#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dadata_okved_filter.py — выгрузка данных об организациях из DaData по списку ИНН
и фильтрация по префиксам ОКВЭД (по умолчанию "56" и "46").

Что делает скрипт:
 1. Читает Excel-файл (лист --sheet, колонка --column, по умолчанию первый лист
    и колонка "ИНН").
 2. Нормализует ИНН, локально проверяет длину (10 или 12 цифр). Невалидные и
    пустые ИНН запросов к DaData НЕ тратят.
 3. Последовательно (по одному запросу, без потоков и асинхронщины) запрашивает
    POST https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party
 4. Строго соблюдает лимиты DaData:
    - не более 30 запросов в минуту (скользящее окно 60 секунд + минимальный
      интервал 60 / max_per_minute, для 30 rpm это 2 секунды);
    - не более 10 000 запросов в сутки (счётчик с привязкой к дате и часовому поясу).
    Любое фактическое обращение к API, включая повторы при 429/5xx, учитывается и
    в окне, и в суточном счётчике.
 5. Извлекает основной и дополнительные ОКВЭД, считает признаки совпадения
    по префиксам (56/46) и оставляет подходящие организации.
 6. Пишет результат в Excel (листы "Все", "Отфильтровано", "Статистика"),
    ведёт state-файл (checkpoint) и умеет продолжать работу после прерывания.

Примеры запуска:

    # Linux/macOS
    export DADATA_TOKEN="ваш_токен"
    python3 dadata_okved_filter.py --input список_инн.xlsx --output результат.xlsx --column ИНН

    # Windows cmd
    set DADATA_TOKEN=ваш_токен
    python dadata_okved_filter.py --input список_инн.xlsx --output результат.xlsx --column ИНН

    # Windows PowerShell
    $env:DADATA_TOKEN="ваш_токен"
    python dadata_okved_filter.py --input список_инн.xlsx --output результат.xlsx --column ИНН

    # тестовый прогон на 100 строках
    python3 dadata_okved_filter.py --input список_инн.xlsx --output результат_test.xlsx --limit 100

    # план без единого запроса к API
    python3 dadata_okved_filter.py --input список_инн.xlsx --output план.xlsx --dry-run

    # продолжение после дневного лимита / падения
    python3 dadata_okved_filter.py --input список_инн.xlsx --output результат.xlsx --resume

Коды возврата: 0 — успех, 1 — ошибка данных/токена/состояния, 2 — дневной лимит
или квота DaData, 3 — ошибка авторизации, 130 — прерывание (Ctrl+C).
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import random
import re
import sys
import time
from collections import deque
from datetime import datetime, timedelta, tzinfo
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Callable, Optional
from urllib.parse import urlparse

import pandas as pd
import requests

try:  # python-dotenv необязателен
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None  # type: ignore[assignment]

try:  # zoneinfo есть в Python 3.9+, на Windows нужен пакет tzdata
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover
    ZoneInfo = None  # type: ignore[assignment]

LOG = logging.getLogger("dadata")

# --------------------------------------------------------------------------------------
# Константы
# --------------------------------------------------------------------------------------

DEFAULT_API_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"
DEFAULT_MAX_PER_MINUTE = 30
DEFAULT_DAILY_LIMIT = 10_000
DEFAULT_TIMEZONE = "Europe/Moscow"
DEFAULT_TIMEOUT = 10.0
DEFAULT_MAX_RETRIES = 5
DEFAULT_OKVED_PREFIXES = "56,46"
DEFAULT_PROGRESS_EVERY = 20
DEFAULT_SAVE_EVERY = 50

WINDOW_SECONDS = 60.0          # длина скользящего окна для лимита rpm
BACKOFF_BASE_SECONDS = 2.0     # база экспоненциальной задержки между повторами
BACKOFF_CAP_SECONDS = 60.0     # верхняя граница задержки между повторами
RETRY_STATUSES = frozenset({429, 500, 502, 503, 504})

# ИНН российских организаций: 10 цифр (юрлицо) или 12 (ИП/физлицо).
# Именно ASCII-цифры: str.isdigit() пропускает арабско-индийские и надстрочные цифры.
INN_RE = re.compile(r"[0-9]{10}(?:[0-9]{2})?")
# Мусор, который убираем из ИНН: пробелы (в т.ч. NBSP), дефисы (ASCII и типографские)
INN_JUNK_RE = re.compile(r"[\s\u00a0\u2010-\u2015\u2212-]")

# Статусы обработки строки
ST_OK = "ok"
ST_NOT_FOUND = "not_found"
ST_INVALID_INN = "invalid_inn"
ST_EMPTY_INN = "empty_inn"
ST_ERROR = "error"
ST_AUTH_ERROR = "auth_error"
ST_QUOTA_EXCEEDED = "quota_exceeded"
ST_PENDING = "pending"
ST_DRY_RUN = "dry_run"

# Статусы, результат которых можно кэшировать между запусками.
# Ошибки сети/HTTP НЕ кэшируются: при следующем запуске строка запрашивается снова.
CACHEABLE_STATUSES = frozenset({ST_OK, ST_NOT_FOUND})

# Колонки входного и выходного файла
COL_INN = "ИНН"
COL_INN_NORM = "ИНН (нормализованный)"
COL_NAME = "Название"
COL_OKVED_MAIN = "ОКВЭД основной"
COL_OKVED_EXTRA = "ОКВЭД дополнительные"
COL_MATCH_MAIN = "Совпадает основной"
COL_MATCH_ANY = "Совпадает любой"
COL_STATUS = "Статус"
COL_ERROR = "Ошибка"
COL_USED = "Использовано запросов"
COL_DATE = "Дата обработки"

NEW_COLUMNS = [
    COL_INN_NORM, COL_NAME, COL_OKVED_MAIN, COL_OKVED_EXTRA,
    COL_MATCH_MAIN, COL_MATCH_ANY, COL_STATUS, COL_ERROR, COL_USED, COL_DATE,
]

SHEET_ALL = "Все"
SHEET_FILTERED = "Отфильтровано"
SHEET_STATS = "Статистика"

EXIT_OK = 0
EXIT_USAGE = 1
EXIT_LIMIT = 2
EXIT_AUTH = 3
EXIT_INTERRUPT = 130


# --------------------------------------------------------------------------------------
# Исключения
# --------------------------------------------------------------------------------------


class DaDataAuthError(RuntimeError):
    """HTTP 403 без признаков квоты: неверный токен или нет доступа."""


class DaDataQuotaError(RuntimeError):
    """Квота/дневной лимит на стороне DaData (HTTP 402/403 с признаками лимита)."""


class DailyLimitReached(RuntimeError):
    """Локальный дневной лимит запросов исчерпан."""


class StateSaveError(RuntimeError):
    """Не удалось сохранить state-файл — продолжать нельзя, иначе расход не учтётся."""


class OutputWriteError(RuntimeError):
    """Не удалось записать выходной Excel (например, файл открыт в Excel)."""


# --------------------------------------------------------------------------------------
# Вспомогательные функции
# --------------------------------------------------------------------------------------


def configure_logging(verbose: bool) -> None:
    """Настраивает логирование в stdout (без визуального интерфейса)."""
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%H:%M:%S",
        stream=sys.stdout,
        force=True,
    )


def mask_token(token: Optional[str]) -> str:
    """Маскирует токен для логов: abcdefghij1234567890 -> abcd...7890."""
    if not token:
        return "<не задан>"
    if len(token) <= 8:
        return "*" * len(token)
    return f"{token[:4]}...{token[-4:]}"


def human_duration(seconds: float) -> str:
    """Секунды -> человекочитаемая строка."""
    if seconds is None or seconds != seconds or seconds == float("inf"):
        return "—"
    total = int(max(0, round(seconds)))
    hours, remainder = divmod(total, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours} ч {minutes} мин"
    if minutes:
        return f"{minutes} мин {secs} с"
    return f"{secs} с"


def resolve_timezone(name: str) -> tzinfo:
    """
    Возвращает tzinfo по имени. Если zoneinfo/tzdata недоступны, для Europe/Moscow
    используется фиксированное смещение +03:00 (в России переход на летнее время отменён).
    """
    if ZoneInfo is not None:
        try:
            return ZoneInfo(name)
        except Exception as exc:  # нет tzdata или опечатка в имени
            LOG.warning("Не удалось загрузить часовой пояс %s (%s)", name, exc)
    if name in {"Europe/Moscow", "Europe/Minsk"}:
        LOG.warning("Используется фиксированное смещение +03:00 для пояса %s", name)
        from datetime import timezone as _tz
        return _tz(timedelta(hours=3))
    raise SystemExit(f"Неизвестный часовой пояс: {name}. Установите пакет tzdata или укажите другой --timezone.")


def now_local(tz: tzinfo) -> datetime:
    """Текущее время в указанном поясе."""
    return datetime.now(tz)


def iso_now(tz: tzinfo) -> str:
    """Текущее время в ISO-формате с часовым поясом."""
    return now_local(tz).isoformat(timespec="seconds")


def normalize_inn(raw: Any) -> str:
    """
    Приводит значение из Excel к строковому ИНН: убирает пробелы и дефисы,
    обрабатывает числа, float-значения вида 543321094661.0, NaN/None.
    Пустое значение возвращается как "" (запрос к API не выполняется).
    """
    if raw is None:
        return ""
    if isinstance(raw, bool):  # bool — не ИНН
        return ""
    try:
        if pd.isna(raw):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(raw).strip()
    if text.endswith(".0") and text[:-2].lstrip("+-").isdigit():
        text = text[:-2]  # Excel/float: 543321094661.0 -> 543321094661
    return INN_JUNK_RE.sub("", text)


def is_valid_inn(inn: str) -> bool:
    """ИНН валиден, если это ровно 10 или 12 ASCII-цифр."""
    return bool(INN_RE.fullmatch(inn))


def parse_prefixes(raw: str) -> tuple[str, ...]:
    """'56,46' -> ('56', '46')."""
    prefixes = tuple(part.strip() for part in str(raw).split(",") if part.strip())
    if not prefixes:
        raise SystemExit("Не заданы префиксы ОКВЭД (--okved-prefixes).")
    return prefixes


def okved_matches(code: Optional[str], prefixes: tuple[str, ...]) -> bool:
    """True, если код ОКВЭД начинается с одного из префиксов (именно префикс, не подстрока)."""
    if not code:
        return False
    value = str(code).strip()
    return value.startswith(prefixes)


def extract_okveds(raw: Any) -> list[str]:
    """
    Извлекает дополнительные ОКВЭД из data.okveds.

    Поддерживаются оба варианта структуры DaData:
      - список объектов: [{"code": "46.90", "main": false}, ...]
      - список строк:    ["46.90", "62.01"]
      - строка с разделителями: "46.90; 62.01"
    """
    result: list[str] = []
    if not raw:
        return result
    if isinstance(raw, (list, tuple)):
        for item in raw:
            if isinstance(item, dict):
                code = item.get("code") or item.get("okved") or ""
                if str(code).strip():
                    result.append(str(code).strip())
            elif isinstance(item, str):
                if item.strip():
                    result.append(item.strip())
            elif item is not None:
                text = str(item).strip()
                if text:
                    result.append(text)
    elif isinstance(raw, str):
        result = [part.strip() for part in re.split(r"[;,]", raw) if part.strip()]
    return result


def is_loopback_url(url: str) -> bool:
    """True для адресов локальной машины (127.0.0.0/8, localhost, ::1)."""
    host = (urlparse(str(url)).hostname or "").lower()
    if host in {"localhost", "::1", "0.0.0.0"}:
        return True
    return host.startswith("127.")


def validate_api_url(url: str) -> None:
    """
    Разрешает https и http только для loopback (127.0.0.1/localhost/::1).

    Иначе токен Authorization мог бы уйти открытым текстом на произвольный хост.
    """
    parsed = urlparse(str(url))
    scheme = (parsed.scheme or "").lower()
    if scheme == "https":
        return
    if scheme == "http" and is_loopback_url(url):
        return
    raise ValueError(
        f"Небезопасный адрес API: {url!r}. Разрешён только https "
        "(http допускается исключительно для 127.0.0.1/localhost — режим тестов)."
    )


def neutralize_formulas(worksheet: Any) -> int:
    """
    Защита от инъекции формул в Excel: строки, начинающиеся с '=', openpyxl пишет
    как формулы (=HYPERLINK(...), =cmd|...). Переводим такие ячейки в текстовые,
    значение при этом не меняется.
    """
    fixed = 0
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.data_type == "f" and isinstance(cell.value, str):
                cell.data_type = "s"
                fixed += 1
    return fixed


# --------------------------------------------------------------------------------------
# Состояние (checkpoint) и блокировка
# --------------------------------------------------------------------------------------


class StateLock:
    """
    Эксклюзивная блокировка state-файла, чтобы два одновременных запуска не
    израсходовали дневную квоту дважды.

    POSIX — fcntl.flock, Windows — msvcrt.locking. Файл блокировки не удаляется:
    удаление файла под блокировкой открывает классическую гонку.
    """

    def __init__(self, state_path: Path) -> None:
        self.path = Path(str(state_path) + ".lock")
        self._handle: Optional[Any] = None

    def acquire(self) -> bool:
        """Пытается захватить блокировку без ожидания."""
        try:
            self._handle = open(self.path, "a+", encoding="utf-8")
        except OSError as exc:
            LOG.warning("Не удалось открыть lock-файл %s: %s", self.path, exc)
            return True  # не блокируем работу из-за проблем с файлом блокировки
        try:
            if os.name == "nt":
                import msvcrt
                self._handle.seek(0)
                msvcrt.locking(self._handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self._handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self._handle.close()
            self._handle = None
            return False
        try:
            self._handle.seek(0)
            self._handle.truncate()
            self._handle.write(f"pid={os.getpid()}\n")
            self._handle.flush()
        except OSError:
            pass
        return True

    def release(self) -> None:
        """Снимает блокировку."""
        if self._handle is None:
            return
        try:
            if os.name == "nt":
                import msvcrt
                self._handle.seek(0)
                msvcrt.locking(self._handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(self._handle.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass
        finally:
            try:
                self._handle.close()
            except OSError:
                pass
            self._handle = None


class StateStore:
    """
    Файл состояния: дневной счётчик, кэш результатов и список обработанных ИНН.

    Счётчик привязан к дате в выбранном часовом поясе: при смене суток он
    обнуляется, а кэш и список обработанных ИНН сохраняются.
    """

    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self.date: str = ""
        self.used_today: int = 0
        self.processed: list[str] = []
        self.cache: dict[str, dict[str, Any]] = {}
        self.last_request_time: Optional[str] = None
        self.day_reset: bool = False
        self._processed_set: set[str] = set()

    # ---------- загрузка/сохранение ----------

    @classmethod
    def load(cls, path: Path, today: str, resume: bool = True) -> "StateStore":
        """
        Читает состояние. Каждое поле разбирается независимо: битое поле не должно
        стоить нам дневного счётчика.

        resume=False игнорирует кэш и список обработанных ИНН, но НЕ обнуляет дневной
        счётчик: иначе ключ --no-resume позволял бы обойти лимит 10 000 запросов в сутки.
        """
        store = cls(path)
        raw: dict[str, Any] = {}
        if store.path.exists():
            try:
                loaded = json.loads(store.path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    raw = loaded
                else:
                    LOG.warning("State-файл %s не является объектом JSON — начинаем с чистого состояния.", path)
            except (OSError, ValueError) as exc:
                LOG.warning("Не удалось прочитать state-файл %s (%s) — начинаем с чистого состояния.", path, exc)

        date = raw.get("date")
        store.date = str(date) if isinstance(date, (str, int, float)) else ""

        try:
            store.used_today = max(0, int(raw.get("used_today") or 0))
        except (TypeError, ValueError):
            LOG.warning("Поле used_today в state-файле повреждено — считаем счётчик нулевым.")
            store.used_today = 0

        if resume:
            processed = raw.get("processed_inns")
            if isinstance(processed, list):
                store.processed = [str(item) for item in processed]
                store._processed_set = set(store.processed)
            cache = raw.get("cache")
            if isinstance(cache, dict):
                store.cache = {
                    str(inn): record for inn, record in cache.items() if isinstance(record, dict)
                }
        else:
            store.processed = []
            store.cache = {}

        last = raw.get("last_request_time")
        store.last_request_time = str(last) if isinstance(last, str) and last else None

        store.roll_day(today)
        return store

    def roll_day(self, today: str) -> None:
        """Сбрасывает дневной счётчик при смене суток (кэш и processed сохраняются)."""
        if self.date != today:
            had_previous = bool(self.date)
            self.date = today
            self.used_today = 0
            self.day_reset = had_previous or self.day_reset

    def save(self) -> None:
        """
        Сохраняет состояние атомарно (tmp + os.replace). При ошибке пишет исключение:
        продолжать работу, не учитывая расход, опаснее, чем остановиться.
        """
        payload = {
            "date": self.date,
            "used_today": self.used_today,
            "processed_inns": self.processed,
            "last_request_time": self.last_request_time,
            "cache": self.cache,
        }
        last_error: Optional[Exception] = None
        for attempt in (1, 2):
            try:
                self._write_atomic(payload)
                return
            except OSError as exc:
                last_error = exc
                LOG.error("Не удалось записать state-файл %s (попытка %d): %s", self.path, attempt, exc)
                time.sleep(0.5)
        raise StateSaveError(f"state-файл недоступен для записи: {self.path} ({last_error})")

    def _write_atomic(self, payload: dict[str, Any]) -> None:
        tmp = self.path.with_name(self.path.name + ".tmp")
        tmp.parent.mkdir(parents=True, exist_ok=True)
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, self.path)

    # ---------- операции ----------

    def mark_processed(self, inn: str) -> None:
        """Помечает ИНН как обработанный (без дублей)."""
        if inn not in self._processed_set:
            self._processed_set.add(inn)
            self.processed.append(inn)

    def put_cache(self, inn: str, record: dict[str, Any]) -> None:
        """Кладёт кэшируемый результат в состояние."""
        self.cache[inn] = record
        self.mark_processed(inn)

    def bump_requests(self, count: int = 1) -> None:
        """Увеличивает дневной счётчик запросов."""
        self.used_today += count


# --------------------------------------------------------------------------------------
# Ограничитель частоты
# --------------------------------------------------------------------------------------


class DaDataRateLimiter:
    """
    Ограничитель частоты запросов к DaData.

    Соблюдает два условия одновременно:
      1) не более max_per_minute запросов в любые 60 секунд (скользящее окно);
      2) минимальный интервал между запросами = 60 / max_per_minute
         (для 30 rpm это 2 секунды) плюс, при необходимости, дополнительная
         пауза --sleep: реальная пауза = max(sleep, 60 / max_per_minute).

    Дополнительно следит за дневным лимитом: запрос разрешается только если
    остаток на сутки больше нуля. Любой фактический вызов API (включая повторы)
    проходит через acquire(), поэтому учитывается и в скользящем окне, и в
    дневном счётчике — счётчик инкрементируется и сохраняется ДО отправки запроса.
    """

    def __init__(
        self,
        state: StateStore,
        max_per_minute: int = DEFAULT_MAX_PER_MINUTE,
        daily_limit: int = DEFAULT_DAILY_LIMIT,
        extra_sleep: float = 0.0,
        tz: Optional[tzinfo] = None,
        clock: Callable[[], float] = time.monotonic,
        sleeper: Callable[[float], None] = time.sleep,
        wall_clock: Optional[Callable[[], datetime]] = None,
    ) -> None:
        self.state = state
        self.max_per_minute = max(1, int(max_per_minute))
        self.daily_limit = max(1, int(daily_limit))
        self.extra_sleep = max(0.0, float(extra_sleep or 0.0))
        self.min_interval = WINDOW_SECONDS / self.max_per_minute
        self.window: deque[float] = deque()
        self._tz = tz
        self._clock = clock
        self._sleep = sleeper
        self._wall_clock = wall_clock or (lambda: now_local(self._tz) if self._tz is not None else datetime.now())
        self._seed_window_from_state()

    # ---------- свойства ----------

    @property
    def used_today(self) -> int:
        """Сколько запросов уже израсходовано за текущие сутки."""
        return self.state.used_today

    @property
    def budget_left(self) -> int:
        """Сколько запросов ещё можно сделать сегодня."""
        return max(0, self.daily_limit - self.state.used_today)

    @property
    def effective_pause(self) -> float:
        """Фактическая пауза между запросами: max(минимальный интервал, --sleep)."""
        return max(self.min_interval, self.extra_sleep)

    def requests_in_window(self, now: Optional[float] = None) -> int:
        """Число запросов за последние 60 секунд (скользящее окно)."""
        current = self._clock() if now is None else now
        while self.window and self.window[0] <= current - WINDOW_SECONDS:
            self.window.popleft()
        return len(self.window)

    def pacing_span(self) -> tuple[int, float]:
        """
        Сколько запросов в скользящем окне и какова его протяжённость по времени.

        Нужно для честной оценки скорости. Считать от старта процесса нельзя:
        первый запрос уходит сразу, поэтому «n / (сейчас - старт)» завышает
        результат (3 запроса за 4 с дают 45 rpm), хотя фактический интервал
        между ними — 2 с, то есть ровно 30 rpm.
        """
        count = self.requests_in_window()
        if count < 2 or not self.window:
            return count, 0.0
        return count, max(0.0, self.window[-1] - self.window[0])

    def seconds_until_slot(self) -> float:
        """
        Сколько секунд ждать до следующего запроса: учитывает и переполнение окна
        60 секунд, и минимальный интервал между запросами (после resume тоже).
        """
        return max(0.0, self._wait_time(self._clock()))

    # ---------- внутреннее ----------

    def _seed_window_from_state(self) -> None:
        """
        Учитывает последний запрос предыдущего запуска: если он был меньше 60 секунд
        назад, окно не должно считаться пустым, иначе можно превысить rpm после resume.
        """
        last_iso = self.state.last_request_time
        if not last_iso:
            return
        try:
            previous = datetime.fromisoformat(last_iso)
        except ValueError:
            return
        current = self._wall_clock()
        if (previous.tzinfo is None) != (current.tzinfo is None):
            if previous.tzinfo is None:
                previous = previous.replace(tzinfo=current.tzinfo)
            else:
                current = current.replace(tzinfo=previous.tzinfo)
        try:
            elapsed = (current - previous).total_seconds()
        except TypeError:
            return
        if 0 <= elapsed < WINDOW_SECONDS:
            self.window.append(self._clock() - elapsed)

    def _wait_time(self, now: float) -> float:
        """Сколько нужно подождать перед следующим запросом (0 — можно сразу)."""
        self.requests_in_window(now)  # выбрасываем устаревшие отметки
        waits = [0.0]
        if len(self.window) >= self.max_per_minute:
            waits.append(self.window[0] + WINDOW_SECONDS - now + 0.05)
        if self.window:
            last = self.window[-1]
            waits.append(last + self.min_interval - now)
            if self.extra_sleep:
                waits.append(last + self.extra_sleep - now)
        return max(waits)

    def acquire(self, reason: str = "", quiet: bool = False) -> bool:
        """
        Ждёт освобождения слота и регистрирует запрос.

        Возвращает False, если дневной лимит исчерпан (запрос отправлять нельзя).
        Счётчик увеличивается и сохраняется до фактической отправки.
        quiet=True подавляет лог о паузах (используется в тестах).
        """
        if self.budget_left <= 0:
            return False
        while True:
            now = self._clock()
            wait = self._wait_time(now)
            if wait <= 0:
                break
            if not quiet:
                LOG.info(
                    "Лимит %d запросов/мин: пауза %.1f с%s | в окне 60 с: %d/%d",
                    self.max_per_minute, wait, f" ({reason})" if reason else "",
                    len(self.window), self.max_per_minute,
                )
            self._sleep(min(wait, WINDOW_SECONDS))
        self.window.append(self._clock())
        self.state.bump_requests()
        self.state.last_request_time = self._wall_clock().isoformat(timespec="seconds")
        self.state.save()  # счётчик фиксируется на диске до отправки запроса
        return True


# --------------------------------------------------------------------------------------
# Клиент DaData
# --------------------------------------------------------------------------------------


class DaDataClient:
    """Последовательный клиент DaData: один запрос за раз, с повторами и учётом лимитов."""

    def __init__(
        self,
        token: str,
        limiter: DaDataRateLimiter,
        api_url: str = DEFAULT_API_URL,
        timeout: float = DEFAULT_TIMEOUT,
        max_retries: int = DEFAULT_MAX_RETRIES,
        session: Optional[requests.Session] = None,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self.token = token
        self.limiter = limiter
        self.api_url = api_url
        self.timeout = float(timeout)
        self.max_retries = max(1, int(max_retries))
        self.session = session or requests.Session()
        # Локальные (loopback) адреса — обычно мок-сервер в тестах: их нельзя гнать через
        # прокси из окружения (HTTP_PROXY/ALL_PROXY), иначе запрос уходит наружу и падает.
        # Для внешних адресов прокси из окружения, наоборот, уважается.
        if is_loopback_url(api_url):
            self.session.trust_env = False
        self.attempts_total = 0  # фактических обращений к API (включая повторы)
        self._sleep = sleep

    @property
    def headers(self) -> dict[str, str]:
        """Заголовки запроса (токен никогда не пишется в логи)."""
        return {
            "Authorization": f"Token {self.token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    def fetch_party(self, inn: str) -> dict[str, Any]:
        """
        Запрашивает организацию по ИНН.

        Возвращает запись результата. Исключения:
          - DailyLimitReached  — исчерпан локальный дневной лимит;
          - DaDataAuthError    — HTTP 403 без признаков квоты (неверный токен);
          - DaDataQuotaError   — HTTP 402/403 с признаками лимита на стороне DaData.
        Остальные проблемы (сеть, битый JSON, 4xx/5xx) записываются в поле error.
        """
        payload = {"query": inn, "count": 1}
        last_error = ""
        attempt = 0
        while attempt < self.max_retries:
            attempt += 1
            if not self.limiter.acquire(reason=f"попытка {attempt}/{self.max_retries}"):
                raise DailyLimitReached()
            try:
                response = self.session.post(
                    self.api_url, headers=self.headers, json=payload, timeout=self.timeout
                )
                self.attempts_total += 1
            except requests.RequestException as exc:
                self.attempts_total += 1
                last_error = f"network_error: {type(exc).__name__}: {exc}"
                if attempt >= self.max_retries:
                    break
                LOG.warning("Сетевая ошибка (%s). Повтор %d/%d", type(exc).__name__, attempt, self.max_retries)
                self._backoff(attempt)
                continue

            status = response.status_code
            if status == 200:
                return self._parse_success(response, inn)
            if status == 402:
                raise DaDataQuotaError(f"HTTP 402: {(response.text or '')[:300]}")
            if status == 403:
                body = (response.text or "")[:300]
                if self._looks_like_quota(body):
                    raise DaDataQuotaError(f"HTTP 403: {body}")
                raise DaDataAuthError(f"HTTP 403: {body}")
            if status in RETRY_STATUSES:
                last_error = f"http_{status}: {(response.text or '')[:200]}"
                if attempt >= self.max_retries:
                    break
                wait = self._retry_after(response)
                if status == 429 and wait is None:
                    wait = self.limiter.seconds_until_slot()
                LOG.warning(
                    "HTTP %d. Повтор %d/%d через %.1f с", status, attempt, self.max_retries,
                    wait if wait else BACKOFF_BASE_SECONDS,
                )
                self._backoff(attempt, extra=wait)
                continue
            # Прочие ошибки (400, 404, ...) повторять бессмысленно
            return self._error_record(f"http_{status}", (response.text or "")[:300], attempt)
        return self._error_record(ST_ERROR, last_error or "unknown error", attempt)

    # ---------- внутреннее ----------

    @staticmethod
    def _looks_like_quota(body: str) -> bool:
        """По телу ответа 403 понимает, что дело в лимите/квоте, а не в доступе."""
        text = (body or "").lower()
        markers = (
            "limit exceeded", "limits exceeded", "quota", "too many requests",
            "превышен", "лимит", "баланс", "balance", "daily", "сутк",
        )
        return any(marker in text for marker in markers)

    @staticmethod
    def _retry_after(response: requests.Response) -> Optional[float]:
        """Разбирает заголовок Retry-After (секунды или HTTP-дата)."""
        raw = response.headers.get("Retry-After")
        if not raw:
            return None
        raw = raw.strip()
        try:
            return max(0.0, float(raw))
        except ValueError:
            pass
        try:
            target = parsedate_to_datetime(raw)
        except (TypeError, ValueError):
            return None
        if target is None:
            return None
        if target.tzinfo is None:
            target = target.replace(tzinfo=None)
            now = datetime.utcnow()
        else:
            now = datetime.now(target.tzinfo)
        return max(0.0, (target - now).total_seconds())

    def _backoff(self, attempt: int, extra: Optional[float] = None) -> None:
        """Экспоненциальная задержка между повторами (+ jitter)."""
        if extra is not None:
            delay = extra + random.uniform(0, 0.3)
        else:
            delay = min(BACKOFF_CAP_SECONDS, BACKOFF_BASE_SECONDS * (2 ** (attempt - 1)))
            delay += random.uniform(0, 0.5)
        if delay > 0:
            self._sleep(delay)

    def _parse_success(self, response: requests.Response, inn: str) -> dict[str, Any]:
        """Устойчивый разбор успешного ответа DaData."""
        try:
            payload = response.json()
        except ValueError:
            return self._error_record(ST_ERROR, "invalid_json: не удалось разобрать ответ")
        if not isinstance(payload, dict):
            return self._error_record(ST_ERROR, "invalid_json: ответ не является объектом")

        suggestions = payload.get("suggestions") or []
        if not isinstance(suggestions, list) or not suggestions:
            record = self._base_record(ST_NOT_FOUND)
            return record

        first = suggestions[0]
        if not isinstance(first, dict):
            return self._error_record(ST_ERROR, "invalid_structure: suggestions[0] не объект")

        data = first.get("data")
        if not isinstance(data, dict):
            data = {}

        name = first.get("value") or ""
        if not isinstance(name, str) or not name.strip():
            name_block = data.get("name")
            if isinstance(name_block, dict):
                name = name_block.get("short_with_opf") or name_block.get("full_with_opf") or ""
            elif isinstance(name_block, str):
                name = name_block

        main_okved = data.get("okved")
        main_okved = str(main_okved).strip() if main_okved else ""
        extras = extract_okveds(data.get("okveds"))

        record = self._base_record(ST_OK)
        record["name"] = str(name).strip()
        record["okved_main"] = main_okved
        record["okved_extra"] = extras
        return record

    @staticmethod
    def _base_record(status: str) -> dict[str, Any]:
        return {
            "status": status,
            "name": "",
            "okved_main": "",
            "okved_extra": [],
            "error": "",
            "used": None,
            "date": None,
        }

    def _error_record(self, kind: str, message: str, attempts: int = 1) -> dict[str, Any]:
        """Запись для неуспешной строки (сеть/HTTP/битый JSON)."""
        record = self._base_record(ST_ERROR)
        record["error"] = f"{kind} (попыток: {attempts}): {message}"
        return record

    def close(self) -> None:
        """Закрывает HTTP-сессию."""
        try:
            self.session.close()
        except Exception:  # pragma: no cover
            pass


# --------------------------------------------------------------------------------------
# Счётчики для прогресса (инкрементальные, без повторного сканирования всех записей)
# --------------------------------------------------------------------------------------


class MatchCounters:
    """Инкрементальные счётчики найденных организаций, ошибок и совпадений 56/46."""

    def __init__(self, prefixes: tuple[str, ...]) -> None:
        self.prefixes = prefixes
        self.found = 0
        self.not_found = 0
        self.errors = 0
        self.match_main = 0
        self.match_any = 0

    def add(self, record: Optional[dict[str, Any]]) -> None:
        """Учитывает одну запись результата."""
        if not record:
            return
        status = record.get("status")
        if status == ST_OK:
            self.found += 1
        elif status == ST_NOT_FOUND:
            self.not_found += 1
        elif status == ST_ERROR:
            self.errors += 1
        main = record.get("okved_main") or ""
        extras = record.get("okved_extra") or []
        if okved_matches(main, self.prefixes):
            self.match_main += 1
            self.match_any += 1
        elif any(okved_matches(code, self.prefixes) for code in extras):
            self.match_any += 1


# --------------------------------------------------------------------------------------
# Чтение входа, сборка и запись выхода
# --------------------------------------------------------------------------------------


def read_input(path: Path, sheet: Optional[str], column: str) -> tuple[pd.DataFrame, list[str]]:
    """
    Читает Excel и возвращает (датафрейм, список нормализованных ИНН).

    Колонка с ИНН читается как строки (dtype=str): иначе Excel/float могут
    испортить значение (ведущие нули, экспоненциальная запись).
    """
    if not path.exists():
        raise SystemExit(f"Входной файл не найден: {path}")

    names = pd.ExcelFile(path).sheet_names
    sheet_name = sheet if sheet is not None else names[0]
    if sheet_name not in names:
        raise SystemExit(f"Лист {sheet_name!r} не найден. Доступные листы: {', '.join(names)}")

    header = pd.read_excel(path, sheet_name=sheet_name, nrows=0)
    if column not in header.columns:
        raise SystemExit(
            f"Колонка {column!r} не найдена. Доступные колонки: {', '.join(map(str, header.columns))}"
        )

    frame = pd.read_excel(path, sheet_name=sheet_name, dtype={column: str})
    return frame, [normalize_inn(value) for value in frame[column]]


def record_to_row(record: Optional[dict[str, Any]], inn: str, prefixes: tuple[str, ...]) -> dict[str, Any]:
    """Преобразует запись результата в набор выходных колонок."""
    if record is None:
        record = {"status": ST_PENDING}
    status = record.get("status") or ST_PENDING
    main = record.get("okved_main") or ""
    extras = record.get("okved_extra") or []
    match_main = okved_matches(main, prefixes)
    match_any = match_main or any(okved_matches(code, prefixes) for code in extras)
    if status == ST_PENDING:
        flag_main: Optional[bool] = None
        flag_any: Optional[bool] = None
    else:
        flag_main = bool(match_main)
        flag_any = bool(match_any)
    used = record.get("used")
    return {
        COL_INN_NORM: inn,
        COL_NAME: record.get("name") or "",
        COL_OKVED_MAIN: main,
        COL_OKVED_EXTRA: "; ".join(extras),
        COL_MATCH_MAIN: flag_main,
        COL_MATCH_ANY: flag_any,
        COL_STATUS: status,
        COL_ERROR: record.get("error") or "",
        COL_USED: used,
        COL_DATE: record.get("date") or "",
    }


def build_results_df(
    src_df: pd.DataFrame, inns: list[str], records: list[Optional[dict[str, Any]]], prefixes: tuple[str, ...]
) -> pd.DataFrame:
    """Собирает выходной датафрейм: исходные колонки + служебные."""
    rows = [record_to_row(records[idx], inns[idx], prefixes) for idx in range(len(inns))]
    new_columns = pd.DataFrame(rows, columns=NEW_COLUMNS)
    result = pd.concat([src_df.reset_index(drop=True), new_columns], axis=1)
    if COL_INN not in result.columns:
        # В исходном файле нет колонки "ИНН" — добавляем нормализованные значения
        result.insert(0, COL_INN, inns)
    return result


def write_output(path: Path, df_all: pd.DataFrame, df_filtered: pd.DataFrame,
                 stats_rows: list[tuple[str, Any]]) -> None:
    """
    Атомарно пишет выходной Excel: исходные строки, отфильтрованные строки и статистика.
    Формулоподобные значения обезвреживаются (защита от инъекции формул).
    """
    tmp = path.with_name(path.name + ".tmp")
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
            df_all.to_excel(writer, sheet_name=SHEET_ALL, index=False)
            df_filtered.to_excel(writer, sheet_name=SHEET_FILTERED, index=False)
            pd.DataFrame(stats_rows, columns=["Показатель", "Значение"]).to_excel(
                writer, sheet_name=SHEET_STATS, index=False
            )
            for worksheet in writer.book.worksheets:
                neutralize_formulas(worksheet)
        os.replace(tmp, path)
    except OSError as exc:
        raise OutputWriteError(
            f"не удалось записать {path} ({exc}). Если файл открыт в Excel — закройте его. "
            f"Промежуточная копия: {tmp if tmp.exists() else '—'}"
        ) from exc
    except Exception as exc:
        tmp.unlink(missing_ok=True)
        raise OutputWriteError(f"не удалось собрать {path}: {exc}") from exc


def import_cache_from_output(path: Path) -> dict[str, dict[str, Any]]:
    """
    Восстанавливает кэш из ранее созданного выходного файла (лист "Все").

    Нужно, если state-файл потерян: иначе уже сделанные запросы израсходуются повторно.
    """
    if not path.exists():
        return {}
    try:
        frame = pd.read_excel(path, sheet_name=SHEET_ALL, dtype=str)
    except Exception as exc:
        LOG.warning("Не удалось прочитать %s для восстановления кэша: %s", path, exc)
        return {}
    if COL_STATUS not in frame.columns:
        return {}
    inn_column = COL_INN_NORM if COL_INN_NORM in frame.columns else COL_INN
    if inn_column not in frame.columns:
        return {}

    cache: dict[str, dict[str, Any]] = {}
    for _, row in frame.iterrows():
        status = str(row.get(COL_STATUS) or "").strip()
        if status not in CACHEABLE_STATUSES:
            continue
        inn = normalize_inn(row.get(inn_column))
        if not is_valid_inn(inn):
            continue
        extras_raw = row.get(COL_OKVED_EXTRA)
        extras = [part.strip() for part in str(extras_raw or "").split(";") if part.strip()]
        name = row.get(COL_NAME)
        main = row.get(COL_OKVED_MAIN)
        cache[inn] = {
            "status": status,
            "name": "" if pd.isna(name) else str(name),
            "okved_main": "" if pd.isna(main) else str(main),
            "okved_extra": extras,
            "error": "",
            "used": None,
            "date": "" if pd.isna(row.get(COL_DATE)) else str(row.get(COL_DATE)),
            "from_cache": True,
        }
    return cache


# --------------------------------------------------------------------------------------
# Основной сценарий
# --------------------------------------------------------------------------------------


def build_arg_parser() -> argparse.ArgumentParser:
    """Описывает интерфейс командной строки."""
    parser = argparse.ArgumentParser(
        prog="dadata_okved_filter.py",
        description="Выгрузка организаций из DaData по списку ИНН и фильтр по ОКВЭД 56/46.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
        epilog="Пример: python3 dadata_okved_filter.py --input список_инн.xlsx --output результат.xlsx",
    )
    parser.add_argument("--input", required=True, help="Входной Excel-файл со списком ИНН")
    parser.add_argument("--output", default=None,
                        help="Выходной Excel-файл (по умолчанию <входной>_dadata.xlsx рядом с входным)")
    parser.add_argument("--sheet", default=None, help="Имя листа (по умолчанию первый лист)")
    parser.add_argument("--column", default=COL_INN, help="Имя колонки с ИНН")
    parser.add_argument("--token", default=None, help="Токен DaData (иначе берётся из DADATA_TOKEN)")
    parser.add_argument("--only-main", action="store_true",
                        help="Фильтровать только по основному ОКВЭД (иначе — по основному и дополнительным)")
    parser.add_argument("--max-per-minute", type=int, default=DEFAULT_MAX_PER_MINUTE,
                        help="Максимум запросов в минуту (жёсткий лимит DaData — 30)")
    parser.add_argument("--daily-limit", type=int, default=DEFAULT_DAILY_LIMIT,
                        help="Максимум запросов в сутки (жёсткий лимит DaData — 10000)")
    parser.add_argument("--state-file", default=None,
                        help="Файл состояния/checkpoint (по умолчанию dadata_state.json рядом с входным файлом)")
    parser.add_argument("--resume", action=argparse.BooleanOptionalAction, default=True,
                        help="Пропускать уже обработанные ИНН (--resume/--no-resume). "
                             "--no-resume игнорирует кэш, но НЕ обнуляет суточный счётчик")
    parser.add_argument("--timezone", default=DEFAULT_TIMEZONE, help="Часовой пояс для суточного счётчика")
    parser.add_argument("--sleep", type=float, default=0.0,
                        help="Дополнительная пауза между запросами; фактическая пауза = max(sleep, 60/rpm)")
    parser.add_argument("--limit", type=int, default=0, help="Обработать только первые N строк (тест)")
    parser.add_argument("--dry-run", action="store_true", help="Показать план без запросов к API")
    parser.add_argument("--okved-prefixes", default=DEFAULT_OKVED_PREFIXES, help="Префиксы ОКВЭД через запятую")
    parser.add_argument("--progress-every", type=int, default=DEFAULT_PROGRESS_EVERY,
                        help="Печатать прогресс каждые N строк")
    parser.add_argument("--save-every", type=int, default=DEFAULT_SAVE_EVERY,
                        help="Сохранять состояние и выходной файл каждые N строк")
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT, help="Таймаут HTTP-запроса, сек")
    parser.add_argument("--max-retries", type=int, default=DEFAULT_MAX_RETRIES, help="Максимум попыток запроса")
    parser.add_argument("--api-url", default=os.environ.get("DADATA_API_URL", DEFAULT_API_URL),
                        help="Адрес API (по умолчанию endpoint DaData; env DADATA_API_URL)")
    parser.add_argument("--lock", action=argparse.BooleanOptionalAction, default=True,
                        help="Блокировать state-файл от параллельных запусков (--lock/--no-lock)")
    parser.add_argument("--verbose", action="store_true", help="Подробные логи (токен всё равно маскируется)")
    return parser


def resolve_token(args: argparse.Namespace) -> Optional[str]:
    """Берёт токен из --token или DADATA_TOKEN (в логи попадает только маска)."""
    env_token = os.environ.get("DADATA_TOKEN")
    token = (args.token or env_token or "").strip()
    return token or None


def run(args: argparse.Namespace) -> int:
    """Основной сценарий. Возвращает код выхода."""
    configure_logging(args.verbose)
    started = time.monotonic()

    try:
        validate_api_url(args.api_url)
    except ValueError as exc:
        LOG.error("%s", exc)
        return EXIT_USAGE

    token = resolve_token(args)
    if token is None and not args.dry_run:
        LOG.error(
            "Не задан токен DaData. Укажите --token или переменную окружения DADATA_TOKEN "
            "(пример: export DADATA_TOKEN=\"ваш_токен\"). Токен можно взять в личном кабинете DaData."
        )
        return EXIT_USAGE
    if token is None:
        LOG.warning("Токен не задан, но включён --dry-run: реальные запросы выполняться не будут.")

    try:
        tz = resolve_timezone(args.timezone)
    except SystemExit as exc:
        LOG.error("%s", exc)
        return EXIT_USAGE

    prefixes = parse_prefixes(args.okved_prefixes)
    today = now_local(tz).strftime("%Y-%m-%d")

    input_path = Path(args.input).expanduser()
    output_path = (
        Path(args.output).expanduser() if args.output
        else input_path.with_name(input_path.stem + "_dadata.xlsx")
    )
    # Путь состояния по умолчанию привязан к входному файлу: смена рабочего каталога
    # не должна создавать новый счётчик и «обнулять» дневной лимит.
    state_path = (
        Path(args.state_file).expanduser() if args.state_file
        else input_path.parent / "dadata_state.json"
    )

    lock: Optional[StateLock] = None
    if args.lock and not args.dry_run:
        lock = StateLock(state_path)
        if not lock.acquire():
            LOG.error(
                "Этот state-файл уже используется другим процессом (%s). "
                "Дождитесь завершения или укажите другой --state-file.", lock.path,
            )
            return EXIT_USAGE
    try:
        return _execute(args, token, tz, today, prefixes, input_path, output_path, state_path, started)
    finally:
        if lock is not None:
            lock.release()


def _execute(
    args: argparse.Namespace,
    token: Optional[str],
    tz: tzinfo,
    today: str,
    prefixes: tuple[str, ...],
    input_path: Path,
    output_path: Path,
    state_path: Path,
    started: float,
) -> int:
    """Тело основного сценария (после проверок и захвата блокировки)."""
    state = StateStore.load(state_path, today, resume=args.resume)
    if not args.resume and (state.cache or state.processed or state.used_today):
        LOG.warning(
            "--no-resume: кэш и список обработанных ИНН игнорируются, "
            "но суточный счётчик сохранён (уже использовано: %d).", state.used_today,
        )
    if state.day_reset:
        LOG.info("Новый день (%s): суточный счётчик запросов обнулён.", today)

    src_df, inns = read_input(input_path, args.sheet, args.column)
    if args.limit and args.limit > 0:
        src_df = src_df.iloc[: args.limit].reset_index(drop=True)
        inns = inns[: args.limit]
        LOG.info("Ограничение --limit: обрабатываются первые %d строк", args.limit)

    total_rows = len(src_df)
    records: list[Optional[dict[str, Any]]] = [None] * total_rows
    counters = MatchCounters(prefixes)

    valid_indices: list[int] = []
    empty_count = invalid_count = 0
    for idx, inn in enumerate(inns):
        if not inn:
            empty_count += 1
            records[idx] = {"status": ST_EMPTY_INN, "error": "пустой ИНН"}
            continue
        if not is_valid_inn(inn):
            invalid_count += 1
            records[idx] = {"status": ST_INVALID_INN, "error": f"невалидный ИНН: {inn!r}"}
            continue
        if state.cache.get(inn):
            valid_indices.append(idx)
            continue
        valid_indices.append(idx)

    resumed_from_cache = 0
    run_cache: dict[str, dict[str, Any]] = {}
    for idx in valid_indices:
        cached = state.cache.get(inns[idx])
        if cached:
            run_cache[inns[idx]] = cached
            resumed_from_cache += 1

    if not state.cache and args.resume:
        recovered = import_cache_from_output(output_path)
        if recovered:
            state.cache.update(recovered)
            for inn in recovered:
                state.mark_processed(inn)
            LOG.info("Кэш восстановлен из %s: %d ИНН (запросы по ним повторно не пойдут).",
                     output_path, len(recovered))
            resumed_from_cache = 0
            run_cache = {}
            for idx in valid_indices:
                cached = state.cache.get(inns[idx])
                if cached:
                    run_cache[inns[idx]] = cached
                    resumed_from_cache += 1

    to_request = sum(1 for idx in valid_indices if not state.cache.get(inns[idx]))
    unique_to_request = len({inns[idx] for idx in valid_indices if not state.cache.get(inns[idx])})
    budget = max(0, args.daily_limit - state.used_today)

    LOG.info("=" * 90)
    LOG.info("DaData: фильтрация организаций по ОКВЭД %s", "/".join(prefixes))
    LOG.info("Вход: %s | Выход: %s", input_path, output_path)
    LOG.info("Токен: %s | endpoint: %s", mask_token(token), args.api_url)
    LOG.info(
        "Лимиты: %d запросов/мин (интервал %.1f с), %d запросов/сутки, часовой пояс %s",
        args.max_per_minute, WINDOW_SECONDS / max(1, args.max_per_minute), args.daily_limit, args.timezone,
    )
    LOG.info("Режим: resume=%s, only-main=%s, limit=%s, dry-run=%s, state: %s",
             args.resume, args.only_main, args.limit or "все", args.dry_run, state_path)
    LOG.info("=" * 90)
    LOG.info("Строк в файле: %d | валидных: %d (уникальных %d) | невалидных: %d | пустых: %d",
             total_rows, len(valid_indices), len({inns[idx] for idx in valid_indices}),
             invalid_count, empty_count)
    LOG.info("Из кэша (обработано ранее): %d строк | новых ИНН к запросу: %d", resumed_from_cache, unique_to_request)
    LOG.info("Остаток на сегодня: %d запросов из %d (дата %s)", budget, args.daily_limit, today)
    if unique_to_request:
        LOG.info("Примерное время обработки при %d rpm: %s",
                 args.max_per_minute, human_duration(unique_to_request * max(WINDOW_SECONDS / args.max_per_minute, args.sleep)))

    if args.dry_run:
        for idx in valid_indices:
            if records[idx] is None:
                records[idx] = {"status": ST_DRY_RUN, "error": "запросы не отправлялись (--dry-run)",
                                "date": iso_now(tz)}
        LOG.warning("DRY-RUN: запросы к DaData не отправлялись, state-файл не изменялся.")
        LOG.info(
            "План: %d строк, из них %d валидных ИНН требуют запроса (%d уникальных), "
            "%d невалидных и %d пустых — без запросов.",
            total_rows, to_request, unique_to_request, invalid_count, empty_count,
        )
        df_all = build_results_df(src_df, inns, records, prefixes)
        flag = COL_MATCH_MAIN if args.only_main else COL_MATCH_ANY
        df_filtered = df_all[df_all[flag].fillna(False).astype(bool)]
        stats_rows = build_stats_rows(
            state, args, counters, total_rows, len(valid_indices), invalid_count, empty_count,
            resumed_from_cache, 0, len(valid_indices), 0, 0, today, output_path,
        )
        try:
            write_output(output_path, df_all, df_filtered, stats_rows)
        except OutputWriteError as exc:
            LOG.error("%s", exc)
            return EXIT_USAGE
        LOG.info("План сохранён: %s (листы «%s», «%s», «%s»)",
                 output_path, SHEET_ALL, SHEET_FILTERED, SHEET_STATS)
        LOG.info("Длительность: %s", human_duration(time.monotonic() - started))
        return EXIT_OK

    limiter = DaDataRateLimiter(
        state=state,
        max_per_minute=args.max_per_minute,
        daily_limit=args.daily_limit,
        extra_sleep=args.sleep,
        tz=tz,
    )
    client = DaDataClient(
        token=token or "",
        limiter=limiter,
        api_url=args.api_url,
        timeout=args.timeout,
        max_retries=args.max_retries,
    )

    used_at_start = state.used_today
    rows_processed = 0
    limit_reached = False
    interrupted = False
    auth_failed = False
    quota_message = ""
    progress_every = max(1, args.progress_every)
    save_every = max(1, args.save_every)

    def build_filtered() -> pd.DataFrame:
        df_all = build_results_df(src_df, inns, records, prefixes)
        flag = COL_MATCH_MAIN if args.only_main else COL_MATCH_ANY
        mask = df_all[flag].fillna(False).astype(bool)
        return df_all, df_all[mask]

    def checkpoint(force: bool = False) -> None:
        state.save()
        df_all, df_filtered = build_filtered()
        stats_rows = build_stats_rows(
            state, args, counters, total_rows, len(valid_indices), invalid_count, empty_count,
            resumed_from_cache, rows_processed, len(valid_indices),
            state.used_today - used_at_start, client.attempts_total, today, output_path,
        )
        try:
            write_output(output_path, df_all, df_filtered, stats_rows)
        except OutputWriteError as exc:
            LOG.warning("Промежуточная запись результата не удалась: %s", exc)
        else:
            if force:
                LOG.info("Промежуточный результат сохранён: %s", output_path)

    try:
        for position, idx in enumerate(valid_indices, start=1):
            if records[idx] is not None:
                continue  # строка уже заполнена дубликатом ранее в этом запуске
            inn = inns[idx]
            record = run_cache.get(inn)
            if record is not None:
                records[idx] = dict(record)
                counters.add(records[idx])
                rows_processed += 1
                continue

            try:
                record = client.fetch_party(inn)
            except DailyLimitReached:
                limit_reached = True
                break
            except DaDataAuthError as exc:
                LOG.error("Ошибка авторизации DaData (HTTP 403): %s", exc)
                auth_failed = True
                break
            except DaDataQuotaError as exc:
                quota_message = str(exc)
                limit_reached = True
                break

            record["used"] = state.used_today
            record["date"] = iso_now(tz)
            records[idx] = record
            counters.add(record)
            run_cache[inn] = record  # дубликаты ИНН больше запросов не тратят

            if record.get("status") in CACHEABLE_STATUSES:
                state.put_cache(inn, record)

            # Дубликаты того же ИНН ниже по файлу получают тот же результат
            if record.get("status") in CACHEABLE_STATUSES:
                for other in valid_indices[position:]:
                    if inns[other] == inn and records[other] is None:
                        clone = dict(record)
                        clone["from_cache"] = True
                        records[other] = clone
                        counters.add(clone)

            rows_processed += 1

            if rows_processed % progress_every == 0 or position == len(valid_indices):
                _log_progress(
                    rows_processed=rows_processed,
                    total_to_process=len(valid_indices),
                    requests_run=state.used_today - used_at_start,
                    attempts_run=client.attempts_total,
                    used_today=state.used_today,
                    daily_limit=args.daily_limit,
                    started=started,
                    counters=counters,
                    prefixes=prefixes,
                    limiter=limiter,
                )

            if rows_processed % save_every == 0:
                checkpoint()
    except KeyboardInterrupt:
        interrupted = True
        LOG.warning("Прерывание пользователем (Ctrl+C): сохраняю прогресс…")
    except StateSaveError as exc:
        LOG.error("Остановлено: %s", exc)
        LOG.error("Остановка сделана намеренно: без учёта расхода можно превысить дневную квоту DaData.")
        return EXIT_USAGE
    finally:
        try:
            client.close()
        except Exception:  # pragma: no cover
            pass

    # Итоговое сохранение: состояние + выходной файл со актуальными статусами
    try:
        checkpoint(force=True)
    except StateSaveError as exc:
        LOG.error("Не удалось сохранить состояние: %s", exc)
        return EXIT_USAGE
    except OutputWriteError as exc:
        LOG.error("%s", exc)
        return EXIT_USAGE

    df_all, df_filtered = build_filtered()
    requests_run = state.used_today - used_at_start
    _log_summary(
        state=state, args=args, counters=counters, total_rows=total_rows,
        valid_count=len(valid_indices), invalid_count=invalid_count, empty_count=empty_count,
        resumed=resumed_from_cache, rows_processed=rows_processed,
        plan_size=len(valid_indices), requests_run=requests_run,
        attempts_run=client.attempts_total, started=started, output_path=output_path,
        remaining=int(df_filtered.shape[0]),
    )

    if interrupted:
        LOG.warning("Прерывание пользователем. Прогресс сохранён: %s | состояние: %s",
                    output_path.name, state_path)
        return EXIT_INTERRUPT
    if auth_failed:
        LOG.error("Остановлено из-за ошибки авторизации. Прогресс сохранён.")
        return EXIT_AUTH
    if limit_reached:
        if quota_message:
            LOG.error("Ответ DaData: %s", quota_message)
        LOG.error("Достигнут дневной лимит DaData. "
                  "Продолжите обработку позже или увеличьте лимит в личном кабинете.")
        LOG.error("Прогресс сохранён: %s | состояние: %s", output_path.name, state_path)
        return EXIT_LIMIT

    LOG.info("Готово. Результат: %s | состояние: %s", output_path, state_path)
    return EXIT_OK


# --------------------------------------------------------------------------------------
# Логирование прогресса, статистика и запись результатов
# --------------------------------------------------------------------------------------


def _log_progress(
    rows_processed: int,
    total_to_process: int,
    requests_run: int,
    attempts_run: int,
    used_today: int,
    daily_limit: int,
    started: float,
    counters: MatchCounters,
    prefixes: tuple[str, ...],
    limiter: DaDataRateLimiter,
) -> None:
    """Печатает строку прогресса: сколько сделано, сколько запросов ушло, ETA и совпадения."""
    elapsed = max(1e-6, time.monotonic() - started)
    paced, span = limiter.pacing_span()
    rpm = (paced - 1) / span * 60.0 if paced > 1 and span > 0 else 0.0
    speed_text = f"{rpm:.1f} rpm" if rpm > 0 else "н/д (мало запросов)"
    rate = rpm / 60.0 if rpm > 0 else (requests_run / elapsed if requests_run else 0.0)
    remaining = max(0, total_to_process - rows_processed)
    eta = remaining / rate if rate > 0 else float("inf")
    LOG.info(
        "[прогресс] %d/%d строк | запросов в запуске: %d (обращений с повторами: %d) | "
        "за сутки: %d/%d | в окне 60 с: %d/%d | средняя скорость: %s | "
        "найдено: %d | ошибок: %d | совпадений %s: %d | ETA: %s",
        rows_processed, total_to_process, requests_run, attempts_run, used_today, daily_limit,
        limiter.requests_in_window(), limiter.max_per_minute,
        speed_text, counters.found, counters.errors, "/".join(prefixes), counters.match_any,
        human_duration(eta),
    )


def build_stats_rows(
    state: StateStore,
    args: argparse.Namespace,
    counters: MatchCounters,
    total_rows: int,
    valid_count: int,
    invalid_count: int,
    empty_count: int,
    resumed: int,
    rows_processed: int,
    plan_size: int,
    requests_run: int,
    attempts_run: int,
    today: str,
    output_path: Path,
) -> list[tuple[str, Any]]:
    """Собирает строки листа «Статистика»."""
    return [
        ("Дата обработки (часовой пояс)", f"{today} ({args.timezone})"),
        ("Всего строк в файле", total_rows),
        ("Пустых ИНН", empty_count),
        ("Валидных ИНН", valid_count),
        ("Невалидных ИНН (проверено локально, без запросов)", invalid_count),
        ("Уникальных валидных ИНН", plan_size),
        ("Уже обработано ранее (взято из кэша)", resumed),
        ("Обработано строк в этом запуске", rows_processed),
        ("Запросов отправлено в этом запуске", requests_run),
        ("Фактических обращений к API (с повторами)", attempts_run),
        ("Всего использовано за сутки", state.used_today),
        ("Дневной лимит", args.daily_limit),
        ("Осталось запросов на сутки", max(0, args.daily_limit - state.used_today)),
        ("Организаций найдено", counters.found),
        ("Организаций не найдено", counters.not_found),
        ("Ошибок запросов", counters.errors),
        (f"Совпадений по основному ОКВЭД ({','.join(parse_prefixes(args.okved_prefixes))})", counters.match_main),
        (f"Совпадений по основному или дополнительным ({','.join(parse_prefixes(args.okved_prefixes))})",
         counters.match_any),
        ("Фильтр", "только основной ОКВЭД" if args.only_main else "основной или дополнительные"),
        ("Результат", str(output_path)),
    ]


def _log_summary(*, state: StateStore, args: argparse.Namespace, counters: MatchCounters,
                 total_rows: int, valid_count: int, invalid_count: int, empty_count: int,
                 resumed: int, rows_processed: int, plan_size: int, requests_run: int,
                 attempts_run: int, started: float, output_path: Path, remaining: int) -> None:
    """Печатает итоговую статистику запуска."""
    elapsed = time.monotonic() - started
    LOG.info("-" * 90)
    LOG.info("ИТОГИ")
    LOG.info("Всего ИНН в файле: %d | валидных: %d | невалидных: %d | пустых: %d",
             total_rows, valid_count, invalid_count, empty_count)
    LOG.info("Обработано ранее (кэш): %d | обработано в этом запуске: %d | не обработано: %d",
             resumed, rows_processed, max(0, plan_size - rows_processed))
    LOG.info("Запросов отправлено в этом запуске: %d (обращений с повторами: %d) | за сутки: %d/%d",
             requests_run, attempts_run, state.used_today, args.daily_limit)
    LOG.info("Организаций найдено: %d | не найдено: %d | ошибок: %d",
             counters.found, counters.not_found, counters.errors)
    LOG.info("Подходит по основному ОКВЭД (%s): %d | по основному или дополнительным: %d",
             ",".join(parse_prefixes(args.okved_prefixes)), counters.match_main, counters.match_any)
    LOG.info("В листе «%s»: %d строк | длительность: %s | результат: %s",
             SHEET_FILTERED, remaining, human_duration(elapsed), output_path)


def main(argv: Optional[list[str]] = None) -> int:
    """Точка входа CLI."""
    if load_dotenv is not None:
        load_dotenv()  # читает .env рядом со скриптом/вверх по дереву, не переопределяя env
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    if args.max_per_minute < 1 or args.max_per_minute > 1000:
        parser.error("--max-per-minute должен быть в диапазоне 1..1000")
    if args.daily_limit < 1:
        parser.error("--daily-limit должен быть больше нуля")
    if args.limit < 0:
        parser.error("--limit не может быть отрицательным")
    try:
        return run(args)
    except SystemExit as exc:  # ошибки валидации аргументов/входного файла
        message = str(exc.code) if not isinstance(exc.code, int) else ""
        if message:
            LOG.error("%s", message)
            return EXIT_USAGE
        return int(exc.code)


if __name__ == "__main__":
    sys.exit(main())
