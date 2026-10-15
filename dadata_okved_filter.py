#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dadata_okved_filter.py — выгрузка данных об организациях из DaData по списку
ИНН и/или ОГРН и фильтрация по префиксам ОКВЭД (по умолчанию "56" и "46").

Что делает скрипт:
 1. Читает Excel-файл (лист --sheet). Колонки с ИНН и ОГРН/ОГРНИП находятся
    автоматически (--search-by auto), либо задаются явно: --inn-column,
    --ogrn-column (устаревший синоним --column = --inn-column).
 2. Выбирает идентификатор для каждой строки: сначала ИНН (10/12 цифр), затем
    ОГРН (13) / ОГРНИП (15); при --prefer ogrn порядок обратный. Если валидных
    идентификаторов нет, строка получает статус invalid_id и запрос не уходит.
    Локальные проверки лимиты DaData НЕ расходуют.
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
 6. Пишет результат в Excel (листы "Все", "Отфильтровано", "Статистика"): использованный
    тип поиска, причина выбора идентификатора, найденные ИНН/ОГРН/ОГРНИП/КПП, название,
    ОКВЭД, статус, признак использования фолбэка. Ведёт state-файл (checkpoint) с ключами
    inn:/ogrn:/ogrnip: и умеет продолжать работу после прерывания.

Примеры запуска:

    # Linux/macOS
    export DADATA_TOKEN="ваш_токен"
    python3 dadata_okved_filter.py --input список.xlsx --output результат.xlsx --search-by auto --prefer inn

    # Windows cmd
    set DADATA_TOKEN=ваш_токен
    python dadata_okved_filter.py --input список.xlsx --output результат.xlsx --search-by auto --prefer inn

    # Windows PowerShell
    $env:DADATA_TOKEN="ваш_токен"
    python dadata_okved_filter.py --input список.xlsx --output результат.xlsx --search-by auto --prefer inn

    # приоритет ОГРН; фолбэк на ОГРН, если поиск по ИНН не дал результата
    python3 dadata_okved_filter.py --input список.xlsx --output результат.xlsx --prefer ogrn
    python3 dadata_okved_filter.py --input список.xlsx --output результат.xlsx --prefer inn --fallback-to-ogrn

    # явные колонки
    python3 dadata_okved_filter.py --input список.xlsx --output результат.xlsx --inn-column "ИНН компании" --ogrn-column "ОГРН"

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
from dataclasses import dataclass
from datetime import datetime, timedelta, tzinfo
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Callable, Optional, Sequence
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
# Мусор, который убираем из идентификаторов: пробелы (в т.ч. NBSP), дефисы (ASCII и типографские)
INN_JUNK_RE = re.compile(r"[\s\u00a0\u2010-\u2015\u2212-]")

# ОГРН юрлица — 13 цифр, ОГРНИП индивидуального предпринимателя — 15 цифр.
OGRN_RE = re.compile(r"[0-9]{13}")
OGRNIP_RE = re.compile(r"[0-9]{15}")

# Типы идентификаторов: используются в колонке «Использованный тип поиска»
# и как префикс ключа кэша (inn:…, ogrn:…, ogrnip:…).
KIND_INN = "inn"
KIND_OGRN = "ogrn"
KIND_OGRNIP = "ogrnip"
KINDS = (KIND_INN, KIND_OGRN, KIND_OGRNIP)

# Режим поиска и приоритет идентификатора
SEARCH_BY_AUTO = "auto"
PREFER_INN = "inn"
PREFER_OGRN = "ogrn"

# Статусы, после которых при --fallback-to-ogrn имеет смысл пробовать ОГРН
FALLBACK_STATUSES = ("not_found", "error")

# Причины выбора идентификатора (попадают в отчёт как есть)
REASON_INN_OK = "INN найден и валиден"
REASON_INN_EMPTY_OGRN = "INN пустой, использован ОГРН"
REASON_INN_INVALID_OGRN = "INN невалидный, использован ОГРН"
REASON_INN_FALLBACK_OGRN = "INN not_found, fallback на ОГРН"
REASON_INN_ERROR_FALLBACK_OGRN = "INN error, fallback на ОГРН"
REASON_OGRN_OK = "ОГРН найден и валиден (приоритет ОГРН)"
REASON_OGRN_EMPTY_INN = "ОГРН пустой, использован ИНН"
REASON_OGRN_INVALID_INN = "ОГРН невалидный, использован ИНН"
REASON_BOTH_MISSING = "Оба идентификатора пусты"
REASON_BOTH_INVALID = "Оба идентификатора невалидны"
REASON_INN_EMPTY_OGRN_INVALID = "INN пустой, ОГРН невалидный"
REASON_INN_INVALID_OGRN_EMPTY = "INN невалидный, ОГРН пустой"
REASON_ID_KIND_INN = "Идентификатор распознан как ИНН"
REASON_ID_KIND_OGRN = "Идентификатор распознан как ОГРН"
REASON_ID_KIND_OGRNIP = "Идентификатор распознан как ОГРНИП"
REASON_ID_UNRECOGNIZED = "Идентификатор не распознан"

# Статусы обработки строки
ST_OK = "ok"
ST_NOT_FOUND = "not_found"
ST_INVALID_ID = "invalid_id"
ST_ERROR = "error"
ST_AUTH_ERROR = "auth_error"
ST_QUOTA_EXCEEDED = "quota_exceeded"
ST_PENDING = "pending"
ST_DRY_RUN = "dry_run"

# Версия формата state-файла: 2 — ключи кэша с типом идентификатора (inn:/ogrn:/ogrnip:)
STATE_VERSION = 2

# Статусы, результат которых можно кэшировать между запусками.
# Ошибки сети/HTTP НЕ кэшируются: при следующем запуске строка запрашивается снова.
CACHEABLE_STATUSES = frozenset({ST_OK, ST_NOT_FOUND})

# Колонки входного и выходного файла
COL_INN = "ИНН"
COL_OGRN = "ОГРН"
COL_OGRNIP = "ОГРНИП"
COL_INN_RAW = "ИНН из строки"
COL_OGRN_RAW = "ОГРН из строки"
COL_INN_NORM = "Нормализованный ИНН"
COL_OGRN_NORM = "Нормализованный ОГРН"
COL_SEARCH_KIND = "Использованный тип поиска"
COL_REASON = "Причина выбора идентификатора"
COL_FOUND_INN = "Найденный ИНН"
COL_FOUND_OGRN = "Найденный ОГРН"
COL_FOUND_OGRNIP = "Найденный ОГРНИП"
COL_FOUND_KPP = "Найденный КПП"
COL_NAME = "Название"
COL_OKVED_MAIN = "Основной ОКВЭД"
COL_OKVED_EXTRA = "Дополнительные ОКВЭД"
COL_MATCH_MAIN = "Совпадает основной"
COL_MATCH_ANY = "Совпадает любой"
COL_STATUS = "Статус"
COL_ERROR = "Ошибка"
COL_FALLBACK = "Fallback использован"
COL_CACHE_KEY = "Ключ запроса"
COL_USED = "Использовано запросов"
COL_DATE = "Дата обработки"

# Имена колонок из прежних версий отчёта: нужны только для чтения старых файлов.
COL_INN_NORM_LEGACY = "ИНН (нормализованный)"
COL_OKVED_MAIN_LEGACY = "ОКВЭД основной"
COL_OKVED_EXTRA_LEGACY = "ОКВЭД дополнительные"

NEW_COLUMNS = [
    COL_SEARCH_KIND, COL_REASON, COL_INN_RAW, COL_OGRN_RAW, COL_INN_NORM, COL_OGRN_NORM,
    COL_FOUND_INN, COL_FOUND_OGRN, COL_FOUND_OGRNIP, COL_FOUND_KPP,
    COL_NAME, COL_OKVED_MAIN, COL_OKVED_EXTRA, COL_STATUS, COL_ERROR, COL_FALLBACK,
    COL_CACHE_KEY, COL_MATCH_MAIN, COL_MATCH_ANY, COL_USED, COL_DATE,
]

# Автопоиск колонок с идентификаторами (порядок из ТЗ важен)
AUTO_ID_COLUMNS = ("ИНН", "ОГРН", "ОГРНИП", "Идентификатор", "Код")
# Слабые кандидаты: берутся только при точном совпадении имени колонки,
# иначе «Идентификатор СПАРК» ушёл бы в DaData как идентификатор организации.
WEAK_ID_COLUMNS = ("Идентификатор", "Код")

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


def _normalize_digits(raw: Any) -> str:
    """
    Приводит значение из Excel к строке цифр: убирает пробелы и дефисы,
    обрабатывает числа, float-значения вида 543321094661.0, NaN/None.
    Пустое значение возвращается как "" (запрос к API не выполняется).
    """
    if raw is None:
        return ""
    if isinstance(raw, bool):  # bool — не идентификатор
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


def normalize_inn(raw: Any) -> str:
    """Нормализованный ИНН: только цифры, пусто если значения нет."""
    return _normalize_digits(raw)


def normalize_ogrn(raw: Any) -> str:
    """Нормализованный ОГРН/ОГРНИП: только цифры, пусто если значения нет."""
    return _normalize_digits(raw)


def cell_to_text(raw: Any) -> str:
    """Значение ячейки как есть (для колонок «… из строки»), без обрезки цифр."""
    if raw is None:
        return ""
    try:
        if pd.isna(raw):
            return ""
    except (TypeError, ValueError):
        pass
    return str(raw).strip()


def is_valid_inn(inn: str) -> bool:
    """ИНН валиден, если это ровно 10 или 12 ASCII-цифр."""
    return bool(INN_RE.fullmatch(inn))


def is_valid_ogrn(ogrn: str) -> bool:
    """ОГРН юрлица валиден, если это ровно 13 ASCII-цифр."""
    return bool(OGRN_RE.fullmatch(ogrn))


def is_valid_ogrnip(ogrnip: str) -> bool:
    """ОГРНИП валиден, если это ровно 15 ASCII-цифр."""
    return bool(OGRNIP_RE.fullmatch(ogrnip))


def ogrn_kind(value: str) -> str:
    """Тип идентификатора по длине значения: 'ogrn' (13 цифр), 'ogrnip' (15) или ''."""
    if is_valid_ogrn(value):
        return KIND_OGRN
    if is_valid_ogrnip(value):
        return KIND_OGRNIP
    return ""


def cache_key(kind: str, value: str) -> str:
    """Ключ кэша: 'inn:7707083893', 'ogrn:1027700000000', 'ogrnip:304770000012345'."""
    if not kind or not value:
        return ""
    return f"{kind}:{value}"


def migrate_cache_key(key: Any) -> str:
    """
    Ключи прежних версий (голый ИНН без префикса) приводится к 'inn:…'.
    Ключи с префиксом возвращаются как есть — миграция идемпотентна.
    """
    text = str(key).strip()
    if not text or ":" in text:
        return text
    if is_valid_inn(text):
        return cache_key(KIND_INN, text)
    kind = ogrn_kind(text)
    return cache_key(kind, text) if kind else text


def key_value(key: str) -> tuple[str, str]:
    """'inn:7707083893' -> ('inn', '7707083893'); ключ без префикса -> ('', ключ)."""
    kind, sep, value = str(key).partition(":")
    if sep and value:
        return kind, value
    return "", str(key)


def _dedupe_keys(keys: Sequence[Any]) -> list[str]:
    """Список ключей без дублей с сохранением порядка (миграция может их породить)."""
    seen: set[str] = set()
    result: list[str] = []
    for key in keys:
        text = str(key)
        if text and text not in seen:
            seen.add(text)
            result.append(text)
    return result


def split_ogrn_fields(data: dict[str, Any]) -> dict[str, str]:
    """
    Разбирает реквизиты найденной организации: ИНН, КПП, ОГРН и ОГРНИП.

    DaData отдаёт и ОГРН юрлица, и ОГРНИП ИП в одном поле data.ogrn (отдельного
    поля ogrnip в ответе нет — проверено запросом по ОГРНИП 326547600194488),
    поэтому тип определяется по длине значения: 13 цифр — ОГРН, 15 — ОГРНИП.
    Поле data.type (LEGAL/INDIVIDUAL) используется как подсказка.
    """
    inn = str(data.get("inn") or "").strip()
    kpp = str(data.get("kpp") or "").strip()
    ogrn_value = str(data.get("ogrn") or "").strip()
    ogrnip_direct = str(data.get("ogrnip") or "").strip()
    party_type = str(data.get("type") or "").strip().upper()

    found_ogrn = ""
    found_ogrnip = ogrnip_direct
    if ogrn_value:
        if is_valid_ogrnip(ogrn_value) or (party_type == "INDIVIDUAL" and not is_valid_ogrn(ogrn_value)):
            found_ogrnip = found_ogrnip or ogrn_value
        else:
            # 13 цифр или неизвестный формат — не теряем значение, пишем в ОГРН
            found_ogrn = ogrn_value
    return {
        "found_inn": inn,
        "found_ogrn": found_ogrn,
        "found_ogrnip": found_ogrnip,
        "found_kpp": kpp,
    }


@dataclass
class RowIdentifier:
    """Каким идентификатором и почему ищем организацию по конкретной строке."""

    inn_raw: str = ""
    inn_norm: str = ""
    ogrn_raw: str = ""
    ogrn_norm: str = ""
    kind: str = ""
    value: str = ""
    reason: str = ""
    fallback_kind: str = ""
    fallback_value: str = ""

    @property
    def has_valid_inn(self) -> bool:
        return is_valid_inn(self.inn_norm)

    @property
    def has_valid_ogrn(self) -> bool:
        return bool(ogrn_kind(self.ogrn_norm))

    @property
    def key(self) -> str:
        """Ключ кэша по фактически выбранному идентификатору."""
        return cache_key(self.kind, self.value)

    @property
    def fallback_key(self) -> str:
        """Ключ кэша для ОГРН, который можно использовать при fallback."""
        return cache_key(self.fallback_kind, self.fallback_value)


def _missing_reason(inn_norm: str, ogrn_norm: str) -> str:
    """Причина, когда ни ИНН, ни ОГРН не годятся для запроса."""
    if not inn_norm and not ogrn_norm:
        return REASON_BOTH_MISSING
    if not inn_norm:
        return REASON_INN_EMPTY_OGRN_INVALID
    if not ogrn_norm:
        return REASON_INN_INVALID_OGRN_EMPTY
    return REASON_BOTH_INVALID


def choose_identifier(inn_norm: str, ogrn_norm: str, prefer: str = PREFER_INN) -> RowIdentifier:
    """
    Выбирает идентификатор для запроса к DaData.

    prefer=inn  — сначала ИНН; ОГРН берётся, только если ИНН пустой или невалидный.
    prefer=ogrn — сначала ОГРН; ИНН берётся, только если ОГРН пустой или невалидный.

    Флаг --fallback-to-ogrn здесь не учитывается: валидный ОГРН всегда кладётся
    в fallback_*, а решает, использовать ли его, уже основной цикл обработки.
    """
    ident = RowIdentifier(inn_norm=inn_norm, ogrn_norm=ogrn_norm)
    ident.fallback_kind = ogrn_kind(ogrn_norm)
    ident.fallback_value = ogrn_norm if ident.fallback_kind else ""

    inn_ok = is_valid_inn(inn_norm)
    ogrn_ok = ogrn_kind(ogrn_norm)

    if prefer == PREFER_OGRN:
        if ogrn_ok:
            ident.kind, ident.value, ident.reason = ogrn_ok, ogrn_norm, REASON_OGRN_OK
            return ident
        if inn_ok:
            ident.kind, ident.value = KIND_INN, inn_norm
            ident.reason = REASON_OGRN_EMPTY_INN if not ogrn_norm else REASON_OGRN_INVALID_INN
            return ident
        ident.reason = _missing_reason(inn_norm, ogrn_norm)
        return ident

    if inn_ok:
        ident.kind, ident.value, ident.reason = KIND_INN, inn_norm, REASON_INN_OK
        return ident
    if ogrn_ok:
        ident.kind, ident.value = ogrn_ok, ogrn_norm
        ident.reason = REASON_INN_EMPTY_OGRN if not inn_norm else REASON_INN_INVALID_OGRN
        return ident
    ident.reason = _missing_reason(inn_norm, ogrn_norm)
    return ident


def choose_from_weak_column(value: str) -> RowIdentifier:
    """
    Строка из колонки «Идентификатор»/«Код»: тип определяется по длине значения
    (10/12 цифр — ИНН, 13 — ОГРН, 15 — ОГРНИП), иначе идентификатор не распознан.
    """
    ident = RowIdentifier()
    kind = KIND_INN if is_valid_inn(value) else ogrn_kind(value)
    if not kind:
        ident.reason = REASON_ID_UNRECOGNIZED
        return ident
    ident.kind, ident.value = kind, value
    if kind == KIND_INN:
        ident.inn_raw = ident.inn_norm = value
        ident.reason = REASON_ID_KIND_INN
    else:
        ident.ogrn_raw = ident.ogrn_norm = value
        ident.fallback_kind, ident.fallback_value = kind, value
        ident.reason = REASON_ID_KIND_OGRN if kind == KIND_OGRN else REASON_ID_KIND_OGRNIP
    return ident


REQUISITE_FIELD_BY_KIND = {KIND_INN: "found_inn", KIND_OGRN: "found_ogrn", KIND_OGRNIP: "found_ogrnip"}


def enrich_found_from_key(record: dict[str, Any], key: str) -> dict[str, Any]:
    """
    Дополняет запись из старого кэша реквизитами, которые тогда не сохраняли.

    Значение берётся из ключа (inn:/ogrn:/ogrnip:): при успешном поиске DaData
    возвращает ровно тот идентификатор, по которому организация найдена.
    """
    if record.get("status") != ST_OK:
        return record
    kind, value = key_value(key)
    field = REQUISITE_FIELD_BY_KIND.get(kind)
    if field and value and not record.get(field):
        record[field] = value
    return record


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
                # Ключи прежних версий (голый ИНН) мигрируем к 'inn:…'.
                store.processed = _dedupe_keys([migrate_cache_key(item) for item in processed])
                store._processed_set = set(store.processed)
            cache = raw.get("cache")
            if isinstance(cache, dict):
                store.cache = {}
                legacy_keys = False
                for key, record in cache.items():
                    if not isinstance(record, dict):
                        continue
                    migrated_key = migrate_cache_key(key)
                    legacy_keys = legacy_keys or migrated_key != str(key)
                    store.cache[migrated_key] = record
                if legacy_keys:
                    LOG.info(
                        "Ключи кэша состояния перенесены в формат с типом идентификатора "
                        "(inn:/ogrn:/ogrnip:)."
                    )
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
            "version": STATE_VERSION,
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

    def mark_processed(self, key: str) -> None:
        """Помечает ключ идентификатора ('inn:…'/'ogrn:…') обработанным, без дублей."""
        if key not in self._processed_set:
            self._processed_set.add(key)
            self.processed.append(key)

    def put_cache(self, key: str, record: dict[str, Any]) -> None:
        """Кладёт кэшируемый результат в состояние под ключом с типом идентификатора."""
        self.cache[key] = record
        self.mark_processed(key)

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

    def fetch_party(self, query: str, kind: str = KIND_INN) -> dict[str, Any]:
        """
        Запрашивает организацию по идентификатору и помечает запись типом поиска.

        query — ИНН (10 или 12 цифр), ОГРН юрлица (13) или ОГРНИП (15);
        kind  — тип идентификатора (inn/ogrn/ogrnip): попадает в отчёт и в ключ кэша.
        """
        record = self._request_party(query)
        record["search_kind"] = kind
        record["cache_key"] = cache_key(kind, query)
        return record

    def _request_party(self, query: str) -> dict[str, Any]:
        """
        Отправляет запрос к findById/party (повторы, лимиты, разбор ответа).

        Возвращает запись результата. Исключения:
          - DailyLimitReached  — исчерпан локальный дневной лимит;
          - DaDataAuthError    — HTTP 403 без признаков квоты (неверный токен);
          - DaDataQuotaError   — HTTP 402/403 с признаками лимита на стороне DaData.
        Остальные проблемы (сеть, битый JSON, 4xx/5xx) записываются в поле error.
        """
        payload = {"query": query, "count": 1}
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
                return self._parse_success(response)
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

    def _parse_success(self, response: requests.Response) -> dict[str, Any]:
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
        record.update(split_ogrn_fields(data))
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
            # найденные реквизиты организации (по данным DaData, не из строки файла)
            "found_inn": "",
            "found_ogrn": "",
            "found_ogrnip": "",
            "found_kpp": "",
            # служебное: тип фактического поиска и ключ кэша
            "search_kind": "",
            "cache_key": "",
            "fallback": False,
            "reason": "",
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
        self.invalid_id = 0
        self.fallbacks = 0
        self.by_kind: dict[str, int] = {kind: 0 for kind in KINDS}
        self.match_main = 0
        self.match_any = 0

    def add(self, record: Optional[dict[str, Any]]) -> None:
        """Учитывает одну запись результата: статус, тип поиска, фолбэк, совпадения."""
        if not record:
            return
        status = record.get("status")
        if status == ST_OK:
            self.found += 1
        elif status == ST_NOT_FOUND:
            self.not_found += 1
        elif status == ST_INVALID_ID:
            self.invalid_id += 1
        elif status == ST_ERROR:
            self.errors += 1
        if record.get("fallback"):
            self.fallbacks += 1
        kind = record.get("search_kind") or ""
        if kind in self.by_kind:
            self.by_kind[kind] += 1
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


@dataclass
class IdColumns:
    """Колонки, из которых берутся идентификаторы строк."""

    inn: Optional[str] = None
    ogrn: Optional[str] = None
    weak: Optional[str] = None  # «Идентификатор»/«Код»: тип определяется по значению
    source: str = "автопоиск"

    def describe(self) -> str:
        """Описание выбранных колонок для лога и листа «Статистика»."""
        parts = [f"{COL_INN}: {self.inn or '—'}", f"{COL_OGRN}: {self.ogrn or '—'}"]
        if self.weak:
            parts.append(f"общий идентификатор: {self.weak}")
        return f"{'; '.join(parts)} ({self.source})"


def _header_key(name: Any) -> str:
    """Нормализованное имя колонки для сопоставления: «ё» → «е», верхний регистр."""
    return str(name).strip().upper().replace("Ё", "Е")


def cell_to_text(value: Any) -> str:
    """Значение ячейки как оно есть в файле (для колонок «… из строки»)."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def _count_valid(frame: pd.DataFrame, column: str, validator: Callable[[str], Any]) -> int:
    """Сколько значений колонки проходят проверку (для выбора лучшего кандидата)."""
    return sum(1 for value in frame[column] if validator(_normalize_digits(value)))


def _pick_best(
    names: Sequence[str],
    frame: pd.DataFrame,
    predicate: Callable[[str], bool],
    validator: Callable[[str], Any],
) -> Optional[str]:
    """
    Лучший кандидат среди колонок, чьё имя подходит под predicate.

    Лучший = больше валидных значений; при равенстве — самая левая колонка.
    Счёт нужен, чтобы колонка вида «ИНН/КПП» не перебила настоящую колонку ИНН.
    """
    best: Optional[str] = None
    best_score = -1
    for name in names:
        if not predicate(name):
            continue
        score = _count_valid(frame, name, validator)
        if score > best_score:
            best, best_score = name, score
    return best


def detect_id_columns(
    frame: pd.DataFrame, inn_column: Optional[str], ogrn_column: Optional[str]
) -> IdColumns:
    """
    Определяет колонки с ИНН и ОГРН/ОГРНИП.

    Явно указанные колонки (--inn-column/--ogrn-column) берутся как есть, вторая
    при этом доискивается автоматически. Иначе работает автопоиск в порядке
    AUTO_ID_COLUMNS: сначала точное имя, затем «содержит» с выбором по числу
    валидных значений. «Идентификатор»/«Код» берутся только при точном имени —
    иначе в DaData уехал бы, например, «Идентификатор СПАРК».
    """
    names = [str(column) for column in frame.columns]
    available = ", ".join(names) or "(нет колонок)"
    for label, value in (("--inn-column", inn_column), ("--ogrn-column", ogrn_column)):
        if value is not None and value not in names:
            raise SystemExit(
                f"Колонка {value!r} из {label} не найдена. Доступные колонки: {available}"
            )

    columns = IdColumns(
        inn=inn_column,
        ogrn=ogrn_column,
        source="явно" if (inn_column is not None or ogrn_column is not None) else "автопоиск",
    )
    exact = {_header_key(name): name for name in names}

    if columns.inn is None:
        columns.inn = exact.get(_header_key(COL_INN)) or _pick_best(
            names, frame, lambda name: "ИНН" in _header_key(name), is_valid_inn
        )
    if columns.ogrn is None:
        for candidate in ("ОГРНИП", "ОГРН"):
            found = exact.get(_header_key(candidate))
            if found is not None and found != columns.inn:
                columns.ogrn = found
                break
        if columns.ogrn is None:
            columns.ogrn = _pick_best(
                names,
                frame,
                lambda name: name != columns.inn and "ОГРН" in _header_key(name),
                ogrn_kind,
            )
    if columns.inn is None and columns.ogrn is None:
        for candidate in WEAK_ID_COLUMNS:
            found = exact.get(_header_key(candidate))
            if found is not None:
                columns.weak = found
                break
    return columns


def build_identifiers(frame: pd.DataFrame, columns: IdColumns, prefer: str) -> list[RowIdentifier]:
    """Считает идентификатор для каждой строки таблицы (без обращений к API)."""
    inn_values = frame[columns.inn] if columns.inn else None
    ogrn_values = frame[columns.ogrn] if columns.ogrn else None
    weak_values = frame[columns.weak] if columns.weak else None

    identifiers: list[RowIdentifier] = []
    for position in range(len(frame)):
        inn_raw = cell_to_text(inn_values.iloc[position]) if inn_values is not None else ""
        ogrn_raw = cell_to_text(ogrn_values.iloc[position]) if ogrn_values is not None else ""
        if weak_values is not None:
            ident = choose_from_weak_column(_normalize_digits(cell_to_text(weak_values.iloc[position])))
        else:
            ident = choose_identifier(_normalize_digits(inn_raw), _normalize_digits(ogrn_raw), prefer)
            ident.inn_raw = inn_raw
            ident.ogrn_raw = ogrn_raw
        identifiers.append(ident)
    return identifiers


def read_input(
    path: Path,
    sheet: Optional[str],
    inn_column: Optional[str] = None,
    ogrn_column: Optional[str] = None,
    prefer: str = PREFER_INN,
) -> tuple[pd.DataFrame, list[RowIdentifier], IdColumns]:
    """
    Читает Excel и возвращает (датафрейм, идентификаторы строк, выбранные колонки).

    Все колонки читаются как строки (dtype=str): иначе Excel/float могут испортить
    идентификаторы — ведущие нули, экспоненциальная запись, «543321094661.0».
    """
    if not path.exists():
        raise SystemExit(f"Входной файл не найден: {path}")

    names = pd.ExcelFile(path).sheet_names
    sheet_name = sheet if sheet is not None else names[0]
    if sheet_name not in names:
        raise SystemExit(f"Лист {sheet_name!r} не найден. Доступные листы: {', '.join(names)}")

    frame = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
    if not len(frame.columns):
        raise SystemExit(f"Лист {sheet_name!r} пуст — нечего обрабатывать.")

    columns = detect_id_columns(frame, inn_column, ogrn_column)
    identifiers = build_identifiers(frame, columns, prefer)
    return frame, identifiers, columns


def record_to_row(
    record: Optional[dict[str, Any]], ident: RowIdentifier, prefixes: tuple[str, ...]
) -> dict[str, Any]:
    """Преобразует запись результата и идентификатор строки в набор выходных колонок."""
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
    return {
        COL_SEARCH_KIND: record.get("search_kind") or ident.kind,
        COL_REASON: record.get("reason") or ident.reason,
        COL_INN_RAW: ident.inn_raw,
        COL_OGRN_RAW: ident.ogrn_raw,
        COL_INN_NORM: ident.inn_norm,
        COL_OGRN_NORM: ident.ogrn_norm,
        COL_FOUND_INN: record.get("found_inn") or "",
        COL_FOUND_OGRN: record.get("found_ogrn") or "",
        COL_FOUND_OGRNIP: record.get("found_ogrnip") or "",
        COL_FOUND_KPP: record.get("found_kpp") or "",
        COL_NAME: record.get("name") or "",
        COL_OKVED_MAIN: main or None,
        COL_OKVED_EXTRA: "; ".join(extras) or None,
        COL_STATUS: status,
        COL_ERROR: record.get("error") or "",
        COL_FALLBACK: "да" if record.get("fallback") else "нет",
        COL_CACHE_KEY: record.get("cache_key") or ident.key,
        COL_MATCH_MAIN: flag_main,
        COL_MATCH_ANY: flag_any,
        COL_USED: record.get("used"),
        COL_DATE: record.get("date") or "",
    }


def build_results_df(
    src_df: pd.DataFrame,
    identifiers: list[RowIdentifier],
    records: list[Optional[dict[str, Any]]],
    prefixes: tuple[str, ...],
) -> pd.DataFrame:
    """Собирает выходной датафрейм: исходные колонки + служебные."""
    rows = [
        record_to_row(records[idx], identifiers[idx], prefixes)
        for idx in range(len(identifiers))
    ]
    new_columns = pd.DataFrame(rows, columns=NEW_COLUMNS)
    return pd.concat([src_df.reset_index(drop=True), new_columns], axis=1)


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


CACHE_KEY_RE = re.compile(r"^(?:inn|ogrn|ogrnip):[0-9]+$")


def _restore_key(row: pd.Series) -> str:
    """
    Восстанавливает ключ кэша из строки отчёта.

    Новый формат: служебная колонка «Ключ запроса» (inn:/ogrn:/ogrnip:) или пара
    «Использованный тип поиска» + нормализованное значение. Прежний формат отчёта
    (колонка «ИНН (нормализованный)» без типа) читается как inn:…
    """
    raw = row.get(COL_CACHE_KEY) if COL_CACHE_KEY in row.index else None
    text = "" if raw is None or pd.isna(raw) else str(raw).strip()
    if CACHE_KEY_RE.match(text):
        return text

    kind = str(row.get(COL_SEARCH_KIND) or "").strip() if COL_SEARCH_KIND in row.index else ""
    if kind in KINDS:
        column = COL_INN_NORM if kind == KIND_INN else COL_OGRN_NORM
        if column in row.index:
            value = _normalize_digits(row.get(column))
            if kind == KIND_INN and is_valid_inn(value):
                return cache_key(kind, value)
            if kind == KIND_OGRN and is_valid_ogrn(value):
                return cache_key(kind, value)
            if kind == KIND_OGRNIP and is_valid_ogrnip(value):
                return cache_key(kind, value)

    for column in (COL_INN_NORM_LEGACY, COL_INN, COL_INN_NORM):
        if column in row.index:
            value = normalize_inn(row.get(column))
            if is_valid_inn(value):
                return cache_key(KIND_INN, value)
    return ""


def import_cache_from_output(path: Path) -> dict[str, dict[str, Any]]:
    """
    Восстанавливает кэш из ранее созданного выходного файла (лист "Все").

    Нужно, если state-файл потерян: иначе уже сделанные запросы израсходуются повторно.
    Строка, обработанная с фолбэком на ОГРН, даёт ключ по ОГРН — результат по ИНН
    в отчёте не сохраняется, поэтому при повторном запуске он будет запрошен снова.
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

    def text_of(row: pd.Series, column: str) -> str:
        """Значение колонки строкой (пусто, если колонки нет или там NaN)."""
        if column not in frame.columns:
            return ""
        value = row.get(column)
        return "" if value is None or pd.isna(value) else str(value)

    cache: dict[str, dict[str, Any]] = {}
    for _, row in frame.iterrows():
        status = text_of(row, COL_STATUS).strip()
        if status not in CACHEABLE_STATUSES:
            continue
        key = _restore_key(row)
        if not key:
            continue
        extras = [part.strip() for part in text_of(row, COL_OKVED_EXTRA).split(";") if part.strip()]
        main = text_of(row, COL_OKVED_MAIN) or text_of(row, COL_OKVED_MAIN_LEGACY)
        kind = text_of(row, COL_SEARCH_KIND).strip()
        cache[key] = {
            "status": status,
            "name": text_of(row, COL_NAME),
            "okved_main": main,
            "okved_extra": extras,
            "error": "",
            "used": None,
            "date": text_of(row, COL_DATE),
            "from_cache": True,
            "search_kind": kind,
            "reason": text_of(row, COL_REASON).strip(),
            "cache_key": key,
            "fallback": text_of(row, COL_FALLBACK).strip().lower() in ("да", "yes", "true"),
            "found_inn": text_of(row, COL_FOUND_INN),
            "found_ogrn": text_of(row, COL_FOUND_OGRN),
            "found_ogrnip": text_of(row, COL_FOUND_OGRNIP),
            "found_kpp": text_of(row, COL_FOUND_KPP),
        }
    return cache


# --------------------------------------------------------------------------------------
# Основной сценарий
# --------------------------------------------------------------------------------------


def build_arg_parser() -> argparse.ArgumentParser:
    """Описывает интерфейс командной строки."""
    parser = argparse.ArgumentParser(
        prog="dadata_okved_filter.py",
        description=(
            "Выгрузка организаций из DaData по списку ИНН/ОГРН и фильтр по ОКВЭД 56/46. "
            "По умолчанию ищет по ИНН, а ОГРН/ОГРНИП использует только если ИНН пустой, "
            "невалидный или не нашёл организацию (при --fallback-to-ogrn)."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
        epilog=(
            "Примеры: "
            "1) python dadata_okved_filter.py --input список.xlsx --output результат.xlsx "
            "--search-by auto --prefer inn; "
            "2) то же с --prefer ogrn (сначала ОГРН); "
            "3) --prefer inn --fallback-to-ogrn (после not_found по ИНН пробовать ОГРН); "
            "4) --inn-column \"ИНН компании\" --ogrn-column \"ОГРН\" — если колонки названы иначе."
        ),
    )
    parser.add_argument("--input", required=True, help="Входной Excel-файл со списком ИНН и/или ОГРН")
    parser.add_argument("--output", default=None,
                        help="Выходной Excel-файл (по умолчанию <входной>_dadata.xlsx рядом с входным)")
    parser.add_argument("--sheet", default=None, help="Имя листа (по умолчанию первый лист)")
    parser.add_argument("--search-by", choices=(SEARCH_BY_AUTO,), default=SEARCH_BY_AUTO,
                        help="Режим выбора идентификатора: auto — сам находит колонки ИНН/ОГРН/ОГРНИП")
    parser.add_argument("--prefer", choices=(PREFER_INN, PREFER_OGRN), default=PREFER_INN,
                        help="Что искать первым, если в строке есть и ИНН, и ОГРН: inn (ИНН) или ogrn (ОГРН)")
    parser.add_argument("--fallback-to-ogrn", action="store_true",
                        help="Если поиск по ИНН дал not_found/error/timeout, повторить по ОГРН из той же строки")
    parser.add_argument("--inn-column", default=None,
                        help="Колонка с ИНН (иначе определяется автоматически)")
    parser.add_argument("--ogrn-column", default=None,
                        help="Колонка с ОГРН/ОГРНИП (иначе определяется автоматически)")
    parser.add_argument("--column", default=None,
                        help="Устаревший синоним --inn-column (оставлен для совместимости)")
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

    if args.column and not args.inn_column:
        LOG.warning(
            "--column устарел, используйте --inn-column: значение %r принято как колонка ИНН.",
            args.column,
        )
        args.inn_column = args.column
    if args.fallback_to_ogrn and args.prefer == PREFER_OGRN:
        LOG.warning(
            "--fallback-to-ogrn не действует при --prefer ogrn: поиск и так идёт по ОГРН первым."
        )

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
            "--no-resume: кэш и список обработанных идентификаторов игнорируются, "
            "но суточный счётчик сохранён (уже использовано: %d).", state.used_today,
        )
    if state.day_reset:
        LOG.info("Новый день (%s): суточный счётчик запросов обнулён.", today)

    src_df, identifiers, columns = read_input(
        input_path, args.sheet, args.inn_column, args.ogrn_column, args.prefer
    )
    args.inn_column, args.ogrn_column = columns.inn, columns.ogrn
    if args.limit and args.limit > 0:
        src_df = src_df.iloc[: args.limit].reset_index(drop=True)
        identifiers = identifiers[: args.limit]
        LOG.info("Ограничение --limit: обрабатываются первые %d строк", args.limit)

    total_rows = len(src_df)
    records: list[Optional[dict[str, Any]]] = [None] * total_rows
    counters = MatchCounters(prefixes)

    valid_indices: list[int] = []
    invalid_id_count = 0
    for idx, ident in enumerate(identifiers):
        if not ident.kind:
            # Ни ИНН, ни ОГРН не годятся: запрос не отправляем, суточную квоту не тратим.
            invalid_id_count += 1
            records[idx] = {
                "status": ST_INVALID_ID,
                "error": ident.reason or "нет валидного идентификатора",
                "reason": ident.reason,
                "search_kind": "",
                "fallback": False,
            }
            continue
        valid_indices.append(idx)

    to_process = len(valid_indices)
    unique_ids = len({identifiers[idx].key for idx in valid_indices})
    resumed_from_cache = 0
    run_cache: dict[str, dict[str, Any]] = {}
    for idx in valid_indices:
        cached = state.cache.get(identifiers[idx].key)
        if cached:
            run_cache[identifiers[idx].key] = cached
            resumed_from_cache += 1

    if not state.cache and args.resume:
        recovered = import_cache_from_output(output_path)
        if recovered:
            state.cache.update(recovered)
            for key in recovered:
                state.mark_processed(key)
            LOG.info("Кэш восстановлен из %s: %d идентификаторов (запросы по ним повторно не пойдут).",
                     output_path, len(recovered))
            resumed_from_cache = 0
            run_cache = {}
            for idx in valid_indices:
                cached = state.cache.get(identifiers[idx].key)
                if cached:
                    run_cache[identifiers[idx].key] = cached
                    resumed_from_cache += 1

    pending_keys = [
        identifiers[idx].key for idx in valid_indices
        if not state.cache.get(identifiers[idx].key)
    ]
    to_request = len(pending_keys)
    unique_to_request = len(set(pending_keys))
    budget = max(0, args.daily_limit - state.used_today)

    LOG.info("=" * 90)
    LOG.info("DaData: фильтрация организаций по ОКВЭД %s", "/".join(prefixes))
    LOG.info("Вход: %s | Выход: %s", input_path, output_path)
    LOG.info("Токен: %s | endpoint: %s", mask_token(token), args.api_url)
    LOG.info("Колонки идентификаторов: %s", columns.describe())
    for label, column, validator in (
        ("ИНН", columns.inn, is_valid_inn),
        ("ОГРН/ОГРНИП", columns.ogrn, ogrn_kind),
        ("Идентификатор (слабый кандидат)", columns.weak,
         lambda value: bool(is_valid_inn(value) or ogrn_kind(value))),
    ):
        if column:
            LOG.info("  колонка %s: %r — валидных значений %d из %d",
                     label, column, _count_valid(src_df, column, validator), total_rows)
    LOG.info(
        "Лимиты: %d запросов/мин (интервал %.1f с), %d запросов/сутки, часовой пояс %s",
        args.max_per_minute, WINDOW_SECONDS / max(1, args.max_per_minute), args.daily_limit, args.timezone,
    )
    LOG.info("Выбор идентификатора: search-by=%s, prefer=%s, fallback-to-ogrn=%s",
             args.search_by, args.prefer, args.fallback_to_ogrn)
    LOG.info("Режим: resume=%s, only-main=%s, limit=%s, dry-run=%s, state: %s",
             args.resume, args.only_main, args.limit or "все", args.dry_run, state_path)
    LOG.info("=" * 90)
    LOG.info("Строк в файле: %d | с идентификатором: %d (уникальных %d) | без валидного идентификатора: %d",
             total_rows, to_process, unique_ids, invalid_id_count)
    LOG.info("Из кэша (обработано ранее): %d строк | новых идентификаторов к запросу: %d",
             resumed_from_cache, unique_to_request)
    LOG.info("Остаток на сегодня: %d запросов из %d (дата %s)", budget, args.daily_limit, today)
    if unique_to_request:
        LOG.info("Примерное время обработки при %d rpm: %s",
                 args.max_per_minute, human_duration(unique_to_request * max(WINDOW_SECONDS / args.max_per_minute, args.sleep)))

    if args.dry_run:
        for idx in valid_indices:
            if records[idx] is None:
                ident = identifiers[idx]
                records[idx] = {
                    "status": ST_DRY_RUN,
                    "error": "запросы не отправлялись (--dry-run)",
                    "date": iso_now(tz),
                    "search_kind": ident.kind,
                    "reason": ident.reason,
                    "cache_key": ident.key,
                    "fallback": False,
                }
        LOG.warning("DRY-RUN: запросы к DaData не отправлялись, state-файл не изменялся.")
        LOG.info(
            "План: %d строк, из них %d требуют запроса (%d уникальных), "
            "%d без валидного идентификатора — без запросов.",
            total_rows, to_request, unique_to_request, invalid_id_count,
        )
        df_all = build_results_df(src_df, identifiers, records, prefixes)
        flag = COL_MATCH_MAIN if args.only_main else COL_MATCH_ANY
        df_filtered = df_all[df_all[flag].fillna(False).astype(bool)]
        stats_rows = build_stats_rows(
            state, args, counters, total_rows, to_process, invalid_id_count,
            resumed_from_cache, 0, to_process, 0, 0, today, output_path, columns,
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
        df_all = build_results_df(src_df, identifiers, records, prefixes)
        flag = COL_MATCH_MAIN if args.only_main else COL_MATCH_ANY
        mask = df_all[flag].fillna(False).astype(bool)
        return df_all, df_all[mask]

    def checkpoint(force: bool = False) -> None:
        fill_known_duplicates()
        state.save()
        df_all, df_filtered = build_filtered()
        stats_rows = build_stats_rows(
            state, args, counters, total_rows, to_process, invalid_id_count,
            resumed_from_cache, rows_processed, to_process,
            state.used_today - used_at_start, client.attempts_total, today, output_path, columns,
        )
        try:
            write_output(output_path, df_all, df_filtered, stats_rows)
        except OutputWriteError as exc:
            LOG.warning("Промежуточная запись результата не удалась: %s", exc)
        else:
            if force:
                LOG.info("Промежуточный результат сохранён: %s", output_path)

    def cached_result(key: str) -> Optional[dict[str, Any]]:
        """Готовый результат по ключу: из этого запуска или из state-файла (копия)."""
        if not key:
            return None
        cached = run_cache.get(key) or state.cache.get(key)
        return dict(cached) if cached else None

    def request_by(kind: str, value: str, key: str) -> dict[str, Any]:
        """
        Один запрос к DaData по конкретному идентификатору.

        Результат идёт в кэш запуска (дубликаты в файле квоту не тратят) и, если он
        кэшируемый, в state-файл. Сетевые ошибки в state не пишутся: их нужно повторить
        в следующий раз, а не запоминать навсегда.
        """
        record = client.fetch_party(value, kind)
        record["used"] = state.used_today
        record["date"] = iso_now(tz)
        record["cache_key"] = key
        record["search_kind"] = kind
        record["fallback"] = False
        run_cache[key] = dict(record)
        if record.get("status") in CACHEABLE_STATUSES:
            state.put_cache(key, record)
        return record

    def resolve_row(idx: int) -> dict[str, Any]:
        """
        Результат для строки: сначала по основному идентификатору (ИНН или ОГРН),
        затем — если включён --fallback-to-ogrn и основной поиск не удался —
        по ОГРН/ОГРНИП из той же строки.

        Проблемы лимитов и авторизации намеренно не перехватываются: их обрабатывает
        основной цикл, чтобы остановиться и сохранить прогресс.
        """
        ident = identifiers[idx]
        key = ident.key
        record = cached_result(key)
        if record is None:
            record = request_by(ident.kind, ident.value, key)
        else:
            record["search_kind"] = record.get("search_kind") or ident.kind
            record["cache_key"] = record.get("cache_key") or key
            record.setdefault("fallback", False)
            enrich_found_from_key(record, key)

        can_fallback = (
            args.fallback_to_ogrn
            and bool(ident.fallback_kind)
            and ident.fallback_kind != ident.kind
        )
        if not can_fallback or record.get("status") not in FALLBACK_STATUSES:
            if not record.get("reason"):
                record["reason"] = ident.reason
            return record

        primary_status = record.get("status")
        fallback_key = ident.fallback_key
        fallback = cached_result(fallback_key)
        if fallback is None:
            LOG.info("Строка %d: поиск по %s %s дал %s — повторяю по %s %s",
                     idx + 2, ident.kind, ident.value, primary_status,
                     ident.fallback_kind, ident.fallback_value)
            fallback = request_by(ident.fallback_kind, ident.fallback_value, fallback_key)
        fallback = dict(fallback)
        fallback["search_kind"] = ident.fallback_kind
        fallback["cache_key"] = fallback_key
        fallback["fallback"] = True
        fallback["fallback_from"] = primary_status
        fallback["reason"] = (
            REASON_INN_FALLBACK_OGRN if primary_status == ST_NOT_FOUND
            else REASON_INN_ERROR_FALLBACK_OGRN
        )
        return fallback

    def fill_known_duplicates() -> int:
        """Заполняет строки-дубликаты уже известным результатом этого запуска."""
        filled = 0
        for idx in valid_indices:
            if records[idx] is not None:
                continue
            known = run_cache.get(identifiers[idx].key)
            if known is None:
                continue
            clone = dict(known)
            clone["from_cache"] = True
            records[idx] = clone
            counters.add(clone)
            filled += 1
        if filled:
            LOG.info("Дубликаты: %d строк получили результат уже обработанного идентификатора.", filled)
        return filled

    try:
        for position, idx in enumerate(valid_indices, start=1):
            if records[idx] is not None:
                continue
            try:
                record = resolve_row(idx)
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

            records[idx] = record
            counters.add(record)
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

    fill_known_duplicates()
    df_all, df_filtered = build_filtered()
    requests_run = state.used_today - used_at_start
    _log_summary(
        state=state, args=args, counters=counters, total_rows=total_rows,
        valid_count=len(valid_indices), invalid_id_count=invalid_id_count,
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
        "найдено: %d | ошибок: %d | без идентификатора: %d | совпадений %s: %d | ETA: %s",
        rows_processed, total_to_process, requests_run, attempts_run, used_today, daily_limit,
        limiter.requests_in_window(), limiter.max_per_minute,
        speed_text, counters.found, counters.errors, counters.invalid_id,
        "/".join(prefixes), counters.match_any,
        human_duration(eta),
    )


def build_stats_rows(
    state: StateStore,
    args: argparse.Namespace,
    counters: MatchCounters,
    total_rows: int,
    valid_count: int,
    invalid_id_count: int,
    resumed: int,
    rows_processed: int,
    plan_size: int,
    requests_run: int,
    attempts_run: int,
    today: str,
    output_path: Path,
    columns: Optional[IdColumns] = None,
) -> list[tuple[str, Any]]:
    """Собирает строки листа «Статистика»."""
    return [
        ("Дата обработки (часовой пояс)", f"{today} ({args.timezone})"),
        ("Всего строк в файле", total_rows),
        ("Колонки идентификаторов", columns.describe() if columns else "—"),
        ("Выбор идентификатора",
         f"search-by={args.search_by}, prefer={args.prefer}, "
         f"fallback-to-ogrn={'да' if args.fallback_to_ogrn else 'нет'}"),
        ("Строк без валидного идентификатора (invalid_id)", invalid_id_count),
        ("Строк с идентификатором", valid_count),
        ("Уникальных идентификаторов", plan_size),
        ("Из них поиск по ИНН", counters.by_kind.get(KIND_INN, 0)),
        ("Из них поиск по ОГРН (13 цифр)", counters.by_kind.get(KIND_OGRN, 0)),
        ("Из них поиск по ОГРНИП (15 цифр)", counters.by_kind.get(KIND_OGRNIP, 0)),
        ("Запросов-фолбэков на ОГРН (--fallback-to-ogrn)", counters.fallbacks),
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
                 total_rows: int, valid_count: int, invalid_id_count: int,
                 resumed: int, rows_processed: int, plan_size: int, requests_run: int,
                 attempts_run: int, started: float, output_path: Path, remaining: int) -> None:
    """Печатает итоговую статистику запуска."""
    elapsed = time.monotonic() - started
    LOG.info("-" * 90)
    LOG.info("ИТОГИ")
    LOG.info("Всего строк в файле: %d | с идентификатором: %d | без валидного идентификатора: %d",
             total_rows, valid_count, invalid_id_count)
    LOG.info("Поиск: по ИНН %d | по ОГРН %d | по ОГРНИП %d | фолбэков на ОГРН %d",
             counters.by_kind.get(KIND_INN, 0), counters.by_kind.get(KIND_OGRN, 0),
             counters.by_kind.get(KIND_OGRNIP, 0), counters.fallbacks)
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
