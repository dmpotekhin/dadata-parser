# -*- coding: utf-8 -*-
"""
Тесты выбора идентификатора для поиска в DaData: ИНН → ОГРН/ОГРНИП → invalid_id.

Проверяют приоритет ИНН, переход на ОГРН при пустом/невалидном ИНН, статус
invalid_id без расхода квоты, --prefer, --fallback-to-ogrn, ключи кэша
(inn:/ogrn:/ogrnip:), миграцию старых ключей, автопоиск колонок и состав
выходных колонок.

Запуск: python3 -m unittest discover -s tests -v
"""

from __future__ import annotations

import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

try:  # запуск через discover из каталога tests
    from test_cli_integration import MockDaData, read_sheets, run_cli
except ImportError:  # запуск как tests.test_identifier_selection из корня репозитория
    from tests.test_cli_integration import MockDaData, read_sheets, run_cli

# Часовой пояс по умолчанию — как в скрипте (нужен для проверки миграции state).
DEFAULT_TIMEZONE = "Europe/Moscow"


def make_input(path: Path, columns: list[str], rows: list[tuple[str, ...]]) -> Path:
    """Создаёт входной Excel с произвольными колонками."""
    pd.DataFrame(rows, columns=columns).to_excel(path, index=False)
    return path


def cell(sheet: pd.DataFrame, column: str, row: int) -> str:
    """Значение ячейки как строка; пустая ячейка (NaN/None) -> ""."""
    value = sheet[column].iloc[row]
    try:
        if value is None or pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


INN_COLUMNS = ["Компания", "ИНН", "ОГРН"]
INN_ROWS = [
    ("Явно валидный ИНН", "7707083893", "1027700000000"),
    ("Пустой ИНН", "", "1027700000000"),
    ("Невалидный ИНН", "abc", "1027700000000"),
    ("Оба невалидны", "12345", "999"),
    ("Оба пусты", "", ""),
]


class IdentifierSelectionTest(unittest.TestCase):
    """Выбор идентификатора по правилам: ИНН, затем ОГРН/ОГРНИП, затем invalid_id."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def run_case(self, rows, args=(), columns=INN_COLUMNS, name="in.xlsx", mock=None,
                 timeout=240):
        """Запускает CLI на заданных строках и возвращает (result, лист «Все», state, mock)."""
        if mock is None:
            raise AssertionError("mock обязателен: используйте MockDaData()")
        inp = make_input(self.tmp_path / name, columns, rows)
        out = self.tmp_path / "out.xlsx"
        state = self.tmp_path / "state.json"
        cmd = ["--input", str(inp), "--output", str(out), "--state-file", str(state),
               "--max-per-minute", "30", *args]
        result = run_cli(cmd, mock.url, self.tmp_path, timeout=timeout)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        sheet = read_sheets(out)["Все"]
        state_data = json.loads(state.read_text(encoding="utf-8")) if state.exists() else {}
        return result, sheet, state_data, mock.snapshot()

    # ---------- 1. приоритет ИНН ----------

    def test_inn_wins_when_both_identifiers_present(self) -> None:
        with MockDaData() as mock:
            result, sheet, _, stats = self.run_case(
                [("Оба валидны", "7707083893", "1027700000000"), ("Дубль", "7707083893", "1027700000000")],
                mock=mock,
            )
            self.assertEqual(stats["queries"], ["7707083893"], "по умолчанию ищем по ИНН, ОГРН не трогаем")
            self.assertEqual(cell(sheet, "Использованный тип поиска", 0), "inn")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 0), "INN найден и валиден")
            self.assertEqual(cell(sheet, "Ключ запроса", 0), "inn:7707083893")
            self.assertEqual(cell(sheet, "Fallback использован", 0), "нет")
            self.assertEqual(cell(sheet, "Статус", 0), "ok")
            self.assertIn("prefer=inn", result.stdout)

    # ---------- 2. пустой/невалидный ИНН → ОГРН ----------

    def test_empty_and_invalid_inn_fall_through_to_ogrn(self) -> None:
        with MockDaData() as mock:
            _, sheet, _, stats = self.run_case(INN_ROWS, mock=mock)
            self.assertEqual(stats["queries"], ["7707083893", "1027700000000"],
                             "пустой и невалидный ИНН уходят по ОГРН; одинаковый ОГРН в двух "
                             "строках запрашивается один раз, битые строки запросов не тратят")

            self.assertEqual(cell(sheet, "Использованный тип поиска", 1), "ogrn")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 1), "INN пустой, использован ОГРН")
            self.assertEqual(cell(sheet, "Ключ запроса", 1), "ogrn:1027700000000")
            self.assertEqual(cell(sheet, "Нормализованный ОГРН", 1), "1027700000000")
            self.assertEqual(cell(sheet, "Статус", 1), "ok")

            self.assertEqual(cell(sheet, "Использованный тип поиска", 2), "ogrn")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 2), "INN невалидный, использован ОГРН")

    # ---------- 3. оба идентификатора негодны → invalid_id ----------

    def test_invalid_id_rows_do_not_spend_quota(self) -> None:
        with MockDaData() as mock:
            _, sheet, state, stats = self.run_case(INN_ROWS, mock=mock)
            self.assertEqual(cell(sheet, "Статус", 3), "invalid_id")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 3), "Оба идентификатора невалидны")
            self.assertEqual(cell(sheet, "Статус", 4), "invalid_id")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 4), "Оба идентификатора пусты")
            self.assertEqual(cell(sheet, "Ключ запроса", 3), "")
            self.assertEqual(cell(sheet, "Использованный тип поиска", 3), "")
            # 5 строк: 1 по ИНН, 1 уникальный ОГРН (его делят две строки), 2 без запроса
            self.assertEqual(stats["total"], 2)
            self.assertEqual(state["used_today"], 2, "invalid_id не расходует суточный лимит")

    def test_invalid_id_count_in_stats(self) -> None:
        with MockDaData() as mock:
            inp = make_input(self.tmp_path / "in.xlsx", INN_COLUMNS, INN_ROWS)
            out = self.tmp_path / "out.xlsx"
            result = run_cli(
                ["--input", str(inp), "--output", str(out), "--state-file", str(self.tmp_path / "s.json")],
                mock.url, self.tmp_path,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            stats_sheet = read_sheets(out)["Статистика"]
            stats_map = {str(row[0]): row[1] for row in stats_sheet.itertuples(index=False)}
            self.assertIn("Строк без валидного идентификатора (invalid_id)", stats_map)
            self.assertEqual(int(stats_map["Строк без валидного идентификатора (invalid_id)"]), 2)
            self.assertEqual(int(stats_map["Из них поиск по ИНН"]), 1)
            self.assertEqual(int(stats_map["Из них поиск по ОГРН (13 цифр)"]), 2)

    # ---------- 4. ОГРНИП (15 цифр) и заполнение найденных реквизитов ----------

    def test_ogrnip_and_found_requisites(self) -> None:
        with MockDaData() as mock:
            _, sheet, _, stats = self.run_case(
                [
                    ("ОГРНИП ИП", "", "326547600194488"),
                    ("ОГРН юрлица", "", "1027700000000"),
                ],
                mock=mock,
            )
            self.assertEqual(stats["queries"], ["326547600194488", "1027700000000"])
            self.assertEqual(cell(sheet, "Использованный тип поиска", 0), "ogrnip")
            self.assertEqual(cell(sheet, "Ключ запроса", 0), "ogrnip:326547600194488")
            self.assertEqual(cell(sheet, "Найденный ОГРНИП", 0), "326547600194488")
            self.assertEqual(cell(sheet, "Найденный ОГРН", 0), "")
            self.assertEqual(cell(sheet, "Найденный ИНН", 0), "500100732259")
            self.assertEqual(cell(sheet, "Название", 0), "ИП Кузнецов А.А.")

            self.assertEqual(cell(sheet, "Использованный тип поиска", 1), "ogrn")
            self.assertEqual(cell(sheet, "Найденный ОГРН", 1), "1027700000000")
            self.assertEqual(cell(sheet, "Найденный ОГРНИП", 1), "")
            self.assertEqual(cell(sheet, "Найденный КПП", 1), "770701001")

    # ---------- 5. --prefer ogrn ----------

    def test_prefer_ogrn_uses_ogrn_first(self) -> None:
        with MockDaData() as mock:
            _, sheet, _, stats = self.run_case(
                [
                    ("Оба валидны", "7707083893", "1027700000000"),
                    ("ОГРН пустой", "7707083893", ""),
                    ("ОГРН невалидный", "7707083893", "999"),
                ],
                args=["--prefer", "ogrn"], mock=mock,
            )
            self.assertEqual(stats["queries"], ["1027700000000", "7707083893"],
                             "ОГРН первым, затем один запрос по ИНН на два оставшихся ряда")
            self.assertEqual(cell(sheet, "Использованный тип поиска", 0), "ogrn")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 0),
                             "ОГРН найден и валиден (приоритет ОГРН)")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 1), "ОГРН пустой, использован ИНН")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 2), "ОГРН невалидный, использован ИНН")

    # ---------- 6. --fallback-to-ogrn ----------

    def test_not_found_by_inn_stops_without_flag(self) -> None:
        rows = [("ИНН не найден, ОГРН есть", "1234567890", "1027700009999")]
        with MockDaData() as mock:
            _, sheet, state, stats = self.run_case(rows, mock=mock)
            self.assertEqual(stats["queries"], ["1234567890"], "без флага второй запрос не делаем")
            self.assertEqual(cell(sheet, "Статус", 0), "not_found")
            self.assertEqual(cell(sheet, "Fallback использован", 0), "нет")
            self.assertIn("inn:1234567890", state["cache"])

    def test_fallback_to_ogrn_after_not_found(self) -> None:
        rows = [("ИНН не найден, ОГРН есть", "1234567890", "1027700009999")]
        with MockDaData() as mock:
            _, sheet, state, stats = self.run_case(rows, args=["--fallback-to-ogrn"], mock=mock)
            self.assertEqual(stats["queries"], ["1234567890", "1027700009999"])
            self.assertEqual(cell(sheet, "Статус", 0), "ok")
            self.assertEqual(cell(sheet, "Использованный тип поиска", 0), "ogrn")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 0), "INN not_found, fallback на ОГРН")
            self.assertEqual(cell(sheet, "Fallback использован", 0), "да")
            self.assertEqual(cell(sheet, "Ключ запроса", 0), "ogrn:1027700009999")
            self.assertIn("inn:1234567890", state["cache"], "неудачный поиск по ИНН тоже кэшируется")
            self.assertIn("ogrn:1027700009999", state["cache"], "результат фолбэка кэшируется отдельно")

    def test_fallback_after_server_error(self) -> None:
        rows = [("500 по ИНН, ОГРН есть", "9999999999", "1027700000000")]
        with MockDaData() as mock:
            _, sheet, _, stats = self.run_case(
                rows, args=["--fallback-to-ogrn", "--max-retries", "1"], mock=mock,
            )
            # одна попытка (--max-retries 1) → error → второй запрос уже по ОГРН
            self.assertEqual(stats["queries"], ["9999999999", "1027700000000"])
            self.assertEqual(cell(sheet, "Статус", 0), "ok")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 0),
                             "INN error, fallback на ОГРН")
            self.assertEqual(cell(sheet, "Fallback использован", 0), "да")

    def test_fallback_results_are_reused_from_cache(self) -> None:
        rows = [("ИНН не найден, ОГРН есть", "1234567890", "1027700009999")]
        with MockDaData() as mock:
            inp = make_input(self.tmp_path / "in.xlsx", INN_COLUMNS, rows)
            out = self.tmp_path / "out.xlsx"
            state_path = self.tmp_path / "state.json"
            cmd = ["--input", str(inp), "--output", str(out), "--state-file", str(state_path),
                   "--fallback-to-ogrn"]
            first = run_cli(cmd, mock.url, self.tmp_path)
            self.assertEqual(first.returncode, 0, first.stdout + first.stderr)
            self.assertEqual(mock.snapshot()["total"], 2)

            second = run_cli(cmd, mock.url, self.tmp_path)
            self.assertEqual(second.returncode, 0, second.stdout + second.stderr)
            self.assertEqual(mock.snapshot()["total"], 2, "повторный запуск берёт всё из кэша")
            sheet = read_sheets(out)["Все"]
            self.assertEqual(cell(sheet, "Статус", 0), "ok")
            self.assertEqual(cell(sheet, "Fallback использован", 0), "да")

    # ---------- 7. автопоиск колонок ----------

    def test_autodetect_columns_by_partial_name(self) -> None:
        columns = ["Компания", "ИНН компании", "ОГРН_ИП"]
        rows = [("Ромашка", "7707083893", "1027700000000")]
        with MockDaData() as mock:
            result, sheet, _, stats = self.run_case(rows, columns=columns, mock=mock)
            self.assertEqual(stats["queries"], ["7707083893"])
            self.assertEqual(cell(sheet, "Использованный тип поиска", 0), "inn")
            self.assertIn("ИНН компании", result.stdout)
            self.assertIn("ОГРН_ИП", result.stdout)

    def test_explicit_columns_override_autodetect(self) -> None:
        columns = ["Компания", "ИНН", "Мой ОГРН"]
        rows = [("Ромашка", "7707083893", "1027700000000")]
        with MockDaData() as mock:
            inp = make_input(self.tmp_path / "in.xlsx", columns, rows)
            result = run_cli(
                ["--input", str(inp), "--output", str(self.tmp_path / "o.xlsx"),
                 "--state-file", str(self.tmp_path / "s.json"),
                 "--inn-column", "ИНН", "--ogrn-column", "Мой ОГРН"],
                mock.url, self.tmp_path,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertEqual(mock.snapshot()["queries"], ["7707083893"])

    def test_missing_explicit_column_gives_clear_error(self) -> None:
        columns = ["Компания", "ИНН"]
        rows = [("Ромашка", "7707083893")]
        with MockDaData() as mock:
            inp = make_input(self.tmp_path / "in.xlsx", columns, rows)
            result = run_cli(
                ["--input", str(inp), "--output", str(self.tmp_path / "o.xlsx"),
                 "--state-file", str(self.tmp_path / "s.json"), "--ogrn-column", "ОГРН"],
                mock.url, self.tmp_path,
            )
            self.assertEqual(result.returncode, 1)
            self.assertIn("не найдена", result.stdout + result.stderr)
            self.assertEqual(mock.snapshot()["total"], 0)

    def test_spark_identifier_column_is_not_used_for_search(self) -> None:
        """«Идентификатор СПАРК» — служебный номер, в DaData он уходить не должен."""
        columns = ["Идентификатор СПАРК", "Комментарий"]
        rows = [("52794853", "нет ИНН и ОГРН")]
        with MockDaData() as mock:
            _, sheet, _, stats = self.run_case(rows, columns=columns, mock=mock)
            self.assertEqual(stats["total"], 0, "по служебному идентификатору запросов нет")
            self.assertEqual(cell(sheet, "Статус", 0), "invalid_id")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 0), "Оба идентификатора пусты")

    def test_exact_weak_column_is_used_by_value_length(self) -> None:
        columns = ["Компания", "Код"]
        rows = [
            ("Юрлицо по коду", "1027700000000"),
            ("ИП по коду", "326547600194488"),
            ("Мусор в коде", "abc"),
        ]
        with MockDaData() as mock:
            _, sheet, _, stats = self.run_case(rows, columns=columns, mock=mock)
            self.assertEqual(stats["queries"], ["1027700000000", "326547600194488"])
            self.assertEqual(cell(sheet, "Использованный тип поиска", 0), "ogrn")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 0), "Идентификатор распознан как ОГРН")
            self.assertEqual(cell(sheet, "Использованный тип поиска", 1), "ogrnip")
            self.assertEqual(cell(sheet, "Ключ запроса", 1), "ogrnip:326547600194488")
            self.assertEqual(cell(sheet, "Статус", 2), "invalid_id")
            self.assertEqual(cell(sheet, "Причина выбора идентификатора", 2), "Идентификатор не распознан")

    # ---------- 8. ключи кэша и миграция ----------

    def test_legacy_state_keys_are_migrated(self) -> None:
        state_path = self.tmp_path / "state.json"
        today = datetime.now(ZoneInfo(DEFAULT_TIMEZONE)).strftime("%Y-%m-%d")
        state_path.write_text(json.dumps({
            "date": today,
            "used_today": 7,
            "processed_inns": ["7707083893"],
            "last_request_time": 0,
            "cache": {
                "7707083893": {
                    "status": "ok", "name": 'ООО "Ромашка"', "okved_main": "56.10",
                    "okved_extra": ["46.90"], "error": "", "used": 7, "date": "2026-01-01 10:00:00",
                }
            },
        }, ensure_ascii=False), encoding="utf-8")
        rows = [("Из старого кэша", "7707083893", "1027700000000")]
        with MockDaData() as mock:
            inp = make_input(self.tmp_path / "in.xlsx", INN_COLUMNS, rows)
            result = run_cli(
                ["--input", str(inp), "--output", str(self.tmp_path / "o.xlsx"),
                 "--state-file", str(state_path)],
                mock.url, self.tmp_path,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertEqual(mock.snapshot()["total"], 0, "ключ из старого state должен сработать")
            sheet = read_sheets(self.tmp_path / "o.xlsx")["Все"]
            self.assertEqual(cell(sheet, "Статус", 0), "ok")
            self.assertEqual(cell(sheet, "Основной ОКВЭД", 0), "56.10")
            written = json.loads(state_path.read_text(encoding="utf-8"))
            self.assertIn("inn:7707083893", written["cache"])
            self.assertIn("inn:7707083893", written["processed_inns"])
            self.assertEqual(written["used_today"], 7, "суточный счётчик сохраняется")

    def test_cache_keys_carry_identifier_type(self) -> None:
        rows = [
            ("Юрлицо", "7707083893", "1027700000000"),
            ("ИП", "500100732259", ""),
        ]
        with MockDaData() as mock:
            _, _, state, _ = self.run_case(rows, mock=mock)
            self.assertEqual(sorted(state["cache"]), ["inn:500100732259", "inn:7707083893"])
            self.assertEqual(sorted(state["processed_inns"]), ["inn:500100732259", "inn:7707083893"])

    # ---------- 9. состав выходных колонок ----------

    def test_output_columns_match_spec(self) -> None:
        with MockDaData() as mock:
            _, sheet, _, _ = self.run_case([("Ромашка", "7707083893", "1027700000000")], mock=mock)
            expected = [
                "Использованный тип поиска", "Причина выбора идентификатора", "ИНН из строки",
                "ОГРН из строки", "Нормализованный ИНН", "Нормализованный ОГРН", "Найденный ИНН",
                "Найденный ОГРН", "Найденный ОГРНИП", "Найденный КПП", "Название",
                "Основной ОКВЭД", "Дополнительные ОКВЭД", "Статус", "Ошибка",
                "Fallback использован",
            ]
            for column in expected:
                self.assertIn(column, sheet.columns, f"нет колонки {column!r}")
            self.assertEqual(cell(sheet, "ИНН из строки", 0), "7707083893")
            self.assertEqual(cell(sheet, "ОГРН из строки", 0), "1027700000000")

    def test_dry_run_marks_planned_identifier(self) -> None:
        with MockDaData() as mock:
            _, sheet, _, stats = self.run_case(INN_ROWS, args=["--dry-run"], mock=mock)
            self.assertEqual(stats["total"], 0)
            self.assertEqual(cell(sheet, "Статус", 0), "dry_run")
            self.assertEqual(cell(sheet, "Использованный тип поиска", 0), "inn")
            self.assertEqual(cell(sheet, "Ключ запроса", 1), "ogrn:1027700000000")
            self.assertEqual(cell(sheet, "Статус", 3), "invalid_id")


if __name__ == "__main__":
    unittest.main(verbosity=2)