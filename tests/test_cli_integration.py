# -*- coding: utf-8 -*-
"""
Интеграционные тесты CLI-скрипта dadata_okved_filter.py против мок-сервера DaData.

Проверяют сквозной сценарий: чтение Excel, нормализацию ИНН, кэш дубликатов,
повторы при 429/500, обработку невалидного JSON и HTTP 400, разбор ОКВЭД,
фильтрацию 56/46, соблюдение интервала запросов, дневной лимит, resume и --dry-run.

Запуск: python3 -m unittest discover -s tests -v
"""

from __future__ import annotations

import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd  # noqa: E402

from mock_dadata import MockDaData  # noqa: E402

REPO = Path(__file__).resolve().parents[1]
SCRIPT = REPO / "dadata_okved_filter.py"
TEST_TOKEN = "test-token-0123456789abcdef"


def make_input(path: Path, rows: list[tuple[str, str]]) -> Path:
    """Создаёт входной Excel: колонка «Компания» + колонка «ИНН»."""
    pd.DataFrame(
        {"Компания": [company for company, _ in rows], "ИНН": [inn for _, inn in rows]}
    ).to_excel(path, index=False)
    return path


def run_cli(args: list[str], api_url: str | None, cwd: Path, timeout: int = 240,
            with_token: bool = True, script: Path | None = None) -> subprocess.CompletedProcess:
    """
    Запускает CLI отдельным процессом.

    `script` нужен для тестов, которым важна изоляция от `.env` проекта: dotenv
    ищет файл от каталога самого скрипта, поэтому копию скрипта кладут во временный
    каталог и запускают её оттуда.
    """
    env = dict(os.environ)
    if with_token:
        env["DADATA_TOKEN"] = TEST_TOKEN
    else:
        env.pop("DADATA_TOKEN", None)
    if api_url:
        env["DADATA_API_URL"] = api_url
    return subprocess.run(
        [sys.executable, str(script or SCRIPT), *args],
        cwd=str(cwd), env=env, capture_output=True, text=True, timeout=timeout,
    )


def read_sheets(path: Path) -> dict[str, pd.DataFrame]:
    return pd.read_excel(path, sheet_name=None, dtype=str)


def statuses_by_inn(sheet: pd.DataFrame) -> dict[str, str]:
    """Словарь ИНН -> статус (пустая ячейка ИНН становится ключом "")."""
    result: dict[str, str] = {}
    for _, row in sheet.iterrows():
        raw_inn = row["ИНН"]
        inn = "" if pd.isna(raw_inn) else str(raw_inn).strip()
        result[inn] = str(row["Статус"]).strip()
    return result


class CliIntegrationTest(unittest.TestCase):
    """Сквозные сценарии основного запуска."""

    ROWS = [
        ("ООО Ромашка", "7707083893"),
        ("ИП Кузнецов", "500100732259"),
        ("ООО Лютик", "7830002293"),
        ("Нет в реестре", "1234567890"),
        ("Ретрай 500", "9999999999"),
        ("Ретрай 429", "8888888888"),
        ("Дубль Ромашки", "7707083893"),
        ("Пробелы и дефисы", " 0123-4567-89 "),
        ("Невалидный", "12345"),
        ("Пустой", ""),
        ("Битый JSON", "6666666666"),
        ("HTTP 400", "5555555555"),
        ("Без доп. ОКВЭД", "4444444444"),
    ]
    # 13 обращений к API: 10 уникальных валидных ИНН + 2 повтора 500 + 1 повтор 429
    EXPECTED_REQUESTS = 13
    EXPECTED_UNIQUE_INNS = {
        "7707083893", "500100732259", "7830002293", "1234567890", "9999999999",
        "8888888888", "0123456789", "6666666666", "5555555555", "4444444444",
    }

    def test_full_run_filter_and_limits(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", self.ROWS)
            out = tmp_path / "out.xlsx"
            state = tmp_path / "state.json"

            result = run_cli(
                ["--input", str(inp), "--output", str(out), "--state-file", str(state),
                 "--max-per-minute", "120", "--progress-every", "5"],
                mock.url, tmp_path,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

            stats = mock.snapshot()
            self.assertEqual(stats["total"], self.EXPECTED_REQUESTS, f"обращений: {stats['inns']}")
            self.assertEqual(set(stats["inns"]), self.EXPECTED_UNIQUE_INNS)
            # Дубликат ИНН не потратил второй запрос
            self.assertEqual(stats["inns"].count("7707083893"), 1)
            # Невалидный ИНН не отправлялся в API
            self.assertNotIn("12345", stats["inns"])
            # Нормализация: пробелы и дефисы убраны перед запросом
            self.assertIn("0123456789", stats["inns"])
            # Повторы при 500 (дважды) и 429 (один раз)
            self.assertEqual(stats["statuses"].count(500), 2)
            self.assertEqual(stats["statuses"].count(429), 1)

            # Соблюдение интервала: при 120 rpm минимальный интервал 0.5 c
            gaps = [b - a for a, b in zip(stats["times"], stats["times"][1:])]
            self.assertTrue(gaps)
            self.assertGreaterEqual(min(gaps), 0.45, f"интервал нарушен: {min(gaps):.3f} c")

            sheets = read_sheets(out)
            self.assertEqual(set(sheets), {"Все", "Отфильтровано", "Статистика"})
            all_sheet = sheets["Все"]
            self.assertEqual(len(all_sheet), len(self.ROWS))

            by_inn = statuses_by_inn(all_sheet)
            self.assertEqual(by_inn["7707083893"], "ok")
            self.assertEqual(by_inn["500100732259"], "ok")
            self.assertEqual(by_inn["1234567890"], "not_found")
            self.assertEqual(by_inn["12345"], "invalid_inn")
            self.assertEqual(by_inn[""], "empty_inn")
            self.assertEqual(by_inn["6666666666"], "error")
            self.assertEqual(by_inn["5555555555"], "error")

            row = all_sheet.iloc[0]
            self.assertIn("Ромашка", row["Название"])
            self.assertEqual(row["ОКВЭД основной"], "56.10")
            self.assertEqual(row["ОКВЭД дополнительные"], "46.90; 62.01")
            self.assertEqual(row["Совпадает основной"], "True")
            self.assertEqual(row["Совпадает любой"], "True")
            # Строка 3: основной 62.01 не подходит, дополнительный 46.90 — подходит
            third = all_sheet.iloc[2]
            self.assertEqual(third["Совпадает основной"], "False")
            self.assertEqual(third["Совпадает любой"], "True")
            # Строка 13: okveds = null -> дополнительные пустые
            self.assertTrue(pd.isna(all_sheet.iloc[12]["ОКВЭД дополнительные"]))
            self.assertEqual(all_sheet.iloc[12]["Совпадает основной"], "True")
            # Ошибки записаны в колонку "Ошибка"
            self.assertIn("invalid_json", str(all_sheet.iloc[10]["Ошибка"]))
            self.assertIn("http_400", str(all_sheet.iloc[11]["Ошибка"]))

            filtered = sheets["Отфильтровано"]
            # 6 строк: Ромашка (56.10) + её дубликат, ИП Кузнецов (46.73),
            # ООО Лютик (доп. 46.90), Пятый (56.20), Тринадцатый (46.01)
            self.assertEqual(len(filtered), 6, "должны остаться 6 строк с 56/46 (одна из них — дубликат)")

            state_data = json.loads(state.read_text(encoding="utf-8"))
            self.assertEqual(state_data["used_today"], self.EXPECTED_REQUESTS)
            # 13 попыток = 10 уникальных ИНН + 3 повтора (500, 500, 429);
            # в кэш попадают только 8 из них: строки с битым JSON и HTTP 400 не кэшируются
            self.assertEqual(len(state_data["processed_inns"]), 8,
                             state_data["processed_inns"])
            self.assertTrue(state_data["last_request_time"])

    def test_second_run_resume_and_only_main(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", self.ROWS)
            out = tmp_path / "out.xlsx"
            state = tmp_path / "state.json"
            base = ["--input", str(inp), "--output", str(out), "--state-file", str(state),
                    "--max-per-minute", "120"]

            first = run_cli(base, mock.url, tmp_path)
            self.assertEqual(first.returncode, 0, first.stdout)

            second = run_cli(base + ["--only-main", "--resume"], mock.url, tmp_path)
            self.assertEqual(second.returncode, 0, second.stdout)
            stats = mock.snapshot()
            # Кэш не даёт повторных запросов; заново запрашиваются только две строки с ошибками
            self.assertEqual(stats["total"], self.EXPECTED_REQUESTS + 2, stats["inns"])
            self.assertEqual(stats["inns"].count("7707083893"), 1)
            # Из 13 строк 11 валидных: 9 берутся из кэша (2 строки с ошибками запрашиваются снова)
            self.assertIn("Из кэша (обработано ранее): 9 строк", second.stdout)

            sheets = read_sheets(out)
            # --only-main: 5 строк с 56/46 в основном ОКВЭД (4 организации + дубликат)
            self.assertEqual(len(sheets["Отфильтровано"]), 5,
                             "--only-main оставляет только 56/46 в основном ОКВЭД")

    def test_resume_without_state_recovers_cache_from_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", self.ROWS)
            out = tmp_path / "out.xlsx"
            state = tmp_path / "state.json"
            base = ["--input", str(inp), "--output", str(out), "--state-file", str(state),
                    "--max-per-minute", "120"]

            self.assertEqual(run_cli(base, mock.url, tmp_path).returncode, 0)
            before = mock.snapshot()["total"]

            # Имитируем потерю state-файла: кэш должен восстановиться из output.xlsx
            state.unlink()
            result = run_cli(base, mock.url, tmp_path)
            self.assertEqual(result.returncode, 0, result.stdout)
            self.assertIn("Кэш восстановлен", result.stdout)
            after = mock.snapshot()["total"]
            # Запрашиваются только ранее неудачные строки (битый JSON и HTTP 400)
            self.assertEqual(after - before, 2, mock.snapshot()["inns"])


class CliLimitAndResumeTest(unittest.TestCase):
    """Дневной лимит: корректная остановка и продолжение."""

    ROWS = CliIntegrationTest.ROWS

    def test_daily_limit_stops_and_resume_continues(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", self.ROWS)
            out = tmp_path / "out.xlsx"
            state = tmp_path / "state.json"
            base = ["--input", str(inp), "--output", str(out), "--state-file", str(state),
                    "--max-per-minute", "600", "--progress-every", "2"]

            first = run_cli(base + ["--daily-limit", "5"], mock.url, tmp_path)
            self.assertEqual(first.returncode, 2, first.stdout)
            self.assertIn("Достигнут дневной лимит DaData.", first.stdout)
            self.assertEqual(mock.snapshot()["total"], 5, mock.snapshot()["inns"])

            state_data = json.loads(state.read_text(encoding="utf-8"))
            self.assertEqual(state_data["used_today"], 5)

            sheets = read_sheets(out)
            all_sheet = sheets["Все"]
            by_inn = statuses_by_inn(all_sheet)
            self.assertEqual(by_inn["7707083893"], "ok")
            self.assertEqual(by_inn["9999999999"], "pending")
            # Недоделанные строки остаются в статусе pending (6 из 13: остальные — ok/not_found/
            # invalid_inn/empty_inn либо дубликат уже обработанного ИНН)
            pending_rows = (all_sheet["Статус"] == "pending").sum()
            self.assertEqual(pending_rows, 6, f"pending строк: {pending_rows}")

            second = run_cli(base, mock.url, tmp_path)
            self.assertEqual(second.returncode, 0, second.stdout)
            stats = mock.snapshot()
            self.assertEqual(stats["total"], CliIntegrationTest.EXPECTED_REQUESTS, stats["inns"])
            self.assertEqual(stats["inns"].count("7707083893"), 1, "кэш должен исключить повторный запрос")

            final = statuses_by_inn(read_sheets(out)["Все"])
            self.assertNotIn("pending", final.values())
            self.assertEqual(final["9999999999"], "ok")


class CliDryRunTest(unittest.TestCase):
    """--dry-run: план без реальных запросов."""

    def test_dry_run_makes_no_requests(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", CliIntegrationTest.ROWS)
            out = tmp_path / "out.xlsx"
            state = tmp_path / "state.json"

            result = run_cli(
                ["--input", str(inp), "--output", str(out), "--state-file", str(state), "--dry-run"],
                mock.url, tmp_path,
            )
            self.assertEqual(result.returncode, 0, result.stdout)
            self.assertEqual(mock.snapshot()["total"], 0, "dry-run не должен обращаться к API")
            self.assertFalse(state.exists(), "dry-run не должен писать state-файл")

            by_inn = statuses_by_inn(read_sheets(out)["Все"])
            self.assertEqual(by_inn["7707083893"], "dry_run")
            self.assertEqual(by_inn["12345"], "invalid_inn")
            self.assertEqual(by_inn[""], "empty_inn")


class CliInterruptTest(unittest.TestCase):
    """Ctrl+C: прогресс сохраняется, затем работа продолжается с --resume."""

    ROWS = [
        ("Первая", "1111111111"),
        ("Вторая", "2222222222"),
        ("Третья", "1010101010"),
        ("Четвёртая", "2020202020"),
        ("Пятая", "3030303030"),
        ("Шестая", "4040404040"),
    ]

    def test_sigint_saves_progress_then_resume(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", self.ROWS)
            out = tmp_path / "out.xlsx"
            state = tmp_path / "state.json"

            env = dict(os.environ)
            env["DADATA_TOKEN"] = TEST_TOKEN
            env["DADATA_API_URL"] = mock.url
            process = subprocess.Popen(
                [sys.executable, str(SCRIPT), "--input", str(inp), "--output", str(out),
                 "--state-file", str(state), "--max-per-minute", "30"],
                cwd=str(tmp_path), env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
            )
            time.sleep(5.5)
            process.send_signal(signal.SIGINT)
            stdout, _ = process.communicate(timeout=60)

            self.assertEqual(process.returncode, 130, stdout)
            self.assertIn("Прерывание пользователем", stdout)
            self.assertTrue(state.exists(), "state-файл должен быть сохранён при Ctrl+C")
            self.assertTrue(out.exists(), "промежуточный результат должен быть сохранён")
            state_data = json.loads(state.read_text(encoding="utf-8"))
            self.assertGreaterEqual(state_data["used_today"], 1)
            self.assertLess(state_data["used_today"], len(self.ROWS))
            partial = statuses_by_inn(read_sheets(out)["Все"])
            self.assertIn("pending", partial.values())

            before = mock.snapshot()["total"]
            resume = run_cli(
                ["--input", str(inp), "--output", str(out), "--state-file", str(state),
                 "--max-per-minute", "600", "--resume"],
                mock.url, tmp_path,
            )
            self.assertEqual(resume.returncode, 0, resume.stdout)
            final = statuses_by_inn(read_sheets(out)["Все"])
            self.assertNotIn("pending", final.values())
            self.assertEqual(set(final.values()), {"ok"})

            stats = mock.snapshot()
            self.assertEqual(len(set(stats["inns"])), len(self.ROWS))
            # Обработанные до Ctrl+C ИНН не запрашиваются повторно
            already = stats["inns"][:before]
            for inn in set(already):
                self.assertLessEqual(stats["inns"].count(inn), 1, f"{inn} запрошен повторно")


class CliErrorHandlingTest(unittest.TestCase):
    """Токен, маскировка и сетевые ошибки."""

    def test_missing_token_gives_clear_error(self) -> None:
        """
        Рядом со скриптом нет .env, а в окружении нет DADATA_TOKEN — ждём понятную
        ошибку. Скрипт копируется в tmp: у реального проекта рядом лежит работающий
        .env, который dotenv нашёл бы и тест перестал бы проверять свою сцену.
        """
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            local_script = tmp_path / SCRIPT.name
            shutil.copy(SCRIPT, local_script)
            inp = make_input(tmp_path / "in.xlsx", [("Тест", "7707083893")])
            result = run_cli(["--input", str(inp), "--output", str(tmp_path / "o.xlsx")],
                             None, tmp_path, with_token=False, script=local_script)
            self.assertEqual(result.returncode, 1)
            self.assertIn("Не задан токен DaData", result.stdout + result.stderr)
            self.assertIn("DADATA_TOKEN", result.stdout + result.stderr)

    def test_token_is_masked_in_logs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("Тест", "1111111111")])
            result = run_cli(
                ["--input", str(inp), "--output", str(tmp_path / "o.xlsx"),
                 "--token", "abcdefghij1234567890", "--max-per-minute", "120"],
                mock.url, tmp_path, with_token=False,
            )
            self.assertEqual(result.returncode, 0, result.stdout)
            self.assertNotIn("abcdefghij1234567890", result.stdout)
            self.assertIn("abcd...7890", result.stdout)

    def test_network_error_is_recorded_per_row(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("Тест", "1111111111")])
            out = tmp_path / "o.xlsx"
            result = run_cli(
                ["--input", str(inp), "--output", str(out), "--max-retries", "2",
                 "--api-url", "http://127.0.0.1:1/suggestions/api/4_1/rs/findById/party"],
                None, tmp_path,
            )
            self.assertEqual(result.returncode, 0, result.stdout)
            sheet = read_sheets(out)["Все"]
            self.assertEqual(sheet.iloc[0]["Статус"], "error")
            self.assertIn("network_error", str(sheet.iloc[0]["Ошибка"]))

    def test_quota_403_stops_gracefully(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("Квота", "7777777777"), ("Вторая", "1111111111")])
            out = tmp_path / "o.xlsx"
            state = tmp_path / "state.json"
            result = run_cli(
                ["--input", str(inp), "--output", str(out), "--state-file", str(state),
                 "--max-per-minute", "600"],
                mock.url, tmp_path,
            )
            self.assertEqual(result.returncode, 2, result.stdout)
            self.assertIn("Достигнут дневной лимит DaData.", result.stdout)
            self.assertTrue(out.exists())
            self.assertTrue(state.exists())

    def test_auth_403_stops_without_retries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("Нет доступа", "3333333333"), ("Вторая", "1111111111")])
            out = tmp_path / "o.xlsx"
            result = run_cli(
                ["--input", str(inp), "--output", str(out), "--max-per-minute", "600",
                 "--max-retries", "5"],
                mock.url, tmp_path,
            )
            self.assertEqual(result.returncode, 3, result.stdout)
            self.assertEqual(mock.snapshot()["total"], 1, "403 не должен повторяться")
            self.assertIn("Ошибка авторизации", result.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
