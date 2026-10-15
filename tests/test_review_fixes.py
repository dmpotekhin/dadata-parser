# -*- coding: utf-8 -*-
"""
Регрессионные тесты на дефекты, найденные при независимом ревью.

Каждый тест закрывает конкретную дыру:
  1. --no-resume не обнуляет суточный счётчик (иначе лимит 10 000/сутки обходится).
  2. Параллельный запуск с тем же state-файлом блокируется, а не тратит двойную квоту.
  3. Дубликаты строк с ошибочным ИНН не жгут по запросу на каждую копию.
  4. Строки, начинающиеся с "=", не превращаются в формулы Excel (инъекция).
  5. ИНН читается как строка, и в выходе есть нормализованный ИНН.
  6. Не-ASCII цифры не считаются ИНН (str.isdigit() пропускал арабско-индийские цифры).
  7. http на внешний хост запрещён (иначе токен уходил бы открытым текстом).
  8. Сетевые ошибки попадают в счётчик обращений с повторами.
  9. Смена суток сбрасывает счётчик, но сохраняет кэш.
 10. Токен читается из .env рядом со скриптом; переменная окружения имеет приоритет,
     а сам токен не попадает в state-файл.
 11. Заявленная в логе скорость не превышает настроенный лимит rpm.

Запуск: python3 -m unittest discover -s tests -v
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402

from dadata_okved_filter import EXIT_LIMIT, StateStore  # noqa: E402
from mock_dadata import MockDaData  # noqa: E402
from test_cli_integration import make_input, read_sheets, run_cli  # noqa: E402

FORMULA = '=HYPERLINK("http://evil.example","click")'
ARABIC_INN = "\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669\u0660"  # ١٢٣٤٥٦٧٨٩٠


def make_input_with_text_cell(path: Path, value: str, inn: str) -> Path:
    """
    Входной Excel, где значение записано именно как ТЕКСТ.

    Так выглядит реальный файл: pandas при чтении берёт значение ячейки (для формул —
    кэшированный результат), поэтому инъекция приходит из текстовых ячеек.
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Компания", "ИНН"])
    sheet.append([value, inn])
    sheet.cell(row=2, column=1).data_type = "s"  # принудительно текст, не формула
    workbook.save(path)
    return path


class ReviewFixesTest(unittest.TestCase):
    """Регрессии по находкам ревью."""

    def _base_args(self, tmp_path: Path, inp: Path, extra: list[str] | None = None) -> list[str]:
        args = [
            "--input", str(inp), "--output", str(tmp_path / "o.xlsx"),
            "--state-file", str(tmp_path / "state.json"),
            "--max-per-minute", "120", "--sleep", "0", "--progress-every", "1",
        ]
        return args + (extra or [])

    # ---------- 1. суточный счётчик ----------

    def test_no_resume_keeps_daily_counter(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            state = tmp_path / "state.json"
            inp = make_input(tmp_path / "in.xlsx", [("A", "7707083893"), ("B", "4444444444")])
            base = self._base_args(tmp_path, inp)

            first = run_cli(base + ["--daily-limit", "1"], mock.url, tmp_path)
            self.assertEqual(first.returncode, EXIT_LIMIT, first.stdout)
            self.assertEqual(json.loads(state.read_text(encoding="utf-8"))["used_today"], 1)

            second = run_cli(base + ["--no-resume", "--limit", "1"], mock.url, tmp_path)
            self.assertEqual(second.returncode, 0, second.stdout + second.stderr)
            used = json.loads(state.read_text(encoding="utf-8"))["used_today"]
            self.assertEqual(used, 2, "счётчик суток не должен обнуляться при --no-resume")
            self.assertEqual(mock.snapshot()["total"], 2)

    def test_day_roll_resets_counter_but_keeps_cache(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "state.json"
            state = StateStore.load(path, "2026-10-15", resume=False)
            state.used_today = 7
            state.processed = ["7707083893"]
            state.cache = {
                "7707083893": {"status": "ok", "name": "ООО Ромашка", "okved_main": "56.10",
                               "okved_extra": [], "error": ""},
            }
            state.save()

            reloaded = StateStore.load(path, "2026-10-16", resume=True)
            self.assertEqual(reloaded.used_today, 0, "новый день — счётчик обнуляется")
            self.assertEqual(reloaded.date, "2026-10-16")
            self.assertIn("inn:7707083893", reloaded.cache, "кэш на новый день сохраняется")

    # ---------- 2. блокировка состояния ----------

    def test_state_lock_blocks_second_run(self) -> None:
        if os.name == "nt":
            self.skipTest("проверка через fcntl выполняется на POSIX")
        import fcntl

        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("A", "7707083893")])
            handle = open(tmp_path / "state.json.lock", "w")  # noqa: SIM115
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                result = run_cli(self._base_args(tmp_path, inp), mock.url, tmp_path)
            finally:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
                handle.close()

            self.assertEqual(result.returncode, 1, result.stdout)
            self.assertIn("используется другим процессом", (result.stdout + result.stderr).lower())
            self.assertEqual(mock.snapshot()["total"], 0, "второй процесс не должен ходить в API")

    # ---------- 3. дубликаты ошибочных ИНН ----------

    def test_duplicate_failed_inns_cost_one_request(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("A", "5555555555"), ("B", "5555555555"),
                                                    ("C", "5555555555")])
            result = run_cli(self._base_args(tmp_path, inp), mock.url, tmp_path)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertEqual(mock.snapshot()["total"], 1, "дубликат ошибочного ИНН не должен повторяться")
            sheet = read_sheets(tmp_path / "o.xlsx")["Все"]
            self.assertEqual(list(sheet["Статус"]), ["error", "error", "error"])

    # ---------- 4. инъекция формул ----------

    def test_formula_injection_is_neutralized(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input_with_text_cell(tmp_path / "in.xlsx", FORMULA, "7707083893")
            out = tmp_path / "o.xlsx"
            result = run_cli(self._base_args(tmp_path, inp), mock.url, tmp_path)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

            sheet = load_workbook(out)["Все"]
            cells = [cell for row in sheet.iter_rows() for cell in row
                     if isinstance(cell.value, str) and "HYPERLINK" in cell.value]
            self.assertTrue(cells, "значение должно остаться в файле (как текст)")
            for cell in cells:
                self.assertNotEqual(cell.data_type, "f", f"формула попала в {cell.coordinate}")
                self.assertTrue(cell.value.startswith("="), "текст не должен искажаться")

    # ---------- 5. нормализованный ИНН ----------

    def test_raw_and_normalized_inn_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("X", " 7707-083893 ")])
            result = run_cli(self._base_args(tmp_path, inp), mock.url, tmp_path)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

            sheet = read_sheets(tmp_path / "o.xlsx")["Все"]
            self.assertIn("Нормализованный ИНН", sheet.columns)
            self.assertEqual(str(sheet["Нормализованный ИНН"].iloc[0]), "7707083893")
            self.assertEqual(str(sheet["ИНН"].iloc[0]), " 7707-083893 ", "исходный ИНН не портим")
            self.assertEqual(str(sheet["Статус"].iloc[0]), "ok")

    # ---------- 6. не-ASCII цифры ----------

    def test_non_ascii_digits_are_not_inn(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("А", ARABIC_INN)])
            result = run_cli(self._base_args(tmp_path, inp), mock.url, tmp_path)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertEqual(mock.snapshot()["total"], 0, "не-ASCII цифры не должны тратить запрос")
            sheet = read_sheets(tmp_path / "o.xlsx")["Все"]
            self.assertEqual(str(sheet["Статус"].iloc[0]), "invalid_id")

    # ---------- 7. запрет http на внешний хост ----------

    def test_plain_http_api_url_refused(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("A", "7707083893")])
            result = run_cli(
                self._base_args(tmp_path, inp),
                "http://api.example.com/suggestions/api/4_1/rs/findById/party",
                tmp_path,
            )
            self.assertEqual(result.returncode, 1, result.stdout)
            self.assertIn("Небезопасный адрес API", result.stdout + result.stderr)

    # ---------- 8. счётчик попыток ----------

    def test_attempts_counter_includes_network_errors(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            inp = make_input(tmp_path / "in.xlsx", [("A", "7707083893")])
            result = run_cli(
                self._base_args(tmp_path, inp, ["--max-retries", "2"]),
                "http://127.0.0.1:1/suggestions/api/4_1/rs/findById/party",
                tmp_path,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("обращений с повторами: 2", result.stdout,
                          "сетевые ошибки должны считаться как обращения")
            sheet = read_sheets(tmp_path / "o.xlsx")["Все"]
            self.assertEqual(str(sheet["Статус"].iloc[0]), "error")
            self.assertIn("network_error", str(sheet["Ошибка"].iloc[0]))

    # ---------- 10. токен из .env ----------

    def test_token_from_env_file_and_env_var_priority(self) -> None:
        """
        .env рядом со скриптом читается, но переменная окружения его перебивает.

        Скрипт копируется в отдельную папку: python-dotenv ищет .env от каталога
        скрипта и вверх по дереву, поэтому тест в общей папке проекта подхватил бы
        и чужой .env, и .env разработчика.
        """
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            shutil.copy(REPO / "dadata_okved_filter.py", tmp_path / "dadata_okved_filter.py")
            make_input(tmp_path / "in.xlsx", [("A", "7707083893")])
            file_token = "fromenvfile1234567890"
            env_token = "fromenvironment0987654321"
            (tmp_path / ".env").write_text(
                f"DADATA_TOKEN={file_token}\nDADATA_API_URL={mock.url}\n", encoding="utf-8")
            clean_env = {k: v for k, v in os.environ.items() if not k.startswith("DADATA_")}
            cmd = [sys.executable, "dadata_okved_filter.py", "--input", "in.xlsx",
                   "--output", "out.xlsx", "--state-file", "state.json",
                   "--max-per-minute", "120", "--sleep", "0", "--verbose"]

            def mask(token: str) -> str:
                return f"{token[:4]}...{token[-4:]}"

            with_env_file = subprocess.run(cmd, cwd=str(tmp_path), env=clean_env,
                                           capture_output=True, text=True, timeout=60)
            self.assertEqual(with_env_file.returncode, 0,
                             with_env_file.stdout + with_env_file.stderr)
            self.assertIn(mask(file_token), with_env_file.stdout,
                          "токен должен браться из .env")
            self.assertNotIn(file_token, (tmp_path / "state.json").read_text(encoding="utf-8"))

            env_wins = subprocess.run(cmd, cwd=str(tmp_path),
                                      env=dict(clean_env, DADATA_TOKEN=env_token),
                                      capture_output=True, text=True, timeout=60)
            self.assertEqual(env_wins.returncode, 0, env_wins.stdout + env_wins.stderr)
            self.assertIn(mask(env_token), env_wins.stdout,
                          "переменная окружения имеет приоритет над .env")
            self.assertNotIn(mask(file_token), env_wins.stdout)

            (tmp_path / ".env").unlink()
            without_token = subprocess.run(cmd, cwd=str(tmp_path), env=clean_env,
                                           capture_output=True, text=True, timeout=60)
            self.assertEqual(without_token.returncode, 1)
            self.assertIn("Не задан токен", without_token.stdout + without_token.stderr)


    # ---------- 10. отчёт о скорости ----------

    def test_reported_speed_does_not_exceed_limit(self) -> None:
        """
        Скорость в строке прогресса считается по фактическим интервалам между
        запросами. Формула «запросов / время от старта» завышала результат:
        3 запроса за ~4 с выдавали 43 rpm при лимите 30 rpm.
        """
        with tempfile.TemporaryDirectory() as tmp, MockDaData() as mock:
            tmp_path = Path(tmp)
            inp = make_input(
                tmp_path / "in.xlsx",
                [("A", "7707083893"), ("B", "7707083894"), ("C", "7707083895")],
            )
            result = run_cli(
                self._base_args(tmp_path, inp, ["--max-per-minute", "30", "--limit", "3"]),
                mock.url,
                tmp_path,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            lines = [ln for ln in result.stdout.splitlines() if "средняя скорость" in ln]
            self.assertTrue(lines, "строка прогресса не найдена")
            found = re.findall(r"средняя скорость: ([0-9.]+) rpm", lines[-1])
            self.assertTrue(found, f"скорость не распознана: {lines[-1]}")
            rpm = float(found[0])
            self.assertLessEqual(
                rpm, 30.0 + 1e-6,
                f"заявленная скорость выше лимита 30 rpm: {lines[-1]}",
            )
            self.assertGreater(rpm, 0.0, "скорость не измерена")


if __name__ == "__main__":
    unittest.main()
