# -*- coding: utf-8 -*-
"""
Юнит-тесты ограничителя частоты DaDataRateLimiter.

Часы и «сон» подменяются фейковыми (clock/sleeper), поэтому тесты не ждут реальное
время, но проверяют именно арифметику лимитов:
  * не более N запросов в любые 60 секунд;
  * минимальный интервал 60 / max_per_minute;
  * реальная пауза = max(--sleep, 60 / max_per_minute);
  * остановка при исчерпании дневного лимита;
  * учёт последнего запроса предыдущего запуска при resume.

Запуск: python3 -m unittest tests.test_rate_limiter -v
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dadata_okved_filter import DaDataRateLimiter, StateStore, WINDOW_SECONDS  # noqa: E402


class FakeClock:
    """Фейковые monotonic/sleep: сон просто сдвигает время вперёд."""

    def __init__(self, start: float = 1000.0) -> None:
        self.now = start
        self.sleeps: list[float] = []

    def monotonic(self) -> float:
        return self.now

    def sleep(self, seconds: float) -> None:
        self.sleeps.append(seconds)
        self.now += seconds


class RateLimiterTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.state_path = Path(self.tmp.name) / "state.json"
        self.tz = timezone(timedelta(hours=3))

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _limiter(self, clock: FakeClock, max_per_minute: int = 30, daily_limit: int = 10_000,
                 extra_sleep: float = 0.0, state: StateStore | None = None) -> DaDataRateLimiter:
        state = state or StateStore.load(self.state_path, "2026-10-15", resume=False)
        return DaDataRateLimiter(
            state=state,
            max_per_minute=max_per_minute,
            daily_limit=daily_limit,
            extra_sleep=extra_sleep,
            clock=clock.monotonic,
            sleeper=clock.sleep,
            wall_clock=lambda: datetime.now(self.tz),
        )

    @staticmethod
    def _max_in_window(times: list[float], window: float = WINDOW_SECONDS) -> int:
        """Максимум запросов в любом окне длины window."""
        worst = 0
        for start in times:
            worst = max(worst, sum(1 for t in times if start <= t < start + window))
        return worst

    def test_min_interval_and_window_limit_for_30_rpm(self) -> None:
        clock = FakeClock()
        limiter = self._limiter(clock, max_per_minute=30)
        stamps: list[float] = []
        for _ in range(35):
            self.assertTrue(limiter.acquire(quiet=True))
            stamps.append(clock.now)

        # Интервал не меньше 2 секунд (30 rpm)
        gaps = [b - a for a, b in zip(stamps, stamps[1:])]
        self.assertTrue(gaps, "должны быть замеры интервалов")
        self.assertGreaterEqual(min(gaps), 2.0 - 1e-9, f"нарушен минимальный интервал: {min(gaps)}")
        # Ни в одном 60-секундном окне не больше 30 запросов
        self.assertLessEqual(self._max_in_window(stamps), 30)
        # Все запросы учтены в дневном счётчике
        self.assertEqual(limiter.state.used_today, 35)
        # Файл состояния сохраняется после каждого запроса
        self.assertTrue(self.state_path.exists())

    def test_extra_sleep_is_max_of_sleep_and_interval(self) -> None:
        # --sleep 5 при 30 rpm -> пауза 5 секунд (больше минимальных 2)
        clock = FakeClock()
        limiter = self._limiter(clock, max_per_minute=30, extra_sleep=5.0)
        self.assertEqual(limiter.effective_pause, 5.0)
        stamps = []
        for _ in range(4):
            limiter.acquire(quiet=True)
            stamps.append(clock.now)
        gaps = [b - a for a, b in zip(stamps, stamps[1:])]
        self.assertGreaterEqual(min(gaps), 5.0 - 1e-9)

        # --sleep 0.5 при 30 rpm -> пауза остаётся 2 секунды (лимит важнее)
        clock2 = FakeClock()
        limiter2 = self._limiter(clock2, max_per_minute=30, extra_sleep=0.5)
        self.assertEqual(limiter2.effective_pause, 2.0)
        stamps2 = []
        for _ in range(3):
            limiter2.acquire(quiet=True)
            stamps2.append(clock2.now)
        gaps2 = [b - a for a, b in zip(stamps2, stamps2[1:])]
        self.assertGreaterEqual(min(gaps2), 2.0 - 1e-9)

    def test_daily_limit_stops_requests(self) -> None:
        clock = FakeClock()
        limiter = self._limiter(clock, max_per_minute=600, daily_limit=3)
        self.assertEqual(limiter.budget_left, 3)
        for _ in range(3):
            self.assertTrue(limiter.acquire(quiet=True))
        self.assertEqual(limiter.budget_left, 0)
        self.assertFalse(limiter.acquire(quiet=True), "четвёртый запрос должен быть запрещён")
        self.assertEqual(limiter.state.used_today, 3)

    def test_resume_respects_previous_request(self) -> None:
        # Предыдущий запуск сделал запрос только что -> окно не пустое
        state = StateStore.load(self.state_path, "2026-10-15", resume=False)
        state.last_request_time = datetime.now(self.tz).isoformat(timespec="seconds")
        state.save()

        clock = FakeClock()
        limiter = self._limiter(clock, max_per_minute=30, state=state)
        wait = limiter.seconds_until_slot()
        self.assertGreater(wait, 0.5, f"после resume запрос не должен уходить сразу (wait={wait})")
        self.assertLessEqual(wait, 2.0 + 1e-6)

        limiter.acquire(quiet=True)
        self.assertGreaterEqual(clock.sleeps[0] if clock.sleeps else 0.0, 0.5)

    def test_window_never_exceeds_limit_for_tighter_rpm(self) -> None:
        # 6 rpm => интервал 10 секунд, окно никогда не переполняется
        clock = FakeClock()
        limiter = self._limiter(clock, max_per_minute=6)
        stamps = []
        for _ in range(10):
            limiter.acquire(quiet=True)
            stamps.append(clock.now)
        gaps = [b - a for a, b in zip(stamps, stamps[1:])]
        self.assertGreaterEqual(min(gaps), 10.0 - 1e-9)
        self.assertLessEqual(self._max_in_window(stamps), 6)


if __name__ == "__main__":
    unittest.main(verbosity=2)
