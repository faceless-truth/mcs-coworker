"""
Tests for Schedule.daily_at + LoadedPlugin._next_daily_at_run
=============================================================
Phase 1 of the Morning Briefing scheduler fix. Daily-at-hour schedules
are now calendar-based: the scheduler computes the next wall-clock HH:00
in the future instead of `now + 86400`. These tests pin that behaviour
plus the boundary cases (already past today, exactly at HH:00, near
midnight, monthly schedules unaffected).
"""
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

# Use a temp DB so the loader's get_plugin_state() does not touch the
# real user database when constructing a LoadedPlugin.
import config as cfg  # noqa: E402

_tmp_db = tempfile.mktemp(suffix=".db")
cfg.DB_PATH = Path(_tmp_db)
cfg.init_db()

from plugin_base import (  # noqa: E402
    AgentPlugin,
    PluginCategory,
    PluginContext,
    PluginResult,
    Schedule,
)


class TestScheduleShape(unittest.TestCase):
    """Schedule construction sets the right flags for each shape."""

    def test_daily_at_is_calendar_based(self):
        sched = Schedule.daily_at(8)
        self.assertTrue(sched.is_calendar_based())
        self.assertEqual(sched.hour_of_day, 8)
        self.assertEqual(sched.label, "Daily at 08:00")

    def test_daily_at_label_zero_pads_hour(self):
        self.assertEqual(Schedule.daily_at(7).label, "Daily at 07:00")
        self.assertEqual(Schedule.daily_at(17).label, "Daily at 17:00")

    def test_every_minutes_is_not_calendar_based(self):
        sched = Schedule.every_minutes(5)
        self.assertFalse(sched.is_calendar_based())
        self.assertIsNone(sched.hour_of_day)
        self.assertEqual(sched.interval_seconds, 300)

    def test_every_hours_is_not_calendar_based(self):
        sched = Schedule.every_hours(4)
        self.assertFalse(sched.is_calendar_based())
        self.assertIsNone(sched.hour_of_day)

    def test_monthly_on_day_is_calendar_based(self):
        sched = Schedule.monthly_on_day(15)
        self.assertTrue(sched.is_calendar_based())
        # Monthly schedules use day_of_month/months_interval, NOT hour_of_day.
        self.assertIsNone(sched.hour_of_day)
        self.assertEqual(sched.day_of_month, 15)
        self.assertEqual(sched.months_interval, 1)

    def test_quarterly_on_day_is_calendar_based(self):
        sched = Schedule.quarterly_on_day(1)
        self.assertTrue(sched.is_calendar_based())
        self.assertIsNone(sched.hour_of_day)
        self.assertEqual(sched.months_interval, 3)

    def test_manual_only_is_not_calendar_based(self):
        sched = Schedule.manual_only()
        self.assertFalse(sched.is_calendar_based())
        self.assertIsNone(sched.hour_of_day)


# ── Plugin used to exercise LoadedPlugin._next_daily_at_run ──────────────
# Constructed with daily_at(8) so we get a real LoadedPlugin to call into.

class _DailyAt8Plugin(AgentPlugin):
    PLUGIN_ID = "test_daily_at_8"
    name = "Test Daily At 8"
    description = "test fixture"
    icon = "T"
    version = "0.0.1"
    requires_graph = False
    requires_claude = False
    default_schedule = Schedule.daily_at(8)
    category = PluginCategory.UNIVERSAL

    def load(self, context: PluginContext) -> bool:
        return True

    def run(self, context: PluginContext) -> PluginResult:
        return PluginResult(success=True, summary="noop")


class _MonthlyOn15Plugin(AgentPlugin):
    PLUGIN_ID = "test_monthly_on_15"
    name = "Test Monthly 15"
    description = "test fixture"
    icon = "M"
    version = "0.0.1"
    requires_graph = False
    requires_claude = False
    default_schedule = Schedule.monthly_on_day(15)
    category = PluginCategory.UNIVERSAL

    def load(self, context: PluginContext) -> bool:
        return True

    def run(self, context: PluginContext) -> PluginResult:
        return PluginResult(success=True, summary="noop")


class TestNextDailyAtRun(unittest.TestCase):
    """LoadedPlugin._next_daily_at_run computes the right wall-clock target."""

    def setUp(self):
        # Import here so the temp DB above is in place before LoadedPlugin
        # touches the settings table.
        from plugin_loader import LoadedPlugin
        self.lp = LoadedPlugin(_DailyAt8Plugin, "test_daily_at_8")

    def _ts(self, *args) -> float:
        return datetime(*args).timestamp()

    def _expected(self, *args) -> float:
        return datetime(*args).timestamp()

    def test_now_before_target_fires_today(self):
        # 06:00 today, target 08:00 -> fires at 08:00 today.
        now = datetime(2026, 5, 5, 6, 0, 0)
        result = self.lp._next_daily_at_run(8, now_fn=lambda: now)
        self.assertEqual(result, self._expected(2026, 5, 5, 8, 0, 0))

    def test_now_after_target_fires_tomorrow(self):
        # 09:30 today, target 08:00 -> fires at 08:00 tomorrow.
        now = datetime(2026, 5, 5, 9, 30, 0)
        result = self.lp._next_daily_at_run(8, now_fn=lambda: now)
        self.assertEqual(result, self._expected(2026, 5, 6, 8, 0, 0))

    def test_now_exactly_at_target_fires_tomorrow(self):
        # Boundary: exactly 08:00:00 -> tomorrow, not today. A scheduler
        # tick landing on the boundary must not re-fire the same hour.
        now = datetime(2026, 5, 5, 8, 0, 0)
        result = self.lp._next_daily_at_run(8, now_fn=lambda: now)
        self.assertEqual(result, self._expected(2026, 5, 6, 8, 0, 0))

    def test_now_after_target_within_same_minute_fires_tomorrow(self):
        # 08:00:01 -> tomorrow, just to be paranoid about the equality.
        now = datetime(2026, 5, 5, 8, 0, 1)
        result = self.lp._next_daily_at_run(8, now_fn=lambda: now)
        self.assertEqual(result, self._expected(2026, 5, 6, 8, 0, 0))

    def test_late_evening_crosses_midnight_cleanly(self):
        # 23:55, target 08:00 -> 08:00 the next calendar day.
        now = datetime(2026, 5, 5, 23, 55, 0)
        result = self.lp._next_daily_at_run(8, now_fn=lambda: now)
        self.assertEqual(result, self._expected(2026, 5, 6, 8, 0, 0))

    def test_end_of_month_crosses_into_next_month(self):
        # 31 May 2026, 09:00 -> 1 Jun 2026 08:00. Confirms we use
        # timedelta(days=1) rather than naive day arithmetic.
        now = datetime(2026, 5, 31, 9, 0, 0)
        result = self.lp._next_daily_at_run(8, now_fn=lambda: now)
        self.assertEqual(result, self._expected(2026, 6, 1, 8, 0, 0))

    def test_end_of_year_crosses_into_next_year(self):
        # 31 Dec 2026, 09:00 -> 1 Jan 2027 08:00.
        now = datetime(2026, 12, 31, 9, 0, 0)
        result = self.lp._next_daily_at_run(8, now_fn=lambda: now)
        self.assertEqual(result, self._expected(2027, 1, 1, 8, 0, 0))


class TestNextCalendarRunDispatch(unittest.TestCase):
    """_next_calendar_run routes daily-at vs monthly correctly."""

    def test_daily_at_routed_to_daily_calculator(self):
        from plugin_loader import LoadedPlugin
        lp = LoadedPlugin(_DailyAt8Plugin, "test_daily_at_8")
        now = datetime(2026, 5, 5, 9, 30, 0)
        result = lp._next_calendar_run(now_fn=lambda: now)
        # Daily-at(8) at 09:30 -> tomorrow 08:00. If the dispatch were
        # broken and this fell through to the monthly branch, we'd get
        # something else (the monthly branch ignores hour_of_day).
        self.assertEqual(result, datetime(2026, 5, 6, 8, 0, 0).timestamp())

    def test_monthly_on_day_still_uses_month_step_path(self):
        # Regression: existing monthly schedules must keep working.
        from plugin_loader import LoadedPlugin
        lp = LoadedPlugin(_MonthlyOn15Plugin, "test_monthly_on_15")
        # 5 May 2026 -> next 15th in the future is 15 May 2026 08:00.
        now = datetime(2026, 5, 5, 12, 0, 0)
        result = lp._next_calendar_run(now_fn=lambda: now)
        self.assertEqual(result, datetime(2026, 5, 15, 8, 0, 0).timestamp())

        # 20 May 2026 -> 15th already passed, next is 15 Jun 2026.
        now = datetime(2026, 5, 20, 12, 0, 0)
        result = lp._next_calendar_run(now_fn=lambda: now)
        self.assertEqual(result, datetime(2026, 6, 15, 8, 0, 0).timestamp())


if __name__ == "__main__":
    unittest.main()
