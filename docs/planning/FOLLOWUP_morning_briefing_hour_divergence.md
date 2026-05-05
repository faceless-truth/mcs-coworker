# Follow-up: `fix(morning_briefing): Schedule.daily_at hour and settings briefing_hour disagree`

**Status:** filed, deferred to Phase 2 of the Morning Briefing rework.
**Filed:** 2026-05-05.
**Filed by:** Elio + Claude during the daily_at scheduler fix (Phase 1 of two).

## Summary

In `plugins/plugin_morning_briefing.py` two values control "when does the briefing fire?" and they can drift apart silently:

- The class attribute `default_schedule = Schedule.daily_at(8)` (line 82) — read by the scheduler to decide *when to call `run()`*.
- The plugin setting `briefing_hour` (default `"8"`, line 94) — read inside `run()` itself: `target_hour = int(self.get_plugin_setting("briefing_hour", "8"))`. If `now.hour != target_hour`, the function returns immediately.

If an accountant changes `briefing_hour` to, say, `7` in Settings, the **scheduler** still wakes the plugin at 08:00 (because that's baked into the class attribute at module load) and `run()` then bails because `now.hour (8) != target_hour (7)`. The briefing never fires. No error, no log entry — silent.

The reverse is just as bad: if the class attribute is `daily_at(8)` but `briefing_hour` is `9`, the scheduler will tick at 08:00 and `run()` will return without action — and the next scheduler tick is 24 h later, so the 09:00 hour is never observed.

## Why this is filed but not fixed yet

Phase 1 (the `feat(scheduler): make daily_at calendar-based and honour the wall-clock hour` change) is scheduler-only — no plugin files touched. This issue is the natural follow-up.

## Why deferring is fine

Phase 2 of the Morning Briefing rework will redesign the delivery path (move off the Outlook draft, settle on a single source of truth for the trigger time). The cleanest fix for this divergence falls naturally out of that redesign — likely by deleting either the class attribute or the setting, depending on which delivery path Phase 2 picks.

## Acceptance criteria

- One source of truth for "what hour does the Morning Briefing fire?".
- Changing it through whichever surface remains (Settings UI or class attribute) reflects in the next scheduled run on the next app start without a code change to the other.
- A unit test that proves a setting change moves the next-run timestamp.

## Files of interest

- `plugins/plugin_morning_briefing.py:82` — `default_schedule = Schedule.daily_at(8)`
- `plugins/plugin_morning_briefing.py:94` — `target_hour = int(self.get_plugin_setting("briefing_hour", "8"))`
