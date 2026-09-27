"""
Wall-clock time source abstraction - lets the app run against a mocked
starting date/time (advancing in real seconds, just like a real clock)
instead of the actual system clock, for testing time/date-dependent
behavior (Hebrew calendar transitions, zmanim, #CHOLONLY/#NONCHOLONLY
slide switching, etc.) without touching the OS clock.

Design: only the WALL CLOCK used for calendar/zmanim calculations
(datetime.now()) is mockable, via TimeSource.now(). time.monotonic() -
used purely for real-time tick/interval scheduling (clock_tick_seconds,
content_tick_seconds, reload_poll_seconds, slide advance durations, etc.,
all in app.py/presentation.py) - is intentionally left completely
untouched/always real everywhere else in the app, so the app's actual
refresh cadence and pacing behave identically whether or not a mock time
is configured; only the calendar date/time itself is shifted. This
matches "set a mock start date/time, then every real passing second
increments it by one simulated second" (1:1 speed, no acceleration).
"""
from __future__ import annotations

import time
from datetime import datetime, timedelta


class TimeSource:
    """Returns the current wall-clock time. Real/base implementation -
    simply delegates to datetime.now()."""

    def now(self) -> datetime:
        return datetime.now()


class MockTimeSource(TimeSource):
    """Returns a simulated wall-clock time: `start` plus however many real
    seconds have elapsed (measured via time.monotonic(), immune to the
    system clock itself being changed) since this object was constructed.
    1 real second == 1 simulated second (no acceleration/deceleration) -
    lets you test date/time-dependent behavior as if the system clock
    were set to `start` and running normally, without touching the actual
    OS clock at all."""

    def __init__(self, start: datetime):
        self.start = start
        self._monotonic_start = time.monotonic()

    def now(self) -> datetime:
        elapsed = time.monotonic() - self._monotonic_start
        return self.start + timedelta(seconds=elapsed)


def create_time_source(mock_start_datetime: str | None) -> TimeSource:
    """Builds the appropriate TimeSource based on config:
    - `mock_start_datetime` is None/empty -> a real TimeSource.
    - otherwise -> a MockTimeSource seeded by parsing `mock_start_datetime`
      as an ISO-8601 datetime string (e.g. "2026-09-25T14:30:00" for a
      Friday afternoon, to test Erev-Shabbos/candle-lighting/#NONCHOLONLY
      slide-switching behavior without waiting for a real Friday or
      changing the system clock)."""
    if not mock_start_datetime:
        return TimeSource()
    start = datetime.fromisoformat(mock_start_datetime)
    return MockTimeSource(start)
