"""
Sanity tests for luah_signage.time_source (pure Python, no UNO dependency).
Run with: python tests/test_time_source.py (also pytest-compatible)
"""
from __future__ import annotations

import time
from datetime import datetime

from luah_signage.time_source import MockTimeSource, TimeSource, create_time_source


def test_real_time_source_returns_current_time():
    source = TimeSource()
    before = datetime.now()
    result = source.now()
    after = datetime.now()
    assert before <= result <= after


def test_mock_time_source_starts_at_configured_time():
    start = datetime(2026, 9, 25, 14, 30, 0)  # a Friday afternoon
    source = MockTimeSource(start)
    result = source.now()
    # Should be at (or an imperceptibly tiny fraction of a second after)
    # the configured start - allow up to 1 second of test-execution slack.
    assert result >= start
    assert (result - start).total_seconds() < 1.0


def test_mock_time_source_advances_with_real_elapsed_time():
    start = datetime(2026, 9, 25, 14, 30, 0)
    source = MockTimeSource(start)
    time.sleep(0.2)
    result = source.now()
    elapsed = (result - start).total_seconds()
    assert 0.15 <= elapsed <= 1.0  # roughly 0.2s, generous bounds for CI jitter


def test_mock_time_source_is_1to1_speed_not_accelerated():
    start = datetime(2026, 9, 25, 14, 30, 0)
    source = MockTimeSource(start)
    real_before = time.monotonic()
    time.sleep(0.3)
    real_elapsed = time.monotonic() - real_before
    mock_elapsed = (source.now() - start).total_seconds()
    # Mock elapsed should track real elapsed almost exactly (1:1), not be
    # scaled/accelerated.
    assert abs(mock_elapsed - real_elapsed) < 0.1


def test_create_time_source_returns_real_when_unset():
    source = create_time_source(None)
    assert isinstance(source, TimeSource)
    assert not isinstance(source, MockTimeSource)


def test_create_time_source_returns_real_when_empty_string():
    source = create_time_source("")
    assert isinstance(source, TimeSource)
    assert not isinstance(source, MockTimeSource)


def test_create_time_source_returns_mock_when_set():
    source = create_time_source("2026-09-25T14:30:00")
    assert isinstance(source, MockTimeSource)
    assert source.start == datetime(2026, 9, 25, 14, 30, 0)


if __name__ == "__main__":
    import sys
    import traceback

    tests = [obj for name, obj in list(globals().items()) if name.startswith("test_")]
    failures = 0
    for t in tests:
        try:
            t()
            print(f"PASS: {t.__name__}")
        except Exception:
            failures += 1
            print(f"FAIL: {t.__name__}")
            traceback.print_exc()
    print(f"\n{len(tests) - failures}/{len(tests)} passed")
    sys.exit(1 if failures else 0)
