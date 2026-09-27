"""
Sanity tests for luah_signage.presentation's pure (no UNO dependency)
helper `_reconcile_day_modes`. Run with:
python tests/test_presentation_daymode.py (also pytest-compatible)

Everything else in presentation.py requires a live UNO/LibreOffice
connection and is not testable in this environment - see that module's
docstring.
"""
from __future__ import annotations

from luah_signage.presentation import (
    DAY_MODE_CHOL_ONLY,
    DAY_MODE_NONCHOL_ONLY,
    _reconcile_day_modes,
)


def test_no_tags_means_always_visible():
    result = _reconcile_day_modes(set(), set(), slide_count=3)
    assert result == {}


def test_cholonly_only_slide():
    result = _reconcile_day_modes({1}, set(), slide_count=3)
    assert result == {1: DAY_MODE_CHOL_ONLY}


def test_noncholonly_only_slide():
    result = _reconcile_day_modes(set(), {2}, slide_count=3)
    assert result == {2: DAY_MODE_NONCHOL_ONLY}


def test_both_tags_on_same_slide_means_always_visible():
    result = _reconcile_day_modes({0}, {0}, slide_count=1)
    assert result == {}


def test_mixed_deck():
    # slide 0: #CHOLONLY only -> chol-only
    # slide 1: no tags -> always
    # slide 2: #NONCHOLONLY only -> non-chol-only
    # slide 3: both tags -> always
    result = _reconcile_day_modes(
        cholonly_indices={0, 3}, noncholonly_indices={2, 3}, slide_count=4
    )
    assert result == {0: DAY_MODE_CHOL_ONLY, 2: DAY_MODE_NONCHOL_ONLY}


def test_out_of_range_indices_ignored():
    # slide_count=2, but sets reference indices beyond range - should not
    # appear in the result (only indices in range(slide_count) are ever
    # considered).
    result = _reconcile_day_modes({5}, set(), slide_count=2)
    assert result == {}


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
