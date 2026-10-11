"""
Sanity tests for luah_signage.presentation_pdf's pure (no UNO dependency)
helpers `slide_indices_for_day` and `to_page_range`. Run with:
python tests/test_presentation_pdf.py (also pytest-compatible)

Everything else in presentation_pdf.py (export_day_pdfs / _store_to_pdf)
requires a live LibreOffice document and is not testable in this
environment - see that module's docstring.
"""
from __future__ import annotations

from luah_signage.presentation_pdf import (
    DAY_MODE_CHOL_ONLY,
    DAY_MODE_NONCHOL_ONLY,
    DAY_MODE_ALWAYS_HIDDEN,
    slide_indices_for_day,
    to_page_range,
)


def test_weekday_contains_cholonly_and_always():
    # slide 0 always-visible, slide 1 CHOLONLY, slide 2 NONCHOLONLY,
    # slide 3 HIDDEN.
    modes = {
        1: DAY_MODE_CHOL_ONLY,
        2: DAY_MODE_NONCHOL_ONLY,
        3: DAY_MODE_ALWAYS_HIDDEN,
    }
    assert slide_indices_for_day(4, modes, non_chol=False) == [0, 1]


def test_shabbos_contains_noncholonly_and_always():
    modes = {
        1: DAY_MODE_CHOL_ONLY,
        2: DAY_MODE_NONCHOL_ONLY,
        3: DAY_MODE_ALWAYS_HIDDEN,
    }
    assert slide_indices_for_day(4, modes, non_chol=True) == [0, 2]


def test_always_visible_appears_in_both_pdfs():
    # slide 0 always-visible, slide 1 CHOLONLY, slide 2 NONCHOLONLY.
    modes = {1: DAY_MODE_CHOL_ONLY, 2: DAY_MODE_NONCHOL_ONLY}
    assert slide_indices_for_day(3, modes, non_chol=False) == [0, 1]
    assert slide_indices_for_day(3, modes, non_chol=True) == [0, 2]


def test_hidden_never_included():
    modes = {0: DAY_MODE_ALWAYS_HIDDEN, 1: DAY_MODE_ALWAYS_HIDDEN}
    assert slide_indices_for_day(2, modes, non_chol=False) == []
    assert slide_indices_for_day(2, modes, non_chol=True) == []


def test_config_slide_hidden_via_always_hidden():
    # A #CONFIG slide is auto-hidden and lands in ALWAYS_HIDDEN - must
    # never appear in either PDF.
    modes = {0: DAY_MODE_ALWAYS_HIDDEN, 1: DAY_MODE_CHOL_ONLY}
    assert slide_indices_for_day(2, modes, non_chol=False) == [1]
    assert slide_indices_for_day(2, modes, non_chol=True) == []


def test_all_always_visible_both_pdfs_identical():
    assert slide_indices_for_day(3, {}, non_chol=False) == [0, 1, 2]
    assert slide_indices_for_day(3, {}, non_chol=True) == [0, 1, 2]


def test_empty_deck():
    assert slide_indices_for_day(0, {}, non_chol=False) == []
    assert slide_indices_for_day(0, {}, non_chol=True) == []


def test_to_page_range_empty():
    assert to_page_range([]) == ""


def test_to_page_range_single_one_based():
    assert to_page_range([0]) == "1"
    assert to_page_range([6]) == "7"


def test_to_page_range_runs_and_singles():
    assert to_page_range([0, 2, 3, 6]) == "1,3-4,7"
    assert to_page_range([0, 1, 2, 5, 6]) == "1-3,6-7"


def test_to_page_range_unsorted_input():
    assert to_page_range([3, 1, 2]) == "2-4"
    assert to_page_range([4, 0, 1]) == "1-2,5"


def test_day_modes_match_presentation_constants():
    # Keep the exported-values in sync with presentation.py's own constants.
    import luah_signage.presentation as pres

    assert DAY_MODE_CHOL_ONLY == pres.DAY_MODE_CHOL_ONLY
    assert DAY_MODE_NONCHOL_ONLY == pres.DAY_MODE_NONCHOL_ONLY
    assert DAY_MODE_ALWAYS_HIDDEN == pres.DAY_MODE_ALWAYS_HIDDEN


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