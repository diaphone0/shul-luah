"""
Sanity tests for luah_signage.presentation's pure (no UNO dependency)
helper `_next_visible_slide_index`. Run with:
python tests/test_presentation_slide_visibility.py (also pytest-compatible)

Everything else in presentation.py requires a live UNO/LibreOffice
connection and is not testable in this environment - see that module's
docstring.
"""
from __future__ import annotations

from luah_signage.presentation import _next_visible_slide_index


def test_all_visible_simple_forward_step():
    flags = [True, True, True]
    assert _next_visible_slide_index(0, flags) == (1, False)
    assert _next_visible_slide_index(1, flags) == (2, False)


def test_all_visible_wraps_at_end():
    flags = [True, True, True]
    assert _next_visible_slide_index(2, flags) == (0, True)


def test_skips_single_hidden_slide():
    # slide 0 hidden, 1 & 2 visible - matches the user's exact reported
    # scenario (#CHOLONLY slide 0 hidden on a non-chol day).
    flags = [False, True, True]
    # from slide 1 -> slide 2 (no hidden slides in between, no wrap)
    assert _next_visible_slide_index(1, flags) == (2, False)
    # from slide 2 -> wraps around, skips hidden slide 0, lands on slide 1
    assert _next_visible_slide_index(2, flags) == (1, True)


def test_skips_multiple_consecutive_hidden_slides():
    flags = [True, False, False, True]
    assert _next_visible_slide_index(0, flags) == (3, False)
    assert _next_visible_slide_index(3, flags) == (0, True)


def test_hidden_slide_at_very_end_of_deck():
    # last slide hidden - wraparound must be detected correctly even
    # though the hidden slide is exactly at the end, not slide 0.
    flags = [True, True, False]
    assert _next_visible_slide_index(0, flags) == (1, False)
    assert _next_visible_slide_index(1, flags) == (0, True)  # skips hidden slide 2, wraps


def test_only_one_visible_slide_wraps_to_itself():
    flags = [False, True, False]
    # from the one visible slide, the only place to go is itself, via a wrap
    assert _next_visible_slide_index(1, flags) == (1, True)


def test_no_visible_slides_returns_none():
    flags = [False, False, False]
    assert _next_visible_slide_index(0, flags) == (None, False)


def test_empty_deck_returns_none():
    assert _next_visible_slide_index(0, []) == (None, False)


def test_current_index_minus_one_finds_first_visible_slide():
    # start_slideshow()/_warm_up_slides() use current_index=-1 as a
    # "find the first visible slide" query (see presentation.py) - this
    # relies on the scan starting at offset 1 from -1, i.e. index 0, and
    # NEVER reporting `wrapped=True` for this special call (there is no
    # meaningful "wrap" concept when there's no real current slide yet).
    assert _next_visible_slide_index(-1, [True, True, True]) == (0, False)
    # slide 0 hidden (e.g. #CHOLONLY on a non-chol day) - must skip
    # straight to the first VISIBLE slide, not literal index 0. This is
    # the exact scenario that caused the live slideshow to incorrectly
    # start on a hidden #CHOLONLY slide (see repo memory for the full
    # user-reported bug and diagnosis).
    assert _next_visible_slide_index(-1, [False, True, True]) == (1, False)
    assert _next_visible_slide_index(-1, [False, False, True]) == (2, False)
    assert _next_visible_slide_index(-1, [False, False, False]) == (None, False)


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
