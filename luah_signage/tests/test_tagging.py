"""
Sanity tests for luah_signage.tagging (pure Python, no UNO dependency).
Run with: python tests/test_tagging.py (also pytest-compatible)
"""
from __future__ import annotations

from datetime import datetime

from pyzmanim.noaa_calculator import Location

from luah_signage import tagging
from luah_signage.zman_context import build_zman_context

JERUSALEM = Location(latitude=31.7683, longitude=35.2137, elevation=754)


class FakeShape:
    def __init__(self, text: str):
        self._text = text

    def get_text(self) -> str:
        return self._text

    def set_text(self, text: str) -> None:
        self._text = text


def _ctx(dt: datetime):
    return build_zman_context(dt, JERUSALEM, eretz_yisroel=True, timezone="Asia/Jerusalem")


def test_tag_pattern_parses_offset():
    matches = list(tagging.TAG_PATTERN.finditer("#SUNSET-18 and #SUNRISE+5"))
    assert len(matches) == 2
    assert matches[0].group(1) == "SUNSET" and matches[0].group(2) == "-18"
    assert matches[1].group(1) == "SUNRISE" and matches[1].group(2) == "+5"


def test_tag_pattern_parses_hhmm_offset():
    matches = list(tagging.TAG_PATTERN.finditer("#SUNSET+00:30 and #TZAIS-01:15"))
    assert len(matches) == 2
    assert matches[0].group(1) == "SUNSET" and matches[0].group(2) == "+00:30"
    assert matches[1].group(1) == "TZAIS" and matches[1].group(2) == "-01:15"


def test_render_template_sunrise_sunset():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    text = tagging.render_template("Sunrise: #SUNRISE Sunset: #SUNSET", ctx)
    assert "Sunrise: 05" in text  # known ~05:34 IDT solstice sunrise
    assert "Sunset: 19" in text or "Sunset: 20" in text


def test_render_template_offset_changes_result():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    base = tagging.render_template("#SUNSET", ctx)
    minus18 = tagging.render_template("#SUNSET-18", ctx)
    assert base != minus18


def test_hhmm_offset_matches_equivalent_plain_minutes_offset():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    plain_minutes = tagging.render_template("#SUNSET+30", ctx)
    hhmm = tagging.render_template("#SUNSET+00:30", ctx)
    assert plain_minutes == hhmm


def test_hhmm_offset_with_hours_and_minus_sign():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    hhmm = tagging.render_template("#SUNSET-01:15", ctx)
    plain_minutes = tagging.render_template("#SUNSET-75", ctx)
    assert hhmm == plain_minutes


def test_parse_offset_minutes_helper():
    assert tagging._parse_offset_minutes(None) == 0
    assert tagging._parse_offset_minutes("+18") == 18
    assert tagging._parse_offset_minutes("-18") == -18
    assert tagging._parse_offset_minutes("+00:30") == 30
    assert tagging._parse_offset_minutes("-01:15") == -75
    assert tagging._parse_offset_minutes("+02:05") == 125


def test_unknown_tag_left_untouched():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    text = tagging.render_template("keep #NOTAREALTAG please", ctx)
    assert "#NOTAREALTAG" in text


def test_scan_shapes_for_tags():
    tagged = FakeShape("#HEBDATE")
    untagged = FakeShape("just some text")
    tracked = tagging.scan_shapes_for_tags([tagged, untagged])
    assert len(tracked) == 1
    assert tracked[0].shape is tagged


def test_tracked_shape_refresh_reuses_template():
    shape = FakeShape("Time: #SUNRISE")
    tracked = tagging.TrackedShape(shape=shape, template="Time: #SUNRISE")
    tracked.refresh(_ctx(datetime(2024, 6, 21, 12, 0, 0)))
    summer_text = shape.get_text()
    tracked.refresh(_ctx(datetime(2024, 12, 21, 12, 0, 0)))
    winter_text = shape.get_text()
    assert summer_text != winter_text
    assert tracked.template == "Time: #SUNRISE"  # template never mutated


def test_parsha_tag_produces_nonempty_hebrew_text():
    # Thursday before Parshas Shemot (Shabbos Jan 6 2024)
    ctx = _ctx(datetime(2024, 1, 4, 12, 0, 0))
    text = tagging.render_template("#PARSHA", ctx)
    assert text.strip()
    assert "שמות" in text


def test_dafyomi_tag_matches_known_start_date():
    ctx = _ctx(datetime(1923, 9, 11, 12, 0, 0))
    text = tagging.render_template("#DAFYOMI", ctx)
    assert "ברכות" in text


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
