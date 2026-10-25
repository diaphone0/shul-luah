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

    def set_text_range(self, start: int, end: int, text: str) -> None:
        self._text = self._text[:start] + text + self._text[end:]


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


def test_tag_pattern_parses_newline_format_suffix():
    matches = list(tagging.TAG_PATTERN.finditer("#DAYINFO:n and #DAYINFO:na"))
    assert len(matches) == 2
    assert matches[0].group(1) == "DAYINFO" and matches[0].group(5) == "n"
    assert matches[1].group(1) == "DAYINFO" and matches[1].group(5) == "na"


def test_tag_pattern_no_newline_format_suffix_by_default():
    matches = list(tagging.TAG_PATTERN.finditer("#DAYINFO"))
    assert len(matches) == 1
    assert matches[0].group(5) is None


def test_tag_pattern_does_not_misfire_on_unrelated_colon_suffix():
    # ":nice" starts with "n" but is NOT the ":n"/":na" suffix - the
    # trailing \b must prevent a false-positive partial match.
    matches = list(tagging.TAG_PATTERN.finditer("#DAYINFO:nice"))
    assert len(matches) == 1
    assert matches[0].group(0) == "#DAYINFO"
    assert matches[0].group(5) is None


def test_tag_pattern_parses_context_suffix():
    matches = list(tagging.TAG_PATTERN.finditer("#SUNSET-30E and #SUNRISE+0S and #TZAIS-18N"))
    assert len(matches) == 3
    assert matches[0].group(1) == "SUNSET" and matches[0].group(2) == "-30" and matches[0].group(3) == "E"
    assert matches[1].group(1) == "SUNRISE" and matches[1].group(2) == "+0" and matches[1].group(3) == "S"
    assert matches[2].group(1) == "TZAIS" and matches[2].group(2) == "-18" and matches[2].group(3) == "N"


def test_tag_pattern_no_context_suffix_by_default():
    matches = list(tagging.TAG_PATTERN.finditer("#SUNSET-30"))
    assert len(matches) == 1
    assert matches[0].group(3) is None


def test_tag_pattern_context_suffix_requires_an_explicit_offset():
    # With no offset digits, a trailing "S"/"E"/"N" is just part of the tag
    # NAME itself (tag names are free to contain any uppercase letter) -
    # there is no way to request a context without an offset present.
    matches = list(tagging.TAG_PATTERN.finditer("#SUNSETS"))
    assert len(matches) == 1
    assert matches[0].group(1) == "SUNSETS"
    assert matches[0].group(2) is None
    assert matches[0].group(3) is None


def test_apply_newline_format_prepends_and_appends():
    assert tagging._apply_newline_format("value", "n") == "\nvalue"
    assert tagging._apply_newline_format("value", "na") == "value\n"
    assert tagging._apply_newline_format("value", None) == "value"


def test_apply_newline_format_is_noop_for_empty_value():
    assert tagging._apply_newline_format("", "n") == ""
    assert tagging._apply_newline_format("", "na") == ""


def test_render_template_newline_n_prepends_newline():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    base_value = tagging.render_template("#SUNRISE", ctx)
    result = tagging.render_template("Sunrise: #SUNRISE:n", ctx)
    assert result == f"Sunrise: \n{base_value}"


def test_render_template_newline_na_appends_newline():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    base_value = tagging.render_template("#SUNRISE", ctx)
    result = tagging.render_template("#SUNRISE:na after", ctx)
    assert result == f"{base_value}\n after"


def test_render_template_sunrise_sunset():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    text = tagging.render_template("Sunrise: #SUNRISE Sunset: #SUNSET", ctx)
    assert "Sunrise: 05" in text  # known ~05:34 IDT solstice sunrise
    assert "Sunset: 19" in text or "Sunset: 20" in text


def test_fmt_time_minute_ceil_default_truncates():
    from pyzmanim.hebrewcalendar import convert_date

    hd = convert_date(datetime(2024, 6, 21, 18, 30, 45))
    assert tagging._fmt_time(hd) == "18:30"


def test_fmt_time_minute_ceil_rounds_up_at_2_seconds():
    from pyzmanim.hebrewcalendar import convert_date

    hd = convert_date(datetime(2024, 6, 21, 18, 30, 2))
    assert tagging._fmt_time(hd, minute_ceil=True) == "18:31"


def test_fmt_time_minute_ceil_does_not_round_up_below_2_seconds():
    from pyzmanim.hebrewcalendar import convert_date

    hd0 = convert_date(datetime(2024, 6, 21, 18, 30, 0))
    hd1 = convert_date(datetime(2024, 6, 21, 18, 30, 1))
    assert tagging._fmt_time(hd0, minute_ceil=True) == "18:30"
    assert tagging._fmt_time(hd1, minute_ceil=True) == "18:30"


def test_fmt_time_minute_ceil_rolls_over_hour_boundary():
    from pyzmanim.hebrewcalendar import convert_date

    hd = convert_date(datetime(2024, 6, 21, 18, 59, 30))
    assert tagging._fmt_time(hd, minute_ceil=True) == "19:00"


def test_tzaisyeshiva_tag_is_after_sunset():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    sunset_text = tagging.render_template("#SUNSET", ctx)
    tzais_text = tagging.render_template("#TZAISYESHIVA", ctx)
    assert tzais_text.strip()
    assert tzais_text > sunset_text  # later clock time, same day - string compare works for "HH:MM"


def test_tzaisyeshiva_tag_differs_summer_vs_winter_offset_from_sunset():
    # Summer (day > 12h): tzais should be MORE than 18 minutes after
    # sunset (18 "zmanis"/proportional minutes, which run longer than
    # real minutes on a long day). Winter (day <= 12h): exactly 18 fixed
    # real minutes after sunset. Checked directly via pyzmanim (not just
    # string comparison) since minute-level text rendering can't easily
    # distinguish "exactly 18" from "a bit more than 18".
    from pyzmanim import zmanim
    from pyzmanim.hebrewcalendar import hdate_gregorian

    summer_ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    winter_ctx = _ctx(datetime(2024, 12, 21, 12, 0, 0))

    summer_sunset = hdate_gregorian(zmanim.getsunset(summer_ctx.now, JERUSALEM))
    summer_tzais = hdate_gregorian(zmanim.gettzaisyeshiva(summer_ctx.now, JERUSALEM))
    winter_sunset = hdate_gregorian(zmanim.getsunset(winter_ctx.now, JERUSALEM))
    winter_tzais = hdate_gregorian(zmanim.gettzaisyeshiva(winter_ctx.now, JERUSALEM))

    summer_delta = (summer_tzais - summer_sunset).total_seconds() / 60
    winter_delta = (winter_tzais - winter_sunset).total_seconds() / 60
    assert summer_delta > 18
    assert abs(winter_delta - 18) < 0.01


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


def test_context_for_suffix_none_or_n_returns_ctx_unchanged():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    assert tagging._context_for_suffix(ctx, None) is ctx
    assert tagging._context_for_suffix(ctx, "N") is ctx


def test_context_for_suffix_s_swaps_now_for_shabbos():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    effective = tagging._context_for_suffix(ctx, "S")
    assert effective.now == ctx.shabbos
    assert effective.shabbos == ctx.shabbos
    assert effective.erev_shabbos == ctx.erev_shabbos


def test_context_for_suffix_e_swaps_now_for_erev_shabbos():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    effective = tagging._context_for_suffix(ctx, "E")
    assert effective.now == ctx.erev_shabbos


def test_sunset_minus30_erev_shabbos_context_matches_shabbos_tag():
    # "#SUNSET-30E" should be identical to "#SHABBOS-30", since #SHABBOS is
    # literally elevation-sunset on Erev Shabbos.
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    via_context_suffix = tagging.render_template("#SUNSET-30E", ctx)
    via_shabbos_tag = tagging.render_template("#SHABBOS-30", ctx)
    assert via_context_suffix == via_shabbos_tag


def test_sunrise_zero_offset_shabbos_context_matches_direct_computation():
    from pyzmanim import zmanim

    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    via_context_suffix = tagging.render_template("#SUNRISE+0S", ctx)
    expected = tagging._fmt_time(zmanim.getsunrise(ctx.shabbos, JERUSALEM))
    assert via_context_suffix == expected


def test_context_suffix_n_is_equivalent_to_omitting_it():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    with_n = tagging.render_template("#SUNSET-18N", ctx)
    without = tagging.render_template("#SUNSET-18", ctx)
    assert with_n == without


def test_tag_pattern_parses_auto_context_suffix():
    matches = list(tagging.TAG_PATTERN.finditer("#SUNSET-30A"))
    assert len(matches) == 1
    assert matches[0].group(1) == "SUNSET"
    assert matches[0].group(2) == "-30"
    assert matches[0].group(3) == "A"


def test_context_for_suffix_auto_on_friday_matches_erev_shabbos():
    ctx = _ctx(datetime(2024, 6, 21, 10, 0, 0))  # Friday
    assert ctx.now.wday == 6
    effective = tagging._context_for_suffix(ctx, "A")
    assert effective.now == ctx.erev_shabbos


def test_context_for_suffix_auto_on_shabbos_matches_shabbos():
    ctx = _ctx(datetime(2024, 6, 22, 10, 0, 0))  # Saturday
    assert ctx.now.wday == 0
    effective = tagging._context_for_suffix(ctx, "A")
    assert effective.now == ctx.shabbos


def test_context_for_suffix_auto_on_chol_weekday_matches_shabbos():
    ctx = _ctx(datetime(2024, 6, 19, 10, 0, 0))  # Wednesday, ordinary chol day
    assert ctx.today_is_chol
    effective = tagging._context_for_suffix(ctx, "A")
    assert effective.now == ctx.shabbos


def test_sunset_auto_context_matches_s_or_e_per_day_type():
    chol_ctx = _ctx(datetime(2024, 6, 19, 10, 0, 0))
    shabbos_ctx = _ctx(datetime(2024, 6, 22, 10, 0, 0))
    erev_ctx = _ctx(datetime(2024, 6, 21, 10, 0, 0))

    assert tagging.render_template("#SUNSET+0A", chol_ctx) == tagging.render_template("#SUNSET+0S", chol_ctx)
    assert tagging.render_template("#SUNSET+0A", shabbos_ctx) == tagging.render_template("#SUNSET+0S", shabbos_ctx)
    assert tagging.render_template("#SUNSET+0A", erev_ctx) == tagging.render_template("#SUNSET+0E", erev_ctx)


def test_context_for_suffix_auto_on_non_friday_erev_yom_tov_matches_erev_shabbos():
    # Erev Pesach 2024 (22 Apr 2024) falls on a MONDAY, not a Friday - "A"
    # must resolve to the erev-shabbos context here too, since erev_shabbos
    # isn't tied to wday==6 (zman_context.py derives it from
    # is_candle_lighting, which also fires for several Yom Tov eves).
    ctx = _ctx(datetime(2024, 4, 22, 10, 0, 0))
    assert ctx.now.wday != 6
    assert ctx.now == ctx.erev_shabbos
    effective = tagging._context_for_suffix(ctx, "A")
    assert effective.now == ctx.erev_shabbos
    assert tagging.render_template("#SUNSET+0A", ctx) == tagging.render_template("#SUNSET+0E", ctx)


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


class RecordingFakeShape(FakeShape):
    """FakeShape that records every set_text_range call (start, end, text)
    - used to verify TrackedShape.refresh only ever touches a tag's own
    span, never the surrounding static text (the formatting-preservation
    guarantee - see uno_shapes.UnoTextShape.set_text_range's docstring)."""

    def __init__(self, text: str):
        super().__init__(text)
        self.range_calls: list[tuple[int, int, str]] = []

    def set_text_range(self, start: int, end: int, text: str) -> None:
        self.range_calls.append((start, end, text))
        super().set_text_range(start, end, text)


def test_tracked_shape_refresh_only_touches_tag_span_not_static_text():
    # Mirrors the user-reported bug: a bold static label next to a
    # non-bold tag ("בבלי: #DAFYOMIBV") lost the label's bold formatting
    # because the old implementation called shape.set_text(whole_string).
    # The fix must only ever call set_text_range for the TAG's own
    # character span, never touch/replace the "Label: " prefix at all.
    shape = RecordingFakeShape("Label: #SUNRISE")
    tracked = tagging.TrackedShape(shape=shape, template="Label: #SUNRISE")
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    tracked.refresh(ctx)
    assert shape.get_text().startswith("Label: ")
    assert len(shape.range_calls) == 1
    start, end, _ = shape.range_calls[0]
    assert start == len("Label: ")  # only the tag span was replaced
    assert end == len("Label: #SUNRISE")


def test_tracked_shape_refresh_reports_changed_tag_names():
    # refresh returns the set of tag NAMES whose value actually changed -
    # used by PresentationController.refresh_content to decide on PDF
    # exports. A tag whose value is unchanged must NOT appear in the set.
    ctx_summer = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    ctx_winter = _ctx(datetime(2024, 12, 21, 12, 0, 0))
    shape = FakeShape("#SUNRISE")
    tracked = tagging.TrackedShape(shape=shape, template="#SUNRISE")
    # First refresh: raw "#SUNRISE" -> rendered => changed.
    assert tracked.refresh(ctx_summer) == {"SUNRISE"}
    # Same ctx again: no change.
    assert tracked.refresh(ctx_summer) == set()
    # Different ctx: the value changed again => reported.
    assert tracked.refresh(ctx_winter) == {"SUNRISE"}


def test_tracked_shape_refresh_reports_empty_when_no_tags_changed():
    shape = FakeShape("no tags here #SUNRISE")
    tracked = tagging.TrackedShape(shape=shape, template="no tags here #SUNRISE")
    assert tracked.refresh(_ctx(datetime(2024, 6, 21, 12, 0, 0))) == {"SUNRISE"}
    assert tracked.refresh(_ctx(datetime(2024, 6, 21, 12, 0, 0))) == set()


def test_is_clock_text_tag_true_only_for_dicl_clock():
    assert tagging.is_clock_text_tag("DICLOCK") is True
    assert tagging.is_clock_text_tag("SUNSET") is False
    assert tagging.is_clock_text_tag("CMD") is False


def test_tracked_shape_refresh_does_not_rewrite_unchanged_tag_value():
    # Calling refresh twice with the SAME effective ctx must not issue a
    # second set_text_range call for a tag whose rendered value didn't
    # change (avoids an unnecessary redraw/formatting churn, consistent
    # with this project's existing "skip redundant writes" convention).
    shape = RecordingFakeShape("#SUNRISE")
    tracked = tagging.TrackedShape(shape=shape, template="#SUNRISE")
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    tracked.refresh(ctx)
    tracked.refresh(ctx)
    assert len(shape.range_calls) == 1


class NudgeableFakeShape(RecordingFakeShape):
    """RecordingFakeShape that also tracks nudge() calls - see
    uno_shapes.UnoTextShape.nudge/_nudge_position's docstrings for why
    this exists (a LibreOffice slideshow repaint workaround for tagged
    text shapes on a slide with no analog clock)."""

    def __init__(self, text: str):
        super().__init__(text)
        self.nudge_calls = 0

    def nudge(self) -> None:
        self.nudge_calls += 1


def test_tracked_shape_refresh_nudges_shape_only_when_value_changed():
    shape = NudgeableFakeShape("#SUNRISE")
    tracked = tagging.TrackedShape(shape=shape, template="#SUNRISE")
    summer_ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    tracked.refresh(summer_ctx, nudge=True)
    assert shape.nudge_calls == 1
    # Same ctx again - value unchanged, must NOT nudge again.
    tracked.refresh(summer_ctx, nudge=True)
    assert shape.nudge_calls == 1
    # A different ctx whose rendered value differs - nudges again.
    winter_ctx = _ctx(datetime(2024, 12, 21, 12, 0, 0))
    tracked.refresh(winter_ctx, nudge=True)
    assert shape.nudge_calls == 2


def test_tracked_shape_refresh_never_nudges_when_nudge_false():
    shape = NudgeableFakeShape("#SUNRISE")
    tracked = tagging.TrackedShape(shape=shape, template="#SUNRISE")
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    tracked.refresh(ctx, nudge=False)
    assert shape.nudge_calls == 0


def test_tracked_shape_refresh_nudge_true_is_safe_without_nudge_method():
    # A shape (e.g. a plain test double) with no nudge() method at all
    # must not raise even when nudge=True - TrackedShape.refresh looks it
    # up via getattr(..., None) rather than assuming every TextShapeLike
    # implementation provides one.
    shape = RecordingFakeShape("#SUNRISE")
    tracked = tagging.TrackedShape(shape=shape, template="#SUNRISE")
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    tracked.refresh(ctx, nudge=True)  # must not raise
    assert shape.get_text() != "#SUNRISE"


def test_tracked_shape_refresh_handles_multiple_tags_with_varying_lengths():
    # Two tags whose rendered VALUE LENGTHS can legitimately differ across
    # refreshes - confirms the offset bookkeeping correctly accounts for
    # an earlier tag's length change shifting a later tag's position.
    shape = FakeShape("#SUNRISE to #SUNSET today")
    tracked = tagging.TrackedShape(shape=shape, template="#SUNRISE to #SUNSET today")
    tracked.refresh(_ctx(datetime(2024, 6, 21, 12, 0, 0)))
    summer_text = shape.get_text()
    assert summer_text.endswith(" today")
    assert " to " in summer_text
    tracked.refresh(_ctx(datetime(2024, 12, 21, 12, 0, 0)))
    winter_text = shape.get_text()
    assert winter_text.endswith(" today")
    assert " to " in winter_text
    assert summer_text != winter_text


def test_tracked_shape_refresh_applies_newline_format_suffix():
    # The live shape text initially equals the template VERBATIM, including
    # the literal ":n" suffix - confirms TrackedShape's first-refresh
    # bookkeeping (which seeds "current value" from the raw matched tag
    # text) correctly accounts for the suffix being part of that raw match.
    shape = FakeShape("Sunrise: #SUNRISE:n")
    tracked = tagging.TrackedShape(shape=shape, template="Sunrise: #SUNRISE:n")
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    tracked.refresh(ctx)
    base_value = tagging.render_template("#SUNRISE", ctx)
    assert shape.get_text() == f"Sunrise: \n{base_value}"
    # A second refresh with the same ctx must not re-issue a write either.
    tracked2 = tagging.TrackedShape(shape=RecordingFakeShape("#SUNRISE:n"), template="#SUNRISE:n")
    tracked2.refresh(ctx)
    tracked2.refresh(ctx)
    assert len(tracked2.shape.range_calls) == 1



def test_parsha_tag_produces_nonempty_hebrew_text():
    # Thursday before Parshas Shemot (Shabbos Jan 6 2024)
    ctx = _ctx(datetime(2024, 1, 4, 12, 0, 0))
    text = tagging.render_template("#PARSHA", ctx)
    assert text.strip()
    assert "שמות" in text


def test_dayinfo_tag_includes_yom_tov_title_on_non_parshah_weekday_yomtov():
    # Pesach day 1 2024 = 23 Apr 2024 (a Tuesday - not Shabbos, no parshah).
    ctx = _ctx(datetime(2024, 4, 23, 12, 0, 0))
    text = tagging.render_template("#DAYINFO", ctx)
    assert "פסח" in text


def test_dayinfo_tag_omits_yom_tov_title_on_parshah_shabbos():
    # Shabbos Jan 6 2024 (Parshas Shemot) - a parshah day, so #DAYINFO
    # must NOT also include a yom-tov/moed title line (that's #PARSHA's
    # job, not #DAYINFO's, per the VBA logic this mirrors). Note: other
    # #DAYINFO lines (e.g. Shabbos Mevorchim) may legitimately still
    # contain "שבת" - only the parshah NAME itself must be absent.
    ctx = _ctx(datetime(2024, 1, 6, 12, 0, 0))
    text = tagging.render_template("#DAYINFO", ctx)
    assert "שמות" not in text


def test_dafyomi_tag_matches_known_start_date():
    ctx = _ctx(datetime(1923, 9, 11, 12, 0, 0))
    text = tagging.render_template("#DAFYOMI", ctx)
    assert "ברכות" in text


def test_molad_tag_is_registered_and_known():
    assert "MOLAD" in tagging.TAG_REGISTRY
    assert tagging.contains_known_tag("#MOLAD")


def test_molad_tag_two_day_rosh_chodesh_matches_known_real_world_example():
    # 28 Tishrei 5787 (2026-10-09) - Tishrei always has 30 days, so the
    # upcoming Rosh Chodesh (Cheshvan) is a two-day one: 30 Tishrei
    # (Sunday) and 1 Cheshvan (Monday) - a known-good real-world example
    # independently verified against get_molad/molad_format output.
    ctx = _ctx(datetime(2026, 10, 9, 12, 0, 0))
    text = tagging.render_template("#MOLAD", ctx)
    assert text == (
        "ראש חודש חשון\n"
        "בימים ראשון ושני\n"
        "המולד ביום ראשון, שעה 09:43 ו-2 חלקים"
    )


def test_molad_tag_single_day_rosh_chodesh_has_no_second_day():
    from pyzmanim.hebrewcalendar import HDate, convert_date, hdate_add_day, hdate_gregorian, last_day_of_hebrew_month

    # Find a Hebrew month with exactly 29 days (a one-day Rosh Chodesh for
    # the FOLLOWING month) that isn't Elul (so the month after isn't
    # Tishrei/Rosh Hashana, which is handled - and tested - separately).
    hd = convert_date(datetime(2024, 1, 1, 12, 0, 0))
    for _ in range(600):
        if last_day_of_hebrew_month(hd.month, hd.year) == 29 and hd.month != 6:
            break
        hdate_add_day(hd, 1)
    else:
        raise AssertionError("could not find a qualifying 29-day month within range")

    ctx = _ctx(hdate_gregorian(hd))
    text = tagging.render_template("#MOLAD", ctx)
    lines = text.split("\n")
    assert len(lines) == 3
    assert lines[0].startswith("ראש חודש ")
    assert lines[1].startswith("ביום ")  # single day-of-week, no "ו" joining
    assert "בימים" not in lines[1]
    assert " ו" not in lines[1]
    assert lines[2].startswith("המולד ביום")


def test_molad_tag_two_day_rosh_chodesh_has_both_days_joined_by_vav():
    from pyzmanim.hebrewcalendar import HDate, convert_date, hdate_add_day, hdate_gregorian, last_day_of_hebrew_month

    # Find a 30-day Hebrew month (a two-day Rosh Chodesh for the
    # following month) that isn't Elul.
    hd = convert_date(datetime(2024, 1, 1, 12, 0, 0))
    for _ in range(600):
        if last_day_of_hebrew_month(hd.month, hd.year) == 30 and hd.month != 6:
            break
        hdate_add_day(hd, 1)
    else:
        raise AssertionError("could not find a qualifying 30-day month within range")

    ctx = _ctx(hdate_gregorian(hd))
    text = tagging.render_template("#MOLAD", ctx)
    lines = text.split("\n")
    assert len(lines) == 3
    assert lines[0].startswith("ראש חודש ")
    assert lines[1].startswith("בימים ")
    assert " ו" in lines[1]  # second day joined directly onto "ו" (no space)


def test_molad_tag_omits_rosh_chodesh_line_before_rosh_hashana():
    from pyzmanim.hebrewcalendar import convert_date, hdate_add_day, hdate_gregorian

    # Elul 5784 - the upcoming month is Tishrei (Rosh Hashana), which is
    # NOT Rosh Chodesh (no Birkas HaChodesh is said for it) - only the
    # molad line should be present, no "ראש חודש" line at all.
    hd = convert_date(datetime(2024, 1, 1, 12, 0, 0))
    while hd.month != 6:  # 6 = Elul
        hdate_add_day(hd, 1)
    ctx = _ctx(hdate_gregorian(hd))
    text = tagging.render_template("#MOLAD", ctx)
    lines = text.split("\n")
    assert len(lines) == 1
    assert "ראש חודש" not in text
    assert "ראש השנה" not in text  # not asserted BY this tag, just never fabricated
    assert lines[0].startswith("המולד ביום")


def test_molad_tag_molad_line_matches_get_molad_directly():
    from pyzmanim.hebrewcalendar import HDate, get_molad, hdate_add_month
    from pyzmanim.hdateformat import molad_format

    ctx = _ctx(datetime(2024, 1, 15, 12, 0, 0))  # 5 Shevat 5784
    hd_next = HDate(**ctx.now.__dict__)
    hdate_add_month(hd_next, 1)
    # The tag renders the molad with full_date=False (no Hebrew date - see
    # the "Birkas HaChodesh"-style format in tagging._tag_molad), and the
    # day-of-week here is the molad's own weekday, which can differ from
    # the Rosh Chodesh day(s)-of-week lines above.
    expected_molad_line = f"המולד ב{molad_format(get_molad(hd_next.year, hd_next.month), full_date=False)}"
    text = tagging.render_template("#MOLAD", ctx)
    assert expected_molad_line in text.split("\n")


def test_cmd_tag_pattern_captures_args():
    matches = list(tagging.TAG_PATTERN.finditer('#CMD:<-city "Beit Shemesh" -format json>'))
    assert len(matches) == 1
    assert matches[0].group(1) == "CMD"
    assert matches[0].group(2) is None
    assert matches[0].group(4) == '-city "Beit Shemesh" -format json'


def test_cmd_tag_pattern_with_no_args():
    matches = list(tagging.TAG_PATTERN.finditer("#CMD"))
    assert len(matches) == 1
    assert matches[0].group(1) == "CMD"
    assert matches[0].group(4) is None


def test_parse_shell_like_args_plain_whitespace():
    assert tagging._parse_shell_like_args("foo bar baz") == ["foo", "bar", "baz"]


def test_parse_shell_like_args_double_quoted_span():
    assert tagging._parse_shell_like_args('-city "Beit Shemesh" -format json') == [
        "-city",
        "Beit Shemesh",
        "-format",
        "json",
    ]


def test_parse_shell_like_args_single_quoted_span():
    assert tagging._parse_shell_like_args("'single quoted arg' second") == [
        "single quoted arg",
        "second",
    ]


def test_parse_shell_like_args_empty_string():
    assert tagging._parse_shell_like_args("") == []


def test_parse_shell_like_args_collapses_extra_whitespace():
    assert tagging._parse_shell_like_args("  foo    bar  ") == ["foo", "bar"]


def test_contains_known_tag_recognizes_cmd_tag():
    assert tagging.contains_known_tag("#CMD:<foo>")
    assert tagging.contains_known_tag("#CMD")


def test_find_first_config_json_no_match_returns_none():
    assert tagging.find_first_config_json(["no config here", "#HIDDEN"]) is None


def test_find_first_config_json_single_match():
    result = tagging.find_first_config_json(['#CONFIG:{"cmd_tick_seconds": 60}'])
    assert result == '{"cmd_tick_seconds": 60}'


def test_find_first_config_json_strips_whitespace():
    result = tagging.find_first_config_json(['#CONFIG:   {"a": 1}   '])
    assert result == '{"a": 1}'


def test_find_first_config_json_first_match_wins():
    texts = [
        "no tag here",
        '#CONFIG:{"first": true}',
        '#CONFIG:{"second": true}',
    ]
    result = tagging.find_first_config_json(texts)
    assert result == '{"first": true}'


def test_find_first_config_json_tag_not_at_start_of_text():
    result = tagging.find_first_config_json(['some label #CONFIG:{"a": 1}'])
    assert result == '{"a": 1}'


def test_find_first_config_json_empty_list_returns_none():
    assert tagging.find_first_config_json([]) is None


def test_cmd_tag_with_no_cache_renders_as_empty_string():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    text = tagging.render_template("before #CMD:<foo> after", ctx, cmd_cache=None)
    assert text == "before  after"


def test_cmd_tag_with_empty_cache_renders_as_empty_string():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    text = tagging.render_template("before #CMD:<foo> after", ctx, cmd_cache={})
    assert text == "before  after"


def test_cmd_tag_reads_from_cache_by_normalized_args_key():
    ctx = _ctx(datetime(2024, 6, 21, 12, 0, 0))
    cache = {"foo": "resolved value", "": "no-args value"}
    assert tagging.render_template("#CMD:<foo>", ctx, cmd_cache=cache) == "resolved value"
    assert tagging.render_template("#CMD", ctx, cmd_cache=cache) == "no-args value"
    assert tagging.render_template("#CMD:<>", ctx, cmd_cache=cache) == "no-args value"


def test_tracked_shape_refresh_passes_cmd_cache_through():
    shape = FakeShape("#CMD:<foo>")
    tracked = tagging.TrackedShape(shape=shape, template="#CMD:<foo>")
    tracked.refresh(_ctx(datetime(2024, 6, 21, 12, 0, 0)), cmd_cache={"foo": "cached!"})
    assert shape.get_text() == "cached!"


def test_find_cmd_arg_keys_collects_distinct_keys_across_templates():
    keys = tagging.find_cmd_arg_keys(["#CMD:<chol>", "text #CMD:<shabbos> more", "#CMD", "no tag here"])
    assert keys == {"chol", "shabbos", ""}


def test_find_cmd_arg_keys_ignores_non_cmd_tags():
    keys = tagging.find_cmd_arg_keys(["#SUNRISE-18", "#PARSHA"])
    assert keys == set()


def test_find_cmd_arg_keys_empty_when_no_templates():
    assert tagging.find_cmd_arg_keys([]) == set()


def test_refresh_cmd_cache_runs_configured_executable_and_substitutes_stdout():
    import sys

    cmd = f'"{sys.executable}" -c "import sys; print(sys.argv[1])"'
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, cmd, {"hello-world"})
    assert cache["hello-world"].strip() == "hello-world"


def test_refresh_cmd_cache_appends_tag_args_after_configured_prefix():
    import sys

    cmd = f'"{sys.executable}" -c "import sys; print(\' \'.join(sys.argv[1:]))"'
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, cmd, {'-a "b c" d'})
    assert cache['-a "b c" d'].strip() == "-a b c d"


def test_refresh_cmd_cache_global_args_applied_to_every_key():
    import sys

    cmd = f'"{sys.executable}" -c "import sys; print(\' \'.join(sys.argv[1:]))"'
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, cmd, {"one", "two"}, global_args="--shared-flag")
    assert cache["one"].strip() == "--shared-flag one"
    assert cache["two"].strip() == "--shared-flag two"


def test_refresh_cmd_cache_global_args_come_before_tag_args():
    import sys

    cmd = f'"{sys.executable}" -c "import sys; print(\' \'.join(sys.argv[1:]))"'
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, cmd, {"tag-arg"}, global_args='--global "g val"')
    assert cache["tag-arg"].strip() == "--global g val tag-arg"


def test_refresh_cmd_cache_empty_global_args_is_a_noop():
    import sys

    cmd = f'"{sys.executable}" -c "import sys; print(\' \'.join(sys.argv[1:]))"'
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, cmd, {"foo"}, global_args="")
    assert cache["foo"].strip() == "foo"


def test_refresh_cmd_cache_empty_list_global_args_is_a_noop():
    import sys

    cmd = f'"{sys.executable}" -c "import sys; print(\' \'.join(sys.argv[1:]))"'
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, cmd, {"foo"}, global_args=[])
    assert cache["foo"].strip() == "foo"


def test_refresh_cmd_cache_list_global_args_passed_verbatim_no_quote_stripping():
    import sys

    cmd = f'"{sys.executable}" -c "import sys; print(repr(sys.argv[1:]))"'
    cache: dict[str, str] = {}
    # A JSON-list global_args item containing a space must arrive as ONE
    # argv entry with the space intact - no shell-like splitting/quote
    # stripping applied to list items at all (unlike the string form).
    tagging.refresh_cmd_cache(cache, cmd, {"tag"}, global_args=["--sep", " - ", "--swap"])
    assert cache["tag"].strip() == repr(["--sep", " - ", "--swap", "tag"])


def test_normalize_global_args_none_and_empty_variants():
    assert tagging._normalize_global_args(None) == []
    assert tagging._normalize_global_args("") == []
    assert tagging._normalize_global_args([]) == []


def test_normalize_global_args_string_is_shell_parsed():
    assert tagging._normalize_global_args('--sep "a b"') == ["--sep", "a b"]


def test_normalize_global_args_list_is_used_verbatim():
    assert tagging._normalize_global_args(["--sep", " - ", "--swap"]) == ["--sep", " - ", "--swap"]


def test_normalize_global_args_list_does_not_strip_quote_characters():
    # Unlike the string form, a literal quote character inside a list item
    # is NOT stripped - it's part of the argv entry's actual value.
    assert tagging._normalize_global_args(["--sep=' - '"]) == ["--sep=' - '"]


def test_refresh_cmd_cache_resolves_multiple_keys_independently():
    import sys

    cmd = f'"{sys.executable}" -c "import sys; print(sys.argv[1])"'
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, cmd, {"one", "two"})
    assert cache["one"].strip() == "one"
    assert cache["two"].strip() == "two"


def test_refresh_cmd_cache_not_configured_stores_error_placeholder():
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, None, {"foo"})
    assert "#CMD" in cache["foo"]
    assert "not configured" in cache["foo"] or "no cmd_tag_executable" in cache["foo"]


def test_refresh_cmd_cache_missing_executable_stores_error_placeholder():
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, "definitely_not_a_real_executable_xyz", {""})
    assert "not found" in cache[""]


def test_refresh_cmd_cache_nonzero_exit_stores_error_placeholder():
    import sys

    cmd = f'"{sys.executable}" -c "import sys; sys.exit(3)"'
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, cmd, {""})
    assert "exit code 3" in cache[""]


def test_refresh_cmd_cache_timeout_stores_error_placeholder():
    import sys

    cmd = f'"{sys.executable}" -c "import time; time.sleep(5)"'
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, cmd, {""}, timeout_seconds=0.2)
    assert "timed out" in cache[""]


def test_refresh_cmd_cache_keeps_previous_good_value_on_failure():
    # A key that previously resolved successfully must NOT be overwritten
    # with an error placeholder if the NEXT run fails - the display
    # should keep showing the last known-good value instead.
    cache: dict[str, str] = {"foo": "previously resolved value"}
    tagging.refresh_cmd_cache(cache, "definitely_not_a_real_executable_xyz", {"foo"})
    assert cache["foo"] == "previously resolved value"


def test_refresh_cmd_cache_stores_error_if_no_previous_good_value_exists():
    # A key with NO previous value at all still gets the error placeholder
    # stored (so a persistently-broken config remains visibly diagnosable).
    cache: dict[str, str] = {}
    tagging.refresh_cmd_cache(cache, "definitely_not_a_real_executable_xyz", {"foo"})
    assert tagging._is_cmd_error_placeholder(cache["foo"])


def test_refresh_cmd_cache_overwrites_a_previous_error_with_a_new_error():
    # A key whose PREVIOUS value was ALREADY an error placeholder (not a
    # genuine prior success) should still get overwritten by a new error -
    # only a previously-GOOD value is protected from being overwritten.
    cache: dict[str, str] = {"foo": "[#CMD: exit code 1]"}
    tagging.refresh_cmd_cache(cache, "definitely_not_a_real_executable_xyz", {"foo"})
    assert cache["foo"] != "[#CMD: exit code 1]"
    assert tagging._is_cmd_error_placeholder(cache["foo"])


def test_refresh_cmd_cache_overwrites_previous_good_value_on_success():
    import sys

    cmd = f'"{sys.executable}" -c "print(\'fresh value\')"'
    cache: dict[str, str] = {"": "stale old value"}
    tagging.refresh_cmd_cache(cache, cmd, {""})
    assert cache[""] == "fresh value"


def test_is_cmd_error_placeholder():
    assert tagging._is_cmd_error_placeholder("[#CMD: timed out after 10s]")
    assert not tagging._is_cmd_error_placeholder("regular output")
    assert not tagging._is_cmd_error_placeholder("")


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
