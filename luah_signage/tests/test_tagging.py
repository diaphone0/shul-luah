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


def test_cmd_tag_pattern_captures_args():
    matches = list(tagging.TAG_PATTERN.finditer('#CMD:<-city "Beit Shemesh" -format json>'))
    assert len(matches) == 1
    assert matches[0].group(1) == "CMD"
    assert matches[0].group(2) is None
    assert matches[0].group(3) == '-city "Beit Shemesh" -format json'


def test_cmd_tag_pattern_with_no_args():
    matches = list(tagging.TAG_PATTERN.finditer("#CMD"))
    assert len(matches) == 1
    assert matches[0].group(1) == "CMD"
    assert matches[0].group(3) is None


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
