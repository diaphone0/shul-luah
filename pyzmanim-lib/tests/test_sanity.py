"""
Sanity tests for the pyzmanim port, checking against known real-world dates.

Run with: python -m pytest tests/ -v
(or: python tests/test_sanity.py to run without pytest)
"""
from __future__ import annotations

from datetime import datetime

from pyzmanim import dafyomi, hdateformat, shiur, zmanim
from pyzmanim.hebrewcalendar import (
    Parshah,
    YomTov,
    convert_date,
    get_parshah,
    get_yom_tov,
    hdate_gregorian,
)
from pyzmanim.noaa_calculator import Location

JERUSALEM = Location(latitude=31.7683, longitude=35.2137, elevation=754)


def test_gregorian_hebrew_round_trip():
    d = datetime(2024, 4, 23)
    h = convert_date(d)
    assert hdate_gregorian(h) == d


def test_pesach_day1():
    h = convert_date(datetime(2024, 4, 23))
    assert h.year == 5784
    assert h.month == 1
    assert h.day == 15
    assert get_yom_tov(h) == YomTov.PESACH_DAY1


def test_parshah_shemot():
    # 6 Jan 2024 was Shabbos Parshas Shemot
    h = convert_date(datetime(2024, 1, 6))
    assert get_parshah(h) == Parshah.SHEMOT


def test_sunrise_sunset_jerusalem_solstice():
    # Summer solstice 2024, Jerusalem: real sunrise ~05:34 IDT (UTC+3) = 02:34 UTC
    h = convert_date(datetime(2024, 6, 21))
    sunrise = hdate_gregorian(zmanim.getsunrise(h, JERUSALEM))
    sunset = hdate_gregorian(zmanim.getsunset(h, JERUSALEM))
    assert sunrise.hour == 2 and abs(sunrise.minute - 34) <= 2
    assert sunset.hour == 16 and abs(sunset.minute - 47) <= 2


def test_zmanim_with_fractional_shaah_zmanis_multipliers_do_not_crash():
    # Regression test: getminchagedolagra/getminchaketanagra/getplaggra use
    # non-integer shaah-zmanis multipliers (6.5, 9.5, 10.75) internally via
    # calctimeoffset, which previously left HDate.hour/min/sec as floats
    # (breaking hdate_gregorian's datetime() call) unless the offset is
    # rounded to an int before being applied.
    h = convert_date(datetime(2024, 6, 21))
    for func in (
        zmanim.getminchagedolagra,
        zmanim.getminchaketanagra,
        zmanim.getplaggra,
        zmanim.gettefilagra,
        zmanim.getshmagra,
    ):
        result = func(h, JERUSALEM)
        dt = hdate_gregorian(result)
        assert isinstance(dt.hour, int) and isinstance(dt.minute, int) and isinstance(dt.second, int)


def test_tzaisyeshiva_summer_uses_zmanis_minutes_longer_than_18():
    # Summer solstice: day > 12 hours (shaah zmanis > 60 real minutes), so
    # tzais should be MORE than a fixed 18 real minutes after sunset (18
    # "zmanis" minutes, proportionally longer).
    h = convert_date(datetime(2024, 6, 21))
    sunset = hdate_gregorian(zmanim.getsunset(h, JERUSALEM))
    tzais = hdate_gregorian(zmanim.gettzaisyeshiva(h, JERUSALEM))
    delta_minutes = (tzais - sunset).total_seconds() / 60
    assert delta_minutes > 18


def test_tzaisyeshiva_winter_uses_fixed_18_minutes():
    # Winter solstice: day <= 12 hours (shaah zmanis <= 60 real minutes),
    # so tzais should be EXACTLY 18 fixed real minutes after sunset.
    h = convert_date(datetime(2024, 12, 21))
    sunset = hdate_gregorian(zmanim.getsunset(h, JERUSALEM))
    tzais = hdate_gregorian(zmanim.gettzaisyeshiva(h, JERUSALEM))
    delta_minutes = (tzais - sunset).total_seconds() / 60
    assert abs(delta_minutes - 18) < 0.01


def test_tzaisyeshiva_elevation_adjusted_is_later_than_flat():
    # adjust_for_elevation=True should use the (later) elevation-adjusted
    # sunset as its base instead of the flat/level one, making the
    # resulting tzais later too. Default (False) must be unchanged.
    h = convert_date(datetime(2024, 6, 21))
    flat = hdate_gregorian(zmanim.gettzaisyeshiva(h, JERUSALEM))
    flat_explicit = hdate_gregorian(zmanim.gettzaisyeshiva(h, JERUSALEM, False))
    elevated = hdate_gregorian(zmanim.gettzaisyeshiva(h, JERUSALEM, True))
    assert flat == flat_explicit
    assert elevated > flat


def test_num_to_h_char_gematria():
    assert hdateformat.num_to_h_char(5784) == "תשפ\u05f4ד"
    assert hdateformat.num_to_h_char(15) == "ט\u05f4ו"  # avoids יה (divine name)
    assert hdateformat.num_to_h_char(16) == "ט\u05f4ז"  # avoids יו (divine name)


def test_daf_yomi_start_date_is_berachos_2():
    result = dafyomi.get_daf_yomi_bavli(dafyomi.DAF_YOMI_START_DATE)
    assert result.masechta_number == 0  # Berachos
    assert result.page == 2


def test_daf_yomi_before_start_returns_no_daf():
    from datetime import date
    result = dafyomi.get_daf_yomi_bavli(date(1900, 1, 1))
    assert result.page == -1


def test_chumash_and_tehillim_do_not_crash():
    h = convert_date(datetime(2024, 1, 6))
    assert shiur.chumash(h)
    assert shiur.tehillim(h)
    assert shiur.get_tanya(h)
    assert shiur.get_halacha(h)
    assert shiur.get_rambam(h, True)
    assert shiur.get_rambam(h, False)


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
