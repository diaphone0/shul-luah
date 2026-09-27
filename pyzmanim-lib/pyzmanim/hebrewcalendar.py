"""
Core Hebrew calendar engine.

Ported from mod_hebrewcalendar.bas in https://github.com/diaphone1/vbzmanim
(itself ported from yparitcher/libzmanim, a port of KosherJava/Zmanim by
Eliyahu Hershfeld, kosherjava.com).

Provides the `HDate` Hebrew-date type, leap year math, Hebrew<->Gregorian
conversion, holiday/Torah-portion/molad (new moon) determination, and hdate
arithmetic.

Note: the original VBA used a Windows-only `TMStruct`/`mktm`/`mkdate` layer
(mod_TMStruct.bas) to bridge `Date` and Hebrew-date math, relying on Win32
`GetTimeZoneInformation` for DST. That layer is not ported - Python's
`datetime` already provides normalized year/month/day/hour/minute/second
fields, and timezone/DST concerns are left to the caller (e.g. via
`zoneinfo`), matching the offset-in-seconds field already present on `HDate`.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import datetime
from enum import IntEnum


class YomTov(IntEnum):
    CHOL = 0
    PESACH_DAY1 = 1
    PESACH_DAY2 = 2
    SHVEI_SHEL_PESACH = 3
    ACHRON_SHEL_PESACH = 4
    SHAVOUS_DAY1 = 5
    SHAVOUS_DAY2 = 6
    ROSH_HASHANAH_DAY1 = 7
    ROSH_HASHANAH_DAY2 = 8
    YOM_KIPPUR = 9
    SUKKOS_DAY1 = 10
    SUKKOS_DAY2 = 11
    SHMEINI_ATZERES = 12
    SIMCHAS_TORAH = 13
    CHOL_HAMOED_PESACH_DAY1 = 14
    CHOL_HAMOED_PESACH_DAY2 = 15
    CHOL_HAMOED_PESACH_DAY3 = 16
    CHOL_HAMOED_PESACH_DAY4 = 17
    CHOL_HAMOED_PESACH_DAY5 = 18
    CHOL_HAMOED_SUKKOS_DAY1 = 19
    CHOL_HAMOED_SUKKOS_DAY2 = 20
    CHOL_HAMOED_SUKKOS_DAY3 = 21
    CHOL_HAMOED_SUKKOS_DAY4 = 22
    CHOL_HAMOED_SUKKOS_DAY5 = 23
    HOSHANA_RABBAH = 24
    PESACH_SHEINI = 25
    LAG_BAOMER = 26
    TU_BAV = 27
    CHANUKAH_DAY1 = 28
    CHANUKAH_DAY2 = 29
    CHANUKAH_DAY3 = 30
    CHANUKAH_DAY4 = 31
    CHANUKAH_DAY5 = 32
    CHANUKAH_DAY6 = 33
    CHANUKAH_DAY7 = 34
    CHANUKAH_DAY8 = 35
    TU_BISHVAT = 36
    PURIM_KATAN = 37
    SHUSHAN_PURIM_KATAN = 38
    PURIM = 39
    SHUSHAN_PURIM = 40
    SHIVA_ASAR_BTAAMUZ = 41
    TISHA_BAV = 42
    TZOM_GEDALIA = 43
    ASARAH_BTEVES = 44
    TAANIS_ESTER = 45
    EREV_PESACH = 46
    EREV_SHAVOUS = 47
    EREV_ROSH_HASHANAH = 48
    EREV_YOM_KIPPUR = 49
    EREV_SUKKOS = 50
    SHKALIM = 51
    ZACHOR = 52
    PARAH = 53
    HACHODESH = 54
    ROSH_CHODESH = 55
    MACHAR_CHODESH = 56
    SHABBOS_MEVORCHIM = 57
    HAGADOL = 58
    CHAZON = 59
    NACHAMU = 60
    SHUVA = 61
    SHIRA = 62
    SHABBOS_CHOL_HAMOED = 63


class Parshah(IntEnum):
    NOPARSHAH = 0
    BERESHIT = 1
    NOACH = 2
    LECH_LECHA = 3
    VAYEIRA = 4
    CHAYEI_SARAH = 5
    TOLEDOT = 6
    VAYETZE = 7
    VAYISHLACH = 8
    VAYESHEV = 9
    MIKETZ = 10
    VAYIGASH = 11
    VAYECHI = 12
    SHEMOT = 13
    VAEIRA = 14
    BO = 15
    BESHALACH = 16
    YITRO = 17
    MISHPATIM = 18
    TERUMAH = 19
    TETZAVEH = 20
    KI_TISA = 21
    VAYAKHEL = 22
    PEKUDEI = 23
    VAYIKRA = 24
    TZAV = 25
    SHEMINI = 26
    TAZRIA = 27
    METZORA = 28
    ACHAREI_MOT = 29
    KEDOSHIM = 30
    EMOR = 31
    BEHAR = 32
    BECHUKOTAI = 33
    BAMIDBAR = 34
    NASO = 35
    BEHAALOTECHA = 36
    SHLACH = 37
    KORACH = 38
    CHUKAT = 39
    BALAK = 40
    PINCHAS = 41
    MATOT = 42
    MASEI = 43
    DEVARIM = 44
    VAETCHANAN = 45
    EIKEV = 46
    REEH = 47
    SHOFTIM = 48
    KI_TEITZEI = 49
    KI_TAVO = 50
    NITZAVIM = 51
    VAYELECH = 52
    HAAZINU = 53
    VZOT_HABERACHAH = 54
    VAYAKHEL_PEKUDEI = 55
    TAZRIA_METZORA = 56
    ACHAREI_MOT_KEDOSHIM = 57
    BEHAR_BECHUKOTAI = 58
    CHUKAT_BALAK = 59
    MATOT_MASEI = 60
    NITZAVIM_VAYELECH = 61


@dataclass
class HDate:
    """Hebrew date. Mutated in place by the HDateAdd* functions below,
    mirroring the original VBA `ByRef` semantics."""
    year: int = 0
    month: int = 0  # 1=Nissan .. 13=Adar II (leap years only)
    day: int = 0  # 1-30
    dayOfYear: int = 0  # from Tishrei
    wday: int = 0  # 0=Sunday .. 6=Saturday
    leap: int = 0
    hour: int = 0
    min: int = 0
    sec: int = 0
    msec: int = 0
    offset: int = 0  # timezone offset in seconds
    EY: bool = False  # Eretz Yisroel (affects yomtov & parshah)


EMPTY_HDATE = HDate()


def hebrew_leap_year(year: int) -> bool:
    return ((7 * year + 1) % 19) < 7


def hebrew_calendar_elapsed_days(year: int) -> int:
    """Days from Molad Tohu until Rosh Hashana of `year`, with dechiyot
    (postponement rules) applied."""
    months_elapsed = (
        (235 * ((year - 1) // 19))
        + (12 * ((year - 1) % 19))
        + ((7 * ((year - 1) % 19) + 1) // 19)
    )
    parts_elapsed = 204 + 793 * (months_elapsed % 1080)
    hours_elapsed = (
        5 + 12 * months_elapsed + 793 * (months_elapsed // 1080) + (parts_elapsed // 1080)
    )
    conjunction_day = 1 + 29 * months_elapsed + (hours_elapsed // 24)
    conjunction_parts = 1080 * (hours_elapsed % 24) + parts_elapsed % 1080
    cdw = conjunction_day % 7

    if (
        conjunction_parts >= 19440
        or (cdw == 2 and conjunction_parts >= 9924 and not hebrew_leap_year(year))
        or (cdw == 1 and conjunction_parts >= 16789 and hebrew_leap_year(year - 1))
    ):
        alternative_day = conjunction_day + 1
    else:
        alternative_day = conjunction_day

    adw = alternative_day % 7
    if adw in (0, 3, 5):
        return 1 + alternative_day
    return alternative_day


def days_in_hebrew_year(year: int) -> int:
    return hebrew_calendar_elapsed_days(year + 1) - hebrew_calendar_elapsed_days(year)


def long_heshvan(year: int) -> bool:
    return (days_in_hebrew_year(year) % 10) == 5


def short_kislev(year: int) -> bool:
    return (days_in_hebrew_year(year) % 10) == 3


def last_day_of_hebrew_month(month: int, year: int) -> int:
    if (
        month in (2, 4, 6)
        or (month == 8 and not long_heshvan(year))
        or (month == 9 and short_kislev(year))
        or month == 10
        or (month == 12 and not hebrew_leap_year(year))
        or month == 13
    ):
        return 29
    return 30


def nissan_count(year: int) -> int:
    """Day-of-year (from Tishrei) of 1 Nissan."""
    return {
        353: 176, 354: 177, 355: 178,
        383: 206, 384: 207, 385: 208,
    }[days_in_hebrew_year(year)]


def hdate_new(
    year: int = 0, month: int = 0, day: int = 0,
    hour: int = 0, min: int = 0, sec: int = 0,
    msec: int = 0, offset: int = 0,
) -> HDate:
    return HDate(year=year, month=month, day=day, hour=hour, min=min, sec=sec, msec=msec, offset=offset)


def convert_date(date_in: datetime) -> HDate:
    """Convert a Gregorian `datetime` to an `HDate`."""
    result = HDate()

    julian_day = gregorian_julian(date_in.year, date_in.month, date_in.day)
    d = math.floor(julian_day) - 347996
    m = (d * 25920.0) / 765433.0
    year = int((19 * m) / 235)

    while d >= hebrew_calendar_elapsed_days(year + 1):
        year += 1

    ys = hebrew_calendar_elapsed_days(year)
    day_of_year = (d - ys) + 1
    nissan_start = nissan_count(year)

    if day_of_year <= nissan_start:
        month = 7  # Start at Tishrei
        daycount = 0
    else:
        month = 1  # Start at Nissan
        daycount = nissan_start

    while day_of_year > (daycount + last_day_of_hebrew_month(month, year)):
        daycount += last_day_of_hebrew_month(month, year)
        month += 1

    day = day_of_year - daycount

    result.year = year
    result.month = month
    result.day = day
    result.wday = (hebrew_calendar_elapsed_days(year) + day_of_year) % 7
    result.dayOfYear = day_of_year
    result.leap = 1 if hebrew_leap_year(year) else 0
    result.hour = date_in.hour
    result.min = date_in.minute
    result.sec = date_in.second
    return result


def hdate_gregorian(date_in: HDate) -> datetime:
    """Convert an `HDate` to a Gregorian `datetime`."""
    jd = hdate_julian(date_in) + 0.5
    a = math.floor((jd - 1867216.25) / 36524.25)
    b = jd + 1525 + a - math.floor(a / 4)
    c = math.floor((b - 122.1) / 365.25)
    dd = math.floor(c * 365.25)
    e = math.floor((b - dd) / 30.6001)

    month = int(e - 13) if e > 13 else int(e - 1)
    year = int(c - 4716) if month > 2 else int(c - 4715)
    day = int(b - dd - math.floor(e * 30.6001))

    return datetime(year, month, day, date_in.hour, date_in.min, date_in.sec)


def gregorian_julian(year: int, month: int, day: int) -> float:
    """Convert a Gregorian date to a Julian day number."""
    if month <= 2:
        year -= 1
        month += 12
    a = year // 100
    b = 2 - a + a // 4
    return math.floor(365.25 * (year + 4716)) + math.floor(30.6001 * (month + 1)) + day + b - 1524.5


def hdate_julian(date_in: HDate) -> float:
    diff = 347996.5
    yearstart = hebrew_calendar_elapsed_days(date_in.year)
    return (date_in.dayOfYear - 1) + yearstart + diff


def hdate_time_t(date_in: HDate) -> float:
    result = (hebrew_calendar_elapsed_days(date_in.year) + (date_in.dayOfYear - 1)) - 2092591
    result = ((((result * 24) + date_in.hour) * 60 + date_in.min) * 60) + date_in.sec
    return result - date_in.offset


def time_thdate(time: float, offset: int) -> HDate:
    temp = time + offset
    result = HDate()
    result.sec = int(temp % 60)
    temp = temp / 60
    result.min = int(temp % 60)
    temp = temp / 60
    result.hour = int(temp % 24)
    d = int(temp / 24) + 2092591
    m = (d * 25920.0) / 765433.0
    year = int((19.0 * m) / 235.0)

    while d >= hebrew_calendar_elapsed_days(year + 1):
        year += 1

    ys = hebrew_calendar_elapsed_days(year)
    day_of_year = (d - ys) + 1
    nissan_start = nissan_count(year)
    if day_of_year <= nissan_start:
        month = 7
        daycount = 0
    else:
        month = 1
        daycount = nissan_start
    while day_of_year > (daycount + last_day_of_hebrew_month(month, year)):
        daycount += last_day_of_hebrew_month(month, year)
        month += 1
    day = day_of_year - daycount

    result.year = year
    result.month = month
    result.day = day
    result.offset = offset
    hdate_set_doy(result)
    return result


def hdate_compare(date1: HDate, date2: HDate) -> int:
    """0 if equal, 1 if date1 < date2, -1 if date1 > date2."""
    if date1.year != date2.year:
        return 1 if date1.year < date2.year else -1
    if date1.dayOfYear != date2.dayOfYear:
        return 1 if date1.dayOfYear < date2.dayOfYear else -1
    if date1.hour != date2.hour:
        return 1 if date1.hour < date2.hour else -1
    if date1.min != date2.min:
        return 1 if date1.min < date2.min else -1
    if date1.sec != date2.sec:
        return 1 if date1.sec < date2.sec else -1
    if date1.msec != date2.msec:
        return 1 if date1.msec < date2.msec else -1
    return 0


def hdate_set_doy(date_in: HDate) -> None:
    """Normalize an hdate in place and set wday, dayOfYear, and leap.
    Converts month 13 (Adar II) in a non-leap year to 12; converts day 30 of
    Cheshvan/Kislev/Adar to 29 in a year where that month only has 29 days."""
    year = date_in.year
    month = date_in.month
    if month == 13 and not hebrew_leap_year(year):
        month = 12
    if date_in.day == 30 and last_day_of_hebrew_month(month, year) == 29:
        date_in.day = 29
    day = date_in.day
    if month < 7:
        monthcount = 1
        day_of_year = nissan_count(year)
    else:
        monthcount = 7
        day_of_year = 0
    while monthcount < month:
        day_of_year += last_day_of_hebrew_month(monthcount, year)
        monthcount += 1
    day_of_year += day
    date_in.dayOfYear = day_of_year
    date_in.wday = (hebrew_calendar_elapsed_days(year) + day_of_year) % 7
    date_in.leap = 1 if hebrew_leap_year(year) else 0


def hdate_add_year(date_in: HDate, years: int) -> None:
    year = date_in.year
    month = date_in.month
    leap1 = date_in.leap
    year += years
    leap2 = 1 if hebrew_leap_year(year) else 0
    if leap1 != leap2:
        if leap1 and month == 13:
            month = 12
        elif not leap1 and month == 12:
            month = 13
    date_in.year = year
    date_in.month = month
    hdate_set_doy(date_in)


def hdate_add_month(date_in: HDate, months: int) -> None:
    last = date_in.day == 30
    monthcount = months
    while monthcount > 0:
        if date_in.month == 12:
            date_in.month = date_in.month + 1 if date_in.leap == 1 else 1
            monthcount -= 1
        elif date_in.month == 13:
            date_in.month = 1
            monthcount -= 1
        elif date_in.month == 6:
            date_in.month += 1
            hdate_add_year(date_in, 1)
            monthcount -= 1
        else:
            date_in.month += 1
            monthcount -= 1
    while monthcount < 0:
        if date_in.month == 1:
            date_in.month = 13 if date_in.leap == 1 else 12
            monthcount += 1
        elif date_in.month == 7:
            date_in.month -= 1
            hdate_add_year(date_in, -1)
            monthcount += 1
        else:
            date_in.month -= 1
            monthcount += 1
    if last and last_day_of_hebrew_month(date_in.month, date_in.year) == 30:
        date_in.day = 30
    hdate_set_doy(date_in)


def hdate_add_day(date_in: HDate, days: int) -> None:
    daycount = days
    while daycount > 0:
        if date_in.day == 30:
            date_in.day = 1
            hdate_add_month(date_in, 1)
            daycount -= 1
        elif date_in.day == 29:
            if last_day_of_hebrew_month(date_in.month, date_in.year) == 29:
                date_in.day = 1
                hdate_add_month(date_in, 1)
            else:
                date_in.day += 1
            daycount -= 1
        else:
            date_in.day += 1
            daycount -= 1
    while daycount < 0:
        if date_in.day == 1:
            hdate_add_month(date_in, -1)
            date_in.day = 30 if last_day_of_hebrew_month(date_in.month, date_in.year) == 30 else 29
            daycount += 1
        else:
            date_in.day -= 1
            daycount += 1
    hdate_set_doy(date_in)


def divide_and_carry(start: int, divisor: int) -> tuple[int, int]:
    """Returns (finish, carry)."""
    finish = start % divisor
    carry = start // divisor
    if finish < 0:
        finish += divisor
        carry -= 1
    return finish, carry


def hdate_add_hour(date_in: HDate, hours: int) -> None:
    hour = date_in.hour + hours
    date_in.hour, carry = divide_and_carry(hour, 24)
    if carry:
        hdate_add_day(date_in, carry)
    else:
        hdate_set_doy(date_in)


def hdate_add_minute(date_in: HDate, minutes: int) -> None:
    minute = date_in.min + minutes
    date_in.min, carry = divide_and_carry(minute, 60)
    if carry:
        hdate_add_hour(date_in, carry)
    else:
        hdate_set_doy(date_in)


def hdate_add_second(date_in: HDate, seconds: int) -> None:
    second = date_in.sec + seconds
    date_in.sec, carry = divide_and_carry(second, 60)
    if carry:
        hdate_add_minute(date_in, carry)
    else:
        hdate_set_doy(date_in)


def hdate_add_msecond(date_in: HDate, mseconds: int) -> None:
    msecond = date_in.msec + mseconds
    date_in.msec, carry = divide_and_carry(msecond, 1000)
    if carry:
        hdate_add_second(date_in, carry)
    else:
        hdate_set_doy(date_in)


def hdate_add(
    date_in: HDate, years: int = 0, months: int = 0, days: int = 0,
    hours: int = 0, minutes: int = 0, seconds: int = 0, mseconds: int = 0,
) -> None:
    if years:
        hdate_add_year(date_in, years)
    if months:
        hdate_add_month(date_in, months)
    if days:
        hdate_add_day(date_in, days)
    if hours:
        hdate_add_hour(date_in, hours)
    if minutes:
        hdate_add_minute(date_in, minutes)
    if seconds:
        hdate_add_second(date_in, seconds)
    if mseconds:
        hdate_add_msecond(date_in, mseconds)


def get_molad(year: int, month: int) -> HDate:
    """Molad (new moon) of the given month (month 1 = Nissan), in Yerushalayim
    Mean Time. The returned hour/min/sec represent chalakim (1/1080 hour
    units), not standard time - see calcmoladoffset in zmanim.py."""
    result = hdate_new()
    months_elapsed = (
        (235 * ((year - 1) // 19))
        + (12 * ((year - 1) % 19))
        + (7 * ((year - 1) % 19) + 1) // 19
    )

    if month > 6:
        months_elapsed += month - 7
    else:
        months_elapsed += month + 5
        if hebrew_leap_year(year):
            months_elapsed += 1

    parts_elapsed = 204 + 793 * (months_elapsed % 1080)
    hours_elapsed = (
        5 + 12 * months_elapsed + 793 * (months_elapsed // 1080) + parts_elapsed // 1080
    )

    conjunction_day = 29 * months_elapsed + hours_elapsed // 24
    conjunction_hour = hours_elapsed % 24
    conjunction_minute = (parts_elapsed % 1080) // 18
    conjunction_parts = (parts_elapsed % 1080) % 18

    if conjunction_day > 2085362:
        # snapshot for shorter calculation - 1 Nissan 5710
        result.year = 5710
        result.month = 1
        result.day = 1
        result.wday = 1
        hdate_add_day(result, conjunction_day - 2085362)
    else:
        result.year = 1
        result.month = 7
        result.day = 1
        hdate_add_day(result, conjunction_day)

    result.hour = conjunction_hour
    result.min = conjunction_minute
    result.sec = conjunction_parts
    result.offset = 8456
    hdate_add_hour(result, -6)

    return result


def get_year_type(date_in: HDate) -> int:
    """Returns 0-16 index representing one of the 16 possible Hebrew year
    'shapes' (leap/non-leap x year-start weekday x Cheshvan/Kislev length),
    or -1. Used to index into parasha.PARASHA_LIST."""
    year_wday = (hebrew_calendar_elapsed_days(date_in.year) + 1) % 7
    if year_wday == 0:
        year_wday = 7

    if date_in.leap == 1:
        if year_wday == 2:
            if short_kislev(date_in.year):
                return 14 if date_in.EY else 6
            elif long_heshvan(date_in.year):
                return 15 if date_in.EY else 7
        elif year_wday == 3:
            return 15 if date_in.EY else 7
        elif year_wday == 5:
            if short_kislev(date_in.year):
                return 8
            elif long_heshvan(date_in.year):
                return 9
        elif year_wday == 7:
            if short_kislev(date_in.year):
                return 10
            elif long_heshvan(date_in.year):
                return 16 if date_in.EY else 11
    else:
        if year_wday == 2:
            if short_kislev(date_in.year):
                return 0
            elif long_heshvan(date_in.year):
                return 12 if date_in.EY else 1
        elif year_wday == 3:
            return 12 if date_in.EY else 1
        elif year_wday == 5:
            if long_heshvan(date_in.year):
                return 3
            elif not short_kislev(date_in.year):
                return 13 if date_in.EY else 2
        elif year_wday == 7:
            if short_kislev(date_in.year):
                return 4
            elif long_heshvan(date_in.year):
                return 5

    return -1


def get_parshah(date_in: HDate) -> Parshah:
    """If Shabbos, returns the current parshah; otherwise NOPARSHAH."""
    from pyzmanim import parasha  # local import: avoids circular import at module load time

    year_type = get_year_type(date_in)
    year_wday = hebrew_calendar_elapsed_days(date_in.year) % 7
    hdate_set_doy(date_in)
    day = year_wday + date_in.dayOfYear

    if date_in.wday != 0:
        return Parshah.NOPARSHAH

    if year_type >= 0:
        return parasha.PARASHA_LIST[year_type][day // 7]
    return Parshah.NOPARSHAH


def get_yom_tov(date_in: HDate) -> YomTov:
    """If yomtov, returns the current yomtov; otherwise CHOL."""
    result = YomTov.CHOL

    if date_in.month == 1:
        if date_in.day == 14:
            result = YomTov.EREV_PESACH
        elif date_in.day == 15:
            result = YomTov.PESACH_DAY1
        elif date_in.day == 16:
            result = YomTov.CHOL_HAMOED_PESACH_DAY1 if date_in.EY else YomTov.PESACH_DAY2
        elif date_in.day == 17:
            result = YomTov.CHOL_HAMOED_PESACH_DAY2 if date_in.EY else YomTov.CHOL_HAMOED_PESACH_DAY1
        elif date_in.day == 18:
            result = YomTov.CHOL_HAMOED_PESACH_DAY3 if date_in.EY else YomTov.CHOL_HAMOED_PESACH_DAY2
        elif date_in.day == 19:
            result = YomTov.CHOL_HAMOED_PESACH_DAY4 if date_in.EY else YomTov.CHOL_HAMOED_PESACH_DAY3
        elif date_in.day == 20:
            result = YomTov.CHOL_HAMOED_PESACH_DAY5 if date_in.EY else YomTov.CHOL_HAMOED_PESACH_DAY4
        elif date_in.day == 21:
            result = YomTov.SHVEI_SHEL_PESACH
        elif date_in.day == 22 and not date_in.EY:
            result = YomTov.ACHRON_SHEL_PESACH
    elif date_in.month == 2:
        if date_in.day == 14:
            return YomTov.PESACH_SHEINI
        if date_in.day == 18:
            return YomTov.LAG_BAOMER
    elif date_in.month == 3:
        if date_in.day == 5:
            return YomTov.EREV_SHAVOUS
        if date_in.day == 6:
            return YomTov.SHAVOUS_DAY1
        if date_in.day == 7 and not date_in.EY:
            return YomTov.SHAVOUS_DAY2
    elif date_in.month == 4:
        if date_in.day in (17, 18):
            if (date_in.day == 17 and date_in.wday != 0) or (date_in.day == 18 and date_in.wday == 1):
                return YomTov.SHIVA_ASAR_BTAAMUZ
    elif date_in.month == 5:
        if date_in.day in (9, 10):
            if (date_in.day == 9 and date_in.wday != 0) or (date_in.day == 10 and date_in.wday == 1):
                return YomTov.TISHA_BAV
        elif date_in.day == 15:
            return YomTov.TU_BAV
    elif date_in.month == 6:
        if date_in.day == 29:
            return YomTov.EREV_ROSH_HASHANAH
    elif date_in.month == 7:
        if date_in.day == 1:
            result = YomTov.ROSH_HASHANAH_DAY1
        elif date_in.day == 2:
            result = YomTov.ROSH_HASHANAH_DAY2
        elif date_in.day == 3:
            if date_in.wday != 0:
                result = YomTov.TZOM_GEDALIA
        elif date_in.day == 4:
            if date_in.wday == 1:
                result = YomTov.TZOM_GEDALIA
        elif date_in.day == 9:
            result = YomTov.EREV_YOM_KIPPUR
        elif date_in.day == 10:
            result = YomTov.YOM_KIPPUR
        elif date_in.day == 14:
            result = YomTov.EREV_SUKKOS
        elif date_in.day == 15:
            result = YomTov.SUKKOS_DAY1
        elif date_in.day == 16:
            result = YomTov.CHOL_HAMOED_SUKKOS_DAY1 if date_in.EY else YomTov.SUKKOS_DAY2
        elif date_in.day == 17:
            result = YomTov.CHOL_HAMOED_SUKKOS_DAY2 if date_in.EY else YomTov.CHOL_HAMOED_SUKKOS_DAY1
        elif date_in.day == 18:
            result = YomTov.CHOL_HAMOED_SUKKOS_DAY3 if date_in.EY else YomTov.CHOL_HAMOED_SUKKOS_DAY2
        elif date_in.day == 19:
            result = YomTov.CHOL_HAMOED_SUKKOS_DAY4 if date_in.EY else YomTov.CHOL_HAMOED_SUKKOS_DAY3
        elif date_in.day == 20:
            result = YomTov.CHOL_HAMOED_SUKKOS_DAY5 if date_in.EY else YomTov.CHOL_HAMOED_SUKKOS_DAY4
        elif date_in.day == 21:
            result = YomTov.HOSHANA_RABBAH
        elif date_in.day == 22:
            result = YomTov.SHMEINI_ATZERES
        elif date_in.day == 23 and not date_in.EY:
            result = YomTov.SIMCHAS_TORAH
    elif date_in.month == 9:
        if date_in.day == 25:
            result = YomTov.CHANUKAH_DAY1
        elif date_in.day == 26:
            result = YomTov.CHANUKAH_DAY2
        elif date_in.day == 27:
            result = YomTov.CHANUKAH_DAY3
        elif date_in.day == 28:
            result = YomTov.CHANUKAH_DAY4
        elif date_in.day == 29:
            result = YomTov.CHANUKAH_DAY5
        elif date_in.day == 30:
            result = YomTov.CHANUKAH_DAY6
    elif date_in.month == 10:
        if date_in.day == 1:
            return YomTov.CHANUKAH_DAY6 if short_kislev(date_in.year) else YomTov.CHANUKAH_DAY7
        if date_in.day == 2:
            return YomTov.CHANUKAH_DAY7 if short_kislev(date_in.year) else YomTov.CHANUKAH_DAY8
        if date_in.day == 3:
            if short_kislev(date_in.year):
                return YomTov.CHANUKAH_DAY8
        elif date_in.day == 10:
            return YomTov.ASARAH_BTEVES
    elif date_in.month == 11:
        if date_in.day == 15:
            return YomTov.TU_BISHVAT
    elif date_in.month == 12:
        if date_in.day == 11:
            if date_in.leap == 0 and date_in.wday == 5:
                return YomTov.TAANIS_ESTER
        elif date_in.day == 13:
            if date_in.leap == 0 and date_in.wday != 0:
                return YomTov.TAANIS_ESTER
        elif date_in.day == 14:
            return YomTov.PURIM_KATAN if date_in.leap == 1 else YomTov.PURIM
        elif date_in.day == 15:
            return YomTov.SHUSHAN_PURIM_KATAN if date_in.leap == 1 else YomTov.SHUSHAN_PURIM
    elif date_in.month == 13:
        if date_in.day == 11:
            if date_in.wday == 5:
                return YomTov.TAANIS_ESTER
        elif date_in.day == 13:
            if date_in.wday != 0:
                return YomTov.TAANIS_ESTER
        elif date_in.day == 14:
            return YomTov.PURIM
        elif date_in.day == 15:
            return YomTov.SHUSHAN_PURIM

    if (
        YomTov.CHOL_HAMOED_PESACH_DAY1 <= result <= YomTov.CHOL_HAMOED_PESACH_DAY5
        or YomTov.CHOL_HAMOED_SUKKOS_DAY1 <= result <= YomTov.CHOL_HAMOED_SUKKOS_DAY5
    ):
        if date_in.wday == 0:
            result = YomTov.SHABBOS_CHOL_HAMOED

    return result


def get_special_shabbos(date_in: HDate) -> YomTov:
    """If Shabbos, returns the current special parshah; otherwise CHOL."""
    result = YomTov.CHOL
    if date_in.wday == 0:
        if (date_in.month == 11 and date_in.leap != 1) or (date_in.month == 12 and date_in.leap == 1):
            if date_in.day in (25, 27, 29):
                result = YomTov.SHKALIM
        if (date_in.month == 12 and date_in.leap != 1) or date_in.month == 13:
            if date_in.day == 1:
                result = YomTov.SHKALIM
            elif date_in.day in (8, 9, 11, 13):
                result = YomTov.ZACHOR
            elif date_in.day in (18, 20, 22, 23):
                result = YomTov.PARAH
            elif date_in.day in (25, 27, 29):
                result = YomTov.HACHODESH
        if date_in.month == 1:
            if date_in.day == 1:
                result = YomTov.HACHODESH
            if 8 <= date_in.day <= 14:
                result = YomTov.HAGADOL
        if date_in.month == 5:
            if 4 <= date_in.day <= 9:
                result = YomTov.CHAZON
            if 10 <= date_in.day <= 16:
                result = YomTov.NACHAMU
        if date_in.month == 7:
            if 3 <= date_in.day <= 8:
                result = YomTov.SHUVA
        if get_parshah(date_in) == Parshah.BESHALACH:
            result = YomTov.SHIRA
    return result


def get_rosh_chodesh(date_in: HDate) -> YomTov:
    """If Rosh Chodesh returns ROSH_CHODESH; otherwise CHOL."""
    if date_in.day == 30 or (date_in.day == 1 and date_in.month != 7):
        return YomTov.ROSH_CHODESH
    return YomTov.CHOL


def get_machar_chodesh(date_in: HDate) -> YomTov:
    """If machar chodesh returns MACHAR_CHODESH; otherwise CHOL."""
    if date_in.wday:
        return YomTov.CHOL
    if date_in.day in (29, 30):
        return YomTov.MACHAR_CHODESH
    return YomTov.CHOL


def get_shabbos_mevorchim(date_in: HDate) -> YomTov:
    """If Shabbos Mevorchim returns SHABBOS_MEVORCHIM; otherwise CHOL."""
    if date_in.wday:
        return YomTov.CHOL
    if 23 <= date_in.day <= 29:
        return YomTov.SHABBOS_MEVORCHIM
    return YomTov.CHOL


def get_omer(date_in: HDate) -> int:
    """The omer count 1-49, or 0 if none."""
    omer = 0
    if date_in.month == 1 and date_in.day >= 16:
        omer = date_in.day - 15
    elif date_in.month == 2:
        omer = date_in.day + 15
    elif date_in.month == 3 and date_in.day <= 5:
        omer = date_in.day + 44
    return omer


def get_avos(date_in: HDate) -> int:
    """If Shabbos, returns the current chapter of Pirkei Avos (1-6, or 12/34/56
    for a double chapter); otherwise 0."""
    if date_in.wday:
        return 0  # Shabbos only

    avos_start = nissan_count(date_in.year) + 23  # 23 Nissan
    avos_day = date_in.dayOfYear - avos_start
    if avos_day <= 0:
        return 0

    chapter = avos_day // 7
    avos_day_of_week = avos_day % 7

    if avos_day_of_week == 6:  # tisha b'av is Shabbos
        if avos_day == 104:
            return 0
        elif avos_day > 104:
            chapter -= 1
    elif avos_day_of_week == 5:  # tisha b'av is Sunday
        if avos_day == 103:
            return 0
        elif avos_day > 103:
            chapter -= 1
    elif avos_day_of_week == 1 and not date_in.EY:  # 2nd day Shavous is Shabbos outside EY
        if avos_day == 43:
            return 0
        elif avos_day > 43:
            chapter -= 1

    chapter = (chapter % 6) + 1

    if date_in.month == 6:  # Elul double chapters
        if date_in.day > 22:
            chapter = 56
        elif date_in.day > 15:
            chapter = 34
        elif date_in.day > 8 and chapter == 1:
            chapter = 12

    return chapter


def is_ta_anis(date_in: HDate) -> bool:
    current = get_yom_tov(date_in)
    return current == YomTov.YOM_KIPPUR or (YomTov.SHIVA_ASAR_BTAAMUZ <= current <= YomTov.TAANIS_ESTER)


def is_assur_be_melachah(date_in: HDate) -> bool:
    current = get_yom_tov(date_in)
    return date_in.wday == 0 or (YomTov.PESACH_DAY1 <= current <= YomTov.SIMCHAS_TORAH)


def is_candle_lighting(date_in: HDate) -> int:
    """Returns 1 if regular candle lighting, 2 if at nightfall (motzei
    yom-tov/shabbos), 3 if chanukah, or 0 if none."""
    if date_in.wday == 6:
        return 1

    current = get_yom_tov(date_in)

    if (
        YomTov.EREV_PESACH <= current <= YomTov.EREV_SUKKOS
        or (current == YomTov.CHOL_HAMOED_PESACH_DAY4 and not date_in.EY)
        or (current == YomTov.CHOL_HAMOED_PESACH_DAY5 and date_in.EY)
        or current == YomTov.HOSHANA_RABBAH
    ):
        return 2 if date_in.wday == 0 else 1

    if current == YomTov.ROSH_HASHANAH_DAY1:
        return 2

    if current in (
        YomTov.PESACH_DAY1, YomTov.SHVEI_SHEL_PESACH, YomTov.SHAVOUS_DAY1,
        YomTov.SUKKOS_DAY1, YomTov.SHMEINI_ATZERES,
    ):
        return 2 if not date_in.EY else 0

    return 0


def is_birchas_ha_chama(date_in: HDate) -> bool:
    yearstart = hebrew_calendar_elapsed_days(date_in.year)
    day = yearstart + date_in.dayOfYear
    return day % 10227 == 172


def is_shabbos_mevorchim(date_in: HDate) -> bool:
    if date_in.wday != 0:
        return False
    if date_in.month == 6:  # Elul - no shabbos mevorchim
        return False
    return 23 <= date_in.day <= 29


def tekufas_tishrei_elapsed_days(date_in: HDate) -> int:
    """Days since Rosh Hashana year 1 for Tekufas Tishrei purposes (adds 1/2
    day since the first Tekufas Tishrei was 9 hours into the day, so all 4
    years of the solar leap cycle share day 47)."""
    days = hebrew_calendar_elapsed_days(date_in.year) + (date_in.dayOfYear - 1) + 0.5
    solar = (date_in.year - 1) * 365.25
    return round(days - solar)


def is_birchas_ha_shanim(date_in: HDate) -> bool:
    if date_in.EY:
        return date_in.month == 7 and date_in.day == 7
    return tekufas_tishrei_elapsed_days(date_in) == 47


def get_birchas_ha_shanim(date_in: HDate) -> bool:
    """True if Tal Umatar Livracha is said in Birchas HaShanim."""
    if date_in.month == 1 and date_in.day < 15:
        return True
    if date_in.month < 7:
        return False
    if date_in.EY:
        return not (date_in.month == 7 and date_in.day < 7)
    return tekufas_tishrei_elapsed_days(date_in) >= 47
