"""
Halachic time (zmanim) calculations.

Ported from mod_zmanim.bas in https://github.com/diaphone1/vbzmanim.

Computes zmanim (halachic times) - alos (dawn), sunrise, shema/tefila
cutoffs, chatzos (midday), mincha times, sunset, tzais (nightfall), plag
hamincha, molad-related offsets, and shaah zmanis (proportional halachic
hour) - per multiple halachic opinions (GRA, Magen Avraham, Baal HaTanya).

All "get*" functions return an `HDate` (a Hebrew calendar date/time - the
resulting HDate carries the correct real-world Gregorian moment, retrievable
via `hebrewcalendar.hdate_gregorian`).
"""
from __future__ import annotations

from pyzmanim.hebrewcalendar import (
    HDate,
    EMPTY_HDATE,
    hdate_julian,
    hdate_add_day,
    hdate_add_second,
    hdate_add_msecond,
    hebrew_calendar_elapsed_days,
)
from pyzmanim.noaa_calculator import Location, get_utc_sunrise, get_utc_sunset

GEOMETRIC_ZENITH = 90.0
ZENITH_AMITIS = GEOMETRIC_ZENITH + 1.583

ZENITH_26_D = GEOMETRIC_ZENITH + 26.0
ZENITH_19_P_8 = GEOMETRIC_ZENITH + 19.8
ZENITH_18_D = GEOMETRIC_ZENITH + 18
ZENITH_16_P_9 = GEOMETRIC_ZENITH + 16.9
ZENITH_16_P_1 = GEOMETRIC_ZENITH + 16.1
ZENITH_13_P_24 = GEOMETRIC_ZENITH + 13.24

ZENITH_11_P_5 = GEOMETRIC_ZENITH + 11.5
ZENITH_11_D = GEOMETRIC_ZENITH + 11
ZENITH_10_P_2 = GEOMETRIC_ZENITH + 10.2

ZENITH_3_P_65 = GEOMETRIC_ZENITH + 3.65
ZENITH_3_P_676 = GEOMETRIC_ZENITH + 3.676
ZENITH_3_P_7 = GEOMETRIC_ZENITH + 3.7
ZENITH_3_P_8 = GEOMETRIC_ZENITH + 3.8
ZENITH_4_P_37 = GEOMETRIC_ZENITH + 4.37
ZENITH_4_P_61 = GEOMETRIC_ZENITH + 4.61
ZENITH_4_P_8 = GEOMETRIC_ZENITH + 4.8
ZENITH_5_P_88 = GEOMETRIC_ZENITH + 5.88
ZENITH_5_P_95 = GEOMETRIC_ZENITH + 5.95
ZENITH_6_D = GEOMETRIC_ZENITH + 6
ZENITH_7_P_083 = GEOMETRIC_ZENITH + 7.083
ZENITH_8_P_5 = GEOMETRIC_ZENITH + 8.5

MINUTES60 = 60 * 60000
MINUTES72 = 72 * 60000
MINUTES90 = 90 * 60000
MINUTES96 = 96 * 60000
MINUTES120 = 120 * 60000
MINUTES18 = 18 * 60000


def get_local_mean_time_offset(now: HDate, here: Location) -> int:
    return int(here.longitude * 4 * 60 - now.offset)


def get_antimeridian_adjustment(now: HDate, here: Location) -> int:
    local_hours_offset = get_local_mean_time_offset(now, here) / 3600
    if local_hours_offset >= 20:
        return 1
    elif local_hours_offset <= -20:
        return -1
    return 0


def get_date_from_time(current: HDate, time: float, here: Location, is_sunrise: bool) -> HDate:
    result = HDate(
        year=current.year, EY=current.EY, offset=current.offset,
        month=current.month, day=current.day,
    )

    adjustment = get_antimeridian_adjustment(current, here)
    if adjustment != 0:
        hdate_add_day(result, adjustment)

    hours = int(time)
    time = (time - hours) * 60
    minutes = int(time)
    time = (time - minutes) * 60
    seconds = int(time)
    time = (time - seconds) * 1000
    milliseconds = int(time)

    local_time_hours = int(here.longitude / 15)
    if is_sunrise and local_time_hours + hours > 18:
        hdate_add_day(result, -1)
    elif not is_sunrise and local_time_hours + hours < 6:
        hdate_add_day(result, 1)

    result.hour = hours
    result.min = minutes
    result.sec = seconds
    result.msec = milliseconds
    hdate_add_second(result, current.offset)

    return result


def calcsunrise(date_in: HDate, here: Location, zenith: float, adjust_for_elevation: bool) -> HDate:
    sunrise = get_utc_sunrise(hdate_julian(date_in), here, zenith, adjust_for_elevation)
    return get_date_from_time(date_in, sunrise, here, True)


def calcsunset(date_in: HDate, here: Location, zenith: float, adjust_for_elevation: bool) -> HDate:
    sunset = get_utc_sunset(hdate_julian(date_in), here, zenith, adjust_for_elevation)
    return get_date_from_time(date_in, sunset, here, False)


def calcshaahzmanis(startday: HDate, endday: HDate) -> int:
    """Length of a proportional halachic hour, in milliseconds, based on the
    span between startday and endday divided by 12."""
    if startday.year == 0 or endday.year == 0:
        return 0
    start = hebrew_calendar_elapsed_days(startday.year) + (startday.dayOfYear - 1)
    end = hebrew_calendar_elapsed_days(endday.year) + (endday.dayOfYear - 1)
    diff = end - start
    diff = (diff * 24) + (endday.hour - startday.hour)
    diff = (diff * 60) + (endday.min - startday.min)
    diff = (diff * 60) + (endday.sec - startday.sec)
    diff = (diff * 1000) + (endday.msec - startday.msec)
    return diff // 12


def calctimeoffset(time: HDate, offset: int) -> HDate:
    # In the original VBA, `offset As Long` implicitly rounds any Double
    # passed in (e.g. shaahzmanis * 6.5) to the nearest integer number of
    # milliseconds. Python has no such implicit coercion, so callers that
    # pass a float (from a non-integer shaahzmanis multiplier) must be
    # rounded here explicitly - otherwise the float propagates into
    # HDate.hour/min/sec via hdate_add_msecond's carry cascade and breaks
    # datetime construction in hdate_gregorian.
    offset = round(offset)
    if time.year == 0 or offset == 0:
        return HDate()
    result = HDate(**time.__dict__)
    hdate_add_msecond(result, offset)
    return result


def getalos(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, ZENITH_16_P_1, False)


def getalosbaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, ZENITH_16_P_9, False)


def getalos26degrees(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, ZENITH_26_D, False)


def getalos19p8degrees(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, ZENITH_19_P_8, False)


def getalos18degrees(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, ZENITH_18_D, False)


def getalos120(date_in: HDate, here: Location) -> HDate:
    return calctimeoffset(getsunrise(date_in, here), -MINUTES120)


def getalos120zmanis(date_in: HDate, here: Location) -> HDate:
    shaahzmanis = getshaahzmanisgra(date_in, here)
    if shaahzmanis == 0:
        return EMPTY_HDATE
    return calctimeoffset(getsunrise(date_in, here), int(shaahzmanis * -2))


def getalos96(date_in: HDate, here: Location) -> HDate:
    return calctimeoffset(getsunrise(date_in, here), -MINUTES96)


def getalos96zmanis(date_in: HDate, here: Location) -> HDate:
    shaahzmanis = getshaahzmanisgra(date_in, here)
    if shaahzmanis == 0:
        return EMPTY_HDATE
    return calctimeoffset(getsunrise(date_in, here), int(shaahzmanis * -1.6))


def getalos90(date_in: HDate, here: Location) -> HDate:
    return calctimeoffset(getsunrise(date_in, here), -MINUTES90)


def getalos90zmanis(date_in: HDate, here: Location) -> HDate:
    shaahzmanis = getshaahzmanisgra(date_in, here)
    if shaahzmanis == 0:
        return EMPTY_HDATE
    return calctimeoffset(getsunrise(date_in, here), int(shaahzmanis * -1.5))


def getalos72(date_in: HDate, here: Location) -> HDate:
    return calctimeoffset(getsunrise(date_in, here), -MINUTES72)


def getalos72zmanis(date_in: HDate, here: Location) -> HDate:
    shaahzmanis = getshaahzmanisgra(date_in, here)
    if shaahzmanis == 0:
        return EMPTY_HDATE
    return calctimeoffset(getsunrise(date_in, here), int(shaahzmanis * -1.2))


def getalos60(date_in: HDate, here: Location) -> HDate:
    return calctimeoffset(getsunrise(date_in, here), -MINUTES60)


def getmisheyakir11p5degrees(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, ZENITH_11_P_5, False)


def getmisheyakir11degrees(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, ZENITH_11_D, False)


def getmisheyakir10p2degrees(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, ZENITH_10_P_2, False)


def getsunrise(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, GEOMETRIC_ZENITH, False)


def getsunrisebaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, ZENITH_AMITIS, False)


def getelevationsunrise(date_in: HDate, here: Location) -> HDate:
    return calcsunrise(date_in, here, GEOMETRIC_ZENITH, True)


def calcshma(startday: HDate, endday: HDate) -> HDate:
    shaahzmanis = calcshaahzmanis(startday, endday)
    return calctimeoffset(startday, shaahzmanis * 3)


def getshmabaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcshma(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def getshmagra(date_in: HDate, here: Location) -> HDate:
    return calcshma(getsunrise(date_in, here), getsunset(date_in, here))


def getshmamga(date_in: HDate, here: Location) -> HDate:
    return calcshma(getalos72(date_in, here), gettzais72(date_in, here))


def calctefila(startday: HDate, endday: HDate) -> HDate:
    shaahzmanis = calcshaahzmanis(startday, endday)
    return calctimeoffset(startday, shaahzmanis * 4)


def gettefilabaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calctefila(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def gettefilagra(date_in: HDate, here: Location) -> HDate:
    return calctefila(getsunrise(date_in, here), getsunset(date_in, here))


def gettefilamga(date_in: HDate, here: Location) -> HDate:
    return calctefila(getalos72(date_in, here), gettzais72(date_in, here))


def getachilaschometzbaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calctefila(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def getachilaschometzgra(date_in: HDate, here: Location) -> HDate:
    return calctefila(getsunrise(date_in, here), getsunset(date_in, here))


def getachilaschometzmga(date_in: HDate, here: Location) -> HDate:
    return calctefila(getalos72(date_in, here), gettzais72(date_in, here))


def calcbiurchometz(startday: HDate, endday: HDate) -> HDate:
    shaahzmanis = calcshaahzmanis(startday, endday)
    return calctimeoffset(startday, shaahzmanis * 5)


def getbiurchometzbaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcbiurchometz(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def getbiurchometzgra(date_in: HDate, here: Location) -> HDate:
    return calcbiurchometz(getsunrise(date_in, here), getsunset(date_in, here))


def getbiurchometzmga(date_in: HDate, here: Location) -> HDate:
    return calcbiurchometz(getalos72(date_in, here), gettzais72(date_in, here))


def calcchatzos(startday: HDate, endday: HDate) -> HDate:
    shaahzmanis = calcshaahzmanis(startday, endday)
    return calctimeoffset(startday, shaahzmanis * 6)


def getchatzosbaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcchatzos(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def getchatzosgra(date_in: HDate, here: Location) -> HDate:
    return calcchatzos(getsunrise(date_in, here), getsunset(date_in, here))


def calcminchagedola(startday: HDate, endday: HDate) -> HDate:
    shaahzmanis = calcshaahzmanis(startday, endday)
    return calctimeoffset(startday, shaahzmanis * 6.5)


def getminchagedolabaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcminchagedola(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def getminchagedolagra(date_in: HDate, here: Location) -> HDate:
    return calcminchagedola(getsunrise(date_in, here), getsunset(date_in, here))


def getminchagedolamga(date_in: HDate, here: Location) -> HDate:
    return calcminchagedola(getalos72(date_in, here), gettzais72(date_in, here))


def calcminchagedola30min(startday: HDate, endday: HDate) -> HDate:
    shaahzmanis = calcshaahzmanis(startday, endday)
    return calctimeoffset(startday, int(shaahzmanis * 6) + 1800000)


def calcminchagedolagreater30min(startday: HDate, endday: HDate) -> HDate:
    if (calcshaahzmanis(startday, endday) * 0.5) >= 1800000:
        return calcminchagedola(startday, endday)
    return calcminchagedola30min(startday, endday)


def getminchagedolabaalhatanyag30m(date_in: HDate, here: Location) -> HDate:
    return calcminchagedolagreater30min(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def getminchagedolagrag30m(date_in: HDate, here: Location) -> HDate:
    return calcminchagedolagreater30min(getsunrise(date_in, here), getsunset(date_in, here))


def getminchagedolamgag30m(date_in: HDate, here: Location) -> HDate:
    return calcminchagedolagreater30min(getalos72(date_in, here), gettzais72(date_in, here))


def calcminchaketana(startday: HDate, endday: HDate) -> HDate:
    shaahzmanis = calcshaahzmanis(startday, endday)
    return calctimeoffset(startday, shaahzmanis * 9.5)


def getminchaketanabaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcminchaketana(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def getminchaketanagra(date_in: HDate, here: Location) -> HDate:
    return calcminchaketana(getsunrise(date_in, here), getsunset(date_in, here))


def getminchaketanamga(date_in: HDate, here: Location) -> HDate:
    return calcminchaketana(getalos72(date_in, here), gettzais72(date_in, here))


def calcplag(startday: HDate, endday: HDate) -> HDate:
    shaahzmanis = calcshaahzmanis(startday, endday)
    return calctimeoffset(startday, shaahzmanis * 10.75)


def getplagbaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcplag(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def getplaggra(date_in: HDate, here: Location) -> HDate:
    return calcplag(getsunrise(date_in, here), getsunset(date_in, here))


def getplagmga(date_in: HDate, here: Location) -> HDate:
    return calcplag(getalos72(date_in, here), gettzais72(date_in, here))


def getcandlelighting(date_in: HDate, here: Location) -> HDate:
    return calctimeoffset(calcsunset(date_in, here, GEOMETRIC_ZENITH, False), -MINUTES18)


def getsunset(date_in: HDate, here: Location) -> HDate:
    return calcsunset(date_in, here, GEOMETRIC_ZENITH, False)


def getsunsetbaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcsunset(date_in, here, ZENITH_AMITIS, False)


def getelevationsunset(date_in: HDate, here: Location) -> HDate:
    return calcsunset(date_in, here, GEOMETRIC_ZENITH, True)


def gettzaisbaalhatanya(date_in: HDate, here: Location) -> HDate:
    return calcsunset(date_in, here, ZENITH_6_D, True)


def gettzais8p5(date_in: HDate, here: Location) -> HDate:
    return calcsunset(date_in, here, ZENITH_8_P_5, True)


def gettzais72(date_in: HDate, here: Location) -> HDate:
    return calctimeoffset(getsunset(date_in, here), MINUTES72)


def gettzaisyeshiva(date_in: HDate, here: Location, adjust_for_elevation: bool = False) -> HDate:
    """Tzais hakochavim per the ruling of Rav Zalman Baruch Melamed
    shlit"a (Yeshivat Har Hamor): in summer (day length greater than 12
    hours - i.e. a shaah zmanis longer than 60 real minutes), tzais is 18
    "zmanis" (proportional) minutes after the flat/level sunset (no
    elevation adjustment); in winter (day 12 hours or shorter), tzais is
    a fixed 18 real minutes after that same sunset instead - so tzais is
    never less than 18 minutes after sunset either way. Does not (yet)
    implement the stricter Tisha B'Av exception (6.45 degrees below the
    horizon) mentioned in the same ruling - only the general case for
    ordinary fast days.

    adjust_for_elevation=True uses the elevation-adjusted sunset
    (getelevationsunset) as the base sunset instead of the flat/level one
    (getsunset) - matching the same adjust_for_elevation convention used
    by calcsunrise/calcsunset elsewhere in this module. Defaults to False
    (flat sunset), matching the ruling's own wording ("ha-shkiah
    ha-mishorit" - the level/flat sunset, not an elevation-adjusted one)."""
    sunset = getelevationsunset(date_in, here) if adjust_for_elevation else getsunset(date_in, here)
    shaahzmanis = getshaahzmanisgra(date_in, here)
    if shaahzmanis == 0:
        return EMPTY_HDATE
    # shaahzmanis is the length of one proportional hour, in milliseconds;
    # comparing it to MINUTES60 (60 real minutes) tells us whether the day
    # is longer (summer) or shorter-or-equal (winter) than 12 hours - a
    # shaah zmanis longer than a real hour means the whole (12-hour) day is
    # longer than 12 real hours. 18 "zmanis" minutes = 18/60 of one shaah
    # zmanis.
    offset = (shaahzmanis * 18 / 60) if shaahzmanis > MINUTES60 else MINUTES18
    return calctimeoffset(sunset, offset)


def calcmoladoffset(date_in: HDate, offsetsec: int) -> HDate:
    from pyzmanim.hebrewcalendar import get_molad

    result = get_molad(date_in.year, date_in.month)
    tz = (-result.offset) + date_in.offset
    adjustment = (result.sec * 10 // 3) + tz + offsetsec
    result.sec = 0
    hdate_add_second(result, adjustment)
    result.EY = date_in.EY
    result.offset = date_in.offset
    return result


def getmolad7days(date_in: HDate) -> HDate:
    return calcmoladoffset(date_in, 604800)


def getmoladhalfmonth(date_in: HDate) -> HDate:
    return calcmoladoffset(date_in, 1275722)


def getmolad15days(date_in: HDate) -> HDate:
    return calcmoladoffset(date_in, 1296000)


def getshaahzmanisbaalhatanya(date_in: HDate, here: Location) -> int:
    return calcshaahzmanis(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))


def getshaahzmanisgra(date_in: HDate, here: Location) -> int:
    return calcshaahzmanis(getsunrise(date_in, here), getsunset(date_in, here))


def getshaahzmanismga(date_in: HDate, here: Location) -> int:
    return calcshaahzmanis(getalos72(date_in, here), gettzais72(date_in, here))
