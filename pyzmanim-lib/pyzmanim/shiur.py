"""
Daily/weekly Torah study schedule text generators.

Ported from mod_shiur.bas in https://github.com/diaphone1/vbzmanim.

Covers: weekly Torah portion (Chumash) with Rashi, monthly Tehillim (Psalms)
division, daily Tanya chapters, daily Halacha cycle, and daily Rambam
(Mishneh Torah) chapters (both the 1-chapter/day and 3-chapters/day cycles).
"""
from __future__ import annotations

from datetime import date

from pyzmanim._data.shiur_data import (
    HALACHA_YOMI,
    RAMBAM_1CHP,
    RAMBAM_3CHP,
    TANYA_LEAP,
    TANYA_REGULAR,
    TEHILLIM_ARRAY,
)
from pyzmanim.hdateformat import num_to_wday, parshah_format
from pyzmanim.hebrewcalendar import (
    HDate,
    Parshah,
    get_parshah,
    hdate_add_day,
    hdate_gregorian,
    hebrew_calendar_elapsed_days,
    hebrew_leap_year,
    last_day_of_hebrew_month,
)

HALACHA_CYCLE_START = date(1954, 8, 16)
RAMBAM_1CHP_CYCLE_START = date(1984, 4, 29)


def chumash(date_in: HDate) -> str:
    """The weekly Torah portion (with Rashi) for the Shabbos on/following
    date_in, handling the Bereishit/Vezot Habrachah edge case around
    Simchas Torah."""
    shabbos = HDate(**date_in.__dict__)
    while True:
        current = get_parshah(shabbos)
        if current == Parshah.NOPARSHAH:
            hdate_add_day(shabbos, 7 - shabbos.wday)
        if current != Parshah.NOPARSHAH:
            break

    if current == Parshah.BERESHIT:
        st = 22 if shabbos.EY else 23
        if date_in.day < st:
            current = Parshah.VZOT_HABERACHAH
        elif date_in.day == st:
            if date_in.wday == 1:
                return f"פרשת {parshah_format(Parshah.BERESHIT)}\n{num_to_wday(date_in, False)} עם פירש״י"
            elif date_in.wday == 0:
                return f"פרשת {parshah_format(Parshah.VZOT_HABERACHAH)}\n{num_to_wday(date_in, False)} עם פירש״י"
            else:
                return ""

    return f"פרשת {parshah_format(current)}\n{num_to_wday(date_in, False)} עם פירש״י"


def tehillim(date_in: HDate) -> str:
    """Monthly Tehillim (Psalms) division for the day of the Hebrew month.
    Day 29 rolls to index 0 (combined with day 30) in months without a 30th
    day."""
    current = date_in.day
    if current == 29 and last_day_of_hebrew_month(date_in.month, date_in.year) != 30:
        current = 0
    return TEHILLIM_ARRAY[current]


def get_tanya(date_in: HDate) -> str:
    """Daily Tanya chapter reference for date_in."""
    last = date_in.month in (7, 8) or (date_in.month == 9 and date_in.day < 19)
    leap = hebrew_leap_year(date_in.year - 1 if last else date_in.year)
    current = date_in.day
    if current == 29 and last_day_of_hebrew_month(date_in.month, date_in.year) != 30:
        current = 0

    table = TANYA_LEAP if leap else TANYA_REGULAR
    shiur = table[date_in.month - 1][current].split(";")

    result = f"{shiur[0]} {shiur[1]}:\n" + f"ד.ה. {shiur[2]} ... {shiur[3]}"
    if len(shiur) > 4:
        result += f" [עמוד {shiur[4]}]"
    return result


def get_halacha(date_in: HDate) -> str:
    """Daily Halacha (1613-day cycle since 16 August 1954)."""
    days = (hdate_gregorian(date_in).date() - HALACHA_CYCLE_START).days
    if days < 0:
        return "-"
    current = days % 1613
    if current == 1612:
        return HALACHA_YOMI[403]
    parts = HALACHA_YOMI[current // 4].split(";")
    return parts[current % 4]


def get_rambam(date_in: HDate, daily_chapter: bool) -> str:
    """Daily Rambam (Mishneh Torah) chapter.

    `daily_chapter=True` uses the 1-chapter/day cycle (1017-day cycle since
    29 April 1984: 1000 chapters + hakdamah + haggada + seder tefila days).
    `daily_chapter=False` uses the 3-chapters/day cycle, keyed by Hebrew
    calendar day-of-year (339-day cycle)."""
    if daily_chapter:
        days = (hdate_gregorian(date_in).date() - RAMBAM_1CHP_CYCLE_START).days
        if days < 0:
            return "-"
        return RAMBAM_1CHP[days % 1017]

    days = (hebrew_calendar_elapsed_days(date_in.year) + date_in.dayOfYear) - 2097823
    if days < 0:
        return "-"
    return RAMBAM_3CHP[days % 339]
