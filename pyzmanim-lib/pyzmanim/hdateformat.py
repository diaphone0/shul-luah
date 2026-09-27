"""
String formatting for Hebrew dates, holiday names, Torah portion names, molad
announcements, and Hebrew numeral (gematria-style) representation.

Ported from mod_hdateformat.bas in https://github.com/diaphone1/vbzmanim.
"""
from __future__ import annotations

from datetime import datetime, timedelta

from pyzmanim._data.hdateformat_data import HCHAR, HMONTH, HWDAY, PARSHAHCHAR, YOMTOV_FORMAT
from pyzmanim.hebrewcalendar import HDate, Parshah, YomTov, hdate_add_day, hdate_gregorian
from pyzmanim.noaa_calculator import Location


def parshah_format(current: Parshah) -> str:
    return PARSHAHCHAR[int(current)]


def get_h_char(num: int) -> str:
    """Maps a specific numeric value (1-10, 20, 30 ... 90, 100, 200, 300, 400,
    and special markers 99=geresh/999=gershayim) to its Hebrew letter."""
    mapping = {
        1: HCHAR[1], 2: HCHAR[2], 3: HCHAR[3], 4: HCHAR[4], 5: HCHAR[5],
        6: HCHAR[6], 7: HCHAR[7], 8: HCHAR[8], 9: HCHAR[9], 10: HCHAR[10],
        20: HCHAR[11], 30: HCHAR[12], 40: HCHAR[13], 50: HCHAR[14],
        60: HCHAR[15], 70: HCHAR[16], 80: HCHAR[17], 90: HCHAR[18],
        100: HCHAR[19], 200: HCHAR[20], 300: HCHAR[21], 400: HCHAR[22],
        99: HCHAR[23], 999: HCHAR[24],
    }
    return mapping.get(num, HCHAR[0])


def _add_char(year: str, charnum: int, num: int, counter: int, limit: int) -> tuple[str, int, int]:
    """Helper for building gematria strings. Returns the updated
    (year_buffer, num, counter). Special values 99/999 (geresh/gershayim
    separator marks) don't consume the running `num` value."""
    charvalue = 0 if charnum in (99, 999) else charnum
    len1 = limit - counter
    end_pos = counter + (2 if len1 > 2 else len1)
    new_year = year[:counter] + get_h_char(charnum) + year[end_pos:]
    new_counter = counter + (1 if len1 > 1 else len1)
    new_num = num - charvalue
    return new_year, new_num, new_counter


def num_to_h_char(innum: int) -> str:
    """Converts an int to a Hebrew-char gematria representation, e.g. 5779
    becomes תשע"ט."""
    num = innum
    year = "\0" * 13
    counter = 0

    if 1000 <= num <= 10000 and num % 1000 == 0:
        year, num, counter = _add_char(year, num // 1000, num, counter, 13)
        year, num, counter = _add_char(year, 999, num, counter, 13)
        return year[:counter]

    if 1000 <= num <= 10000:
        num %= 1000

    while num > 0 and counter < 13:
        if num == 15 or num == 16:
            year, num, counter = _add_char(year, 9, num, counter, 13)
            year, num, counter = _add_char(year, 99, num, counter, 13)
            year, num, counter = _add_char(year, num, num, counter, 13)
            return year[:counter]
        elif num < 10 or (num < 100 and num % 10 == 0) or (num < 500 and num % 100 == 0):
            if counter != 0:
                year, num, counter = _add_char(year, 99, num, counter, 13)
            year, num, counter = _add_char(year, num, num, counter, 13)
            result = year[:counter]
            if innum < 11:
                result += "'"
            return result
        elif num > 400:
            year, num, counter = _add_char(year, 400, num, counter, 13)
        elif num > 300:
            year, num, counter = _add_char(year, 300, num, counter, 13)
        elif num > 200:
            year, num, counter = _add_char(year, 200, num, counter, 13)
        elif num > 100:
            year, num, counter = _add_char(year, 100, num, counter, 13)
        elif num // 10 > 0:
            year, num, counter = _add_char(year, num - (num % 10), num, counter, 13)

    result = year[:counter]
    if innum < 11:
        result += "'"
    return result


def num_to_wday(date_in: HDate, shabbos: bool) -> str:
    """Convert int-based Hebrew weekday (hdate.wday) to char representation.
    `shabbos=True` uses שבת instead of שביעי for the 7th day."""
    if shabbos and date_in.wday == 0:
        return HWDAY[7]
    return HWDAY[date_in.wday]


def num_to_h_month(month: int, leap: int) -> str:
    """Convert int-based Hebrew month (hdate.month) to char representation."""
    if leap != 0:
        if month == 12:
            return HMONTH[0]
        elif month == 13:
            return HMONTH[month]
    if 0 < month < 13:
        return HMONTH[month]
    return ""


def hdate_format(date_in: HDate) -> str:
    """Convert an HDate to its string representation."""
    day = num_to_h_char(date_in.day)
    month = num_to_h_month(date_in.month, date_in.leap)
    year = num_to_h_char(date_in.year)
    return f"{day} {month} {year}"


def hdate_or_format(date_in: HDate, here: Location) -> str:
    """Convert an HDate to its string representation, with evening
    consideration: if the current real-world time is between tzais and the
    next sunrise, prefix with 'אור ל-' (the eve of) since the Hebrew calendar
    date has already advanced."""
    from pyzmanim import zmanim

    date_next = HDate(**date_in.__dict__)
    current_date = hdate_gregorian(date_in)
    hdate_add_day(date_next, 1)

    sunset_today = hdate_gregorian(zmanim.gettzais8p5(date_in, here))
    sunrise_today = hdate_gregorian(zmanim.getsunrise(date_in, here))

    is_or = False
    if current_date >= sunset_today:
        is_or = True
        date_result = date_next
    else:
        date_result = date_in

    if current_date < sunrise_today:
        is_or = True

    day = num_to_h_char(date_result.day)
    month = num_to_h_month(date_result.month, date_result.leap)
    year = num_to_h_char(date_result.year)
    result = f"{day} {month} {year}"
    if is_or:
        result = "אור ל-" + result
    return result


def molad_format(molad: HDate, full_date: bool = True) -> str:
    """Format an HDate holding molad info for a molad announcement."""
    time_str = f"{molad.hour:02d}:{molad.min:02d}"
    if full_date:
        return (
            f"יום {num_to_wday(molad, True)}, {num_to_h_char(molad.day)} "
            f"{num_to_h_month(molad.month, molad.leap)}, שעה {time_str} ו-{molad.sec} חלקים"
        )
    return f"יום {num_to_wday(molad, True)} שעה {time_str} ו-{molad.sec} חלקים"


def yom_tov_format(current: YomTov) -> str:
    """Convert a YomTov value to its Hebrew display title."""
    return YOMTOV_FORMAT.get(YomTov(current).name, "")


def avos_format(avos: int) -> str:
    """Convert a Pirkei Avos chapter number to its Hebrew char representation."""
    mapping = {1: "א", 2: "ב", 3: "ג", 4: "ד", 5: "ה", 6: "ו", 12: "א-ב", 34: "ג-ד", 56: "ה-ו"}
    return mapping.get(avos, "")


def tround(t: datetime) -> datetime:
    """Rounds seconds of a time to an added minute if seconds > 29."""
    return t + timedelta(minutes=1) if t.second > 29 else t
