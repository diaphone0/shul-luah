"""
Builds the per-refresh "zman context" - the current Hebrew date/time plus a
few derived values (the relevant Shabbos/Erev-Shabbos HDate, and whether
today counts as chol for weekday/Shabbos slide switching) that the tag
renderers in tagging.py need.

This mirrors the hebrewDate/shabbos/erevshabbos/today_is_chol computation in
LuahMain.bas's `luah`/`init_shapes` subs, but computes the UTC offset
dynamically via zoneinfo (DST-aware) instead of a hardcoded fixed offset.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import datetime
from zoneinfo import ZoneInfo

from pyzmanim.hebrewcalendar import (
    HDate,
    convert_date,
    hdate_add_day,
    hdate_gregorian,
    is_assur_be_melachah,
    is_candle_lighting,
)
from pyzmanim.noaa_calculator import Location
from pyzmanim.zmanim import getchatzosgra

from .tzdata_support import ensure_tzdata_available

ensure_tzdata_available()


@dataclass
class ZmanContext:
    now: HDate
    location: Location
    shabbos: HDate
    erev_shabbos: HDate
    today_is_chol: bool


def _utc_offset_seconds(timezone: str, at: datetime) -> int:
    tz = ZoneInfo(timezone)
    aware = at.replace(tzinfo=tz) if at.tzinfo is None else at.astimezone(tz)
    offset = aware.utcoffset()
    return int(offset.total_seconds()) if offset else 0


def build_zman_context(
    now: datetime,
    location: Location,
    eretz_yisroel: bool,
    timezone: str,
) -> ZmanContext:
    hebrew_date = convert_date(now)
    hebrew_date.offset = _utc_offset_seconds(timezone, now)
    hebrew_date.EY = eretz_yisroel

    erev_shabbos = replace(hebrew_date)
    shabbos = replace(hebrew_date)

    if is_assur_be_melachah(erev_shabbos):
        hdate_add_day(erev_shabbos, -1)
    else:
        while is_candle_lighting(erev_shabbos) == 0:
            hdate_add_day(erev_shabbos, 1)
        shabbos = replace(erev_shabbos)
        hdate_add_day(shabbos, 1)

    if hebrew_date.wday == 6:
        today_is_chol = hdate_gregorian(hebrew_date) <= hdate_gregorian(
            getchatzosgra(hebrew_date, location)
        )
    elif hebrew_date.wday == 0:
        today_is_chol = False
    else:
        today_is_chol = True

    return ZmanContext(
        now=hebrew_date,
        location=location,
        shabbos=shabbos,
        erev_shabbos=erev_shabbos,
        today_is_chol=today_is_chol,
    )
