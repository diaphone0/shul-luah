"""
Daily Talmud (Gemara) study cycle - Bavli (Babylonian) and Yerushalmi
(Jerusalem) Talmud, plus Mishna Yomi (daily Mishnah study).

Ported from mod_dafyomi.bas in https://github.com/diaphone1/vbzmanim, which
itself ports the dafyomi list & functions from https://github.com/NykUser/MyZman/.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta

from pyzmanim._data.dafyomi_data import (
    BLATT_PER_MASECHTA,
    BLATT_PER_MASECHTA_YERUSHALMI,
    MASECHTOS_BAVLI,
    MASECHTOS_YERUSHALMI,
    MISHNAYOS_ARR,
)
from pyzmanim.hdateformat import num_to_h_char
from pyzmanim.hebrewcalendar import (
    YomTov,
    convert_date,
    get_yom_tov,
    hdate_add,
    hdate_add_year,
    hdate_gregorian,
    hdate_new,
)

DAF_YOMI_START_DATE = date(1923, 9, 11)
SHEKALIM_CHANGE_DATE = date(1975, 6, 24)
YERUSHALMI_START_DATE = date(1980, 2, 2)
MISHNAH_YOMI_START_DATE = date(1947, 5, 20)
WHOLE_SHAS_DAFS = 1554
MISHNAH_DAYS_PER_CYCLE = 2096  # total of 4192 mishnayos / 2 per day


@dataclass
class Daf:
    masechta_number: int = 0
    page: int = 0  # reused as mishnah number for Mishna Yomi
    has_secondary_mesechta: bool = False
    secondary_mesechta_number: int = 0
    perek_number: int = 0  # reused as perek for Mishna Yomi


def _is_current_blatt_with_next_masechta(masechta: int, blatt: int) -> bool:
    return (masechta == 35 and blatt == 22) or (masechta == 36 and blatt == 25)


def get_daf_yomi_bavli(date_in: date) -> Daf:
    """Returns the Daf Yomi Bavli page for a given date (as Daf). The first
    Daf Yomi cycle started 11 September 1923 and dates before that return
    page=-1 (no daf). The number of blatt in Shekalim changed on the 8th Daf
    Yomi cycle (24 June 1975) from 13 to 22 - blattPerMasechta[4] is adjusted
    per-cycle to account for this."""
    if date_in < DAF_YOMI_START_DATE:
        return Daf(page=-1)

    days_from_start = (date_in - DAF_YOMI_START_DATE).days
    days_from_shekalim_change = (date_in - SHEKALIM_CHANGE_DATE).days

    blatt_per_masechta = list(BLATT_PER_MASECHTA)

    if date_in >= SHEKALIM_CHANGE_DATE:
        cycle_no = 8 + (days_from_shekalim_change // 2711)
        daf_no = days_from_shekalim_change % 2711
    else:
        cycle_no = 1 + (days_from_start // 2702)
        daf_no = days_from_start % 2702

    # Fix Shekalim for old cycles.
    blatt_per_masechta[4] = 13 if cycle_no <= 7 else 22

    total = 0
    for masechta in range(40):
        total += blatt_per_masechta[masechta] - 1
        if daf_no < total:
            blatt = 1 + blatt_per_masechta[masechta] - (total - daf_no)
            # Fiddle with the weird ones near the end (Tamid/Midos/Kinnim).
            if masechta == 36:
                blatt += 21
            elif masechta == 37:
                blatt += 24
            elif masechta == 38:
                blatt += 32

            return Daf(
                masechta_number=masechta,
                page=blatt,
                has_secondary_mesechta=_is_current_blatt_with_next_masechta(masechta, blatt),
                secondary_mesechta_number=masechta + 1,
            )

    return Daf(page=-1)


def get_num_of_special_days(date_start: date, date_end: date) -> int:
    """Number of special days (Yom Kippur and Tisha B'Av, on which there is
    no daf) between date_start and date_end, inclusive."""
    result = 0
    start_year = convert_date(datetime(date_start.year, date_start.month, date_start.day)).year
    end_year = convert_date(datetime(date_end.year, date_end.month, date_end.day)).year

    yom_kipur = hdate_new()
    tisha_beav = hdate_new()
    hdate_add(yom_kipur, years=start_year, months=7, days=10)
    hdate_add(tisha_beav, years=start_year, months=5, days=9)

    for _ in range(start_year, end_year + 1):
        date_check = hdate_gregorian(yom_kipur).date()
        if date_start <= date_check <= date_end:
            result += 1

        date_check = hdate_gregorian(tisha_beav).date()
        if date_start <= date_check <= date_end:
            result += 1

        hdate_add_year(yom_kipur, 1)
        hdate_add_year(tisha_beav, 1)

    return result


def get_daf_yomi_yerushalmi(date_in: date) -> Daf:
    """Returns the Daf Yomi Yerushalmi page for a given date (as Daf). The
    first cycle started 15 Shevat 5740 (2 February 1980); dates prior to
    this, or on Yom Kippur/Tisha B'Av (no daf on those days), return
    page=-1."""
    hdate_check = convert_date(datetime(date_in.year, date_in.month, date_in.day))

    if (
        get_yom_tov(hdate_check) in (YomTov.YOM_KIPPUR, YomTov.TISHA_BAV)
        or date_in < YERUSHALMI_START_DATE
    ):
        return Daf(page=-1)

    next_cycle = YERUSHALMI_START_DATE
    prev_cycle = YERUSHALMI_START_DATE

    while date_in > next_cycle:
        prev_cycle = next_cycle
        next_cycle = next_cycle + timedelta(days=WHOLE_SHAS_DAFS)
        next_cycle = next_cycle + timedelta(days=get_num_of_special_days(prev_cycle, next_cycle))

    daf_no = (date_in - prev_cycle).days
    daf_no -= get_num_of_special_days(prev_cycle, date_in)

    masechta_number = 0
    page = daf_no
    for i in range(39):
        if daf_no < BLATT_PER_MASECHTA_YERUSHALMI[i]:
            masechta_number = i
            page = daf_no
            break
        daf_no -= BLATT_PER_MASECHTA_YERUSHALMI[i]

    return Daf(masechta_number=masechta_number, page=page + 1)


def get_daf_yomi_format(date_in: date, yerushalmi: bool = False) -> str:
    if yerushalmi:
        result = get_daf_yomi_yerushalmi(date_in)
        if result.page == -1:
            return "-"
        return f"{MASECHTOS_YERUSHALMI[result.masechta_number]} דף {num_to_h_char(result.page)}"
    else:
        result = get_daf_yomi_bavli(date_in)
        if result.page == -1:
            return "-"
        return f"{MASECHTOS_BAVLI[result.masechta_number]} דף {num_to_h_char(result.page)}"


def get_mishna_yomi(date_in: date) -> Daf:
    """Returns the Mishna Yomi (2 mishnayos/day) position for a given date.
    Officially the cycle started 6 Sivan 5707 (Shavuos 1947), but since the
    first cycle was 5 days shorter, 1 Sivan (20 May 1947) is used instead."""
    if date_in < MISHNAH_YOMI_START_DATE:
        return Daf(page=-1)

    day_in_cycle = (date_in - MISHNAH_YOMI_START_DATE).days % MISHNAH_DAYS_PER_CYCLE

    total_sum = 0
    for i in range(63):
        parts = MISHNAYOS_ARR[i].split(",")
        nxt_sum = 0
        for j in range(2, len(parts)):  # 0/1 are the masechta's hebrew/english names
            m_since_cycle = day_in_cycle * 2  # mishnayos since start of cycle
            m_since_masechet = m_since_cycle - total_sum  # since start of masechta
            m_since_perek = m_since_masechet - nxt_sum  # since start of perek
            if total_sum <= m_since_cycle < total_sum + nxt_sum + int(parts[j]):
                return Daf(masechta_number=i, perek_number=j - 1, page=m_since_perek + 1)
            nxt_sum += int(parts[j])
        total_sum += nxt_sum

    return Daf(page=-1)  # unreachable for a valid day_in_cycle


def get_mishna_yomi_format(date_in: date) -> str:
    result = get_mishna_yomi(date_in)
    if result.page == -1:
        return "-"
    parts = MISHNAYOS_ARR[result.masechta_number].split(",")
    return f"{parts[0]} פרק {num_to_h_char(result.perek_number)} משנה {num_to_h_char(result.page)}"
