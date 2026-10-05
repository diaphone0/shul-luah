#!/usr/bin/env python3
"""
Print a full Jewish calendar month view (6×7 grid), mirroring the VBA
calendar_utils_get_month_matrix + calendar_utils_get_day_info logic.
Uses pyzmanim for all Hebrew date, parsha, yom tov, molad, omer, and
zmanim calculations.
"""

import sys

# The active console code page (e.g. Windows cp1255) may not cover every
# character this script prints (box-drawing borders, Hebrew, gematria
# punctuation like ״/׳, etc.) - reconfigure stdout/stderr to UTF-8 up front
# so printing never raises UnicodeEncodeError (same fix pattern already
# used in luah_signage/app.py for the same class of problem).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass

sys.path.insert(0, "pyzmanim-lib")

import re
from datetime import datetime, timedelta
from pyzmanim.hebrewcalendar import (
    HDate, convert_date, hdate_add_day, hdate_add_month, hdate_add_year,
    hdate_gregorian, hdate_new, last_day_of_hebrew_month,
    get_molad, get_omer, get_parshah, get_rosh_chodesh,
    get_special_shabbos, get_yom_tov, is_candle_lighting,
    is_shabbos_mevorchim, is_assur_be_melachah,
    YomTov, Parshah,
)
from pyzmanim.hdateformat import (
    hdate_format, molad_format, num_to_h_char, num_to_h_month,
    parshah_format, yom_tov_format,
)
from pyzmanim.zmanim import getelevationsunset, gettzais8p5, getsunrise, getsunset, Location
from pyzmanim.dafyomi import get_daf_yomi_format, get_mishna_yomi_format
from pyzmanim.shiur import get_halacha, get_rambam, get_tanya

# ---- city definitions (mirroring the VBA hereJ/hereT/hereH/hereB) ----------

JERUSALEM = Location(latitude=31.788, longitude=35.218, elevation=800)
TEL_AVIV = Location(latitude=32.06, longitude=34.77, elevation=20)
HAIFA = Location(latitude=32.8, longitude=34.991, elevation=300)
BEER_SHEVA = Location(latitude=31.24, longitude=34.79, elevation=0)

CITIES = [
    ("ירושלים", JERUSALEM, 40),
    ("תל אביב", TEL_AVIV, 22),
    ("חיפה", HAIFA, 30),
    ("באר שבע", BEER_SHEVA, 20),
]

# Named location presets for --location argument
LOCATION_PRESETS: dict[str, tuple[Location, int]] = {
    "jerusalem":  (JERUSALEM,  40),
    "tel-aviv":   (TEL_AVIV,   22),
    "haifa":      (HAIFA,      30),
    "beer-sheva": (BEER_SHEVA, 20),
    "eli": (Location(latitude=32.05, longitude=35.26, elevation=720), 30),
}

NUM_COLS = 7  # Sunday..Saturday
NUM_ROWS = 6  # max rows in a month grid

# Hebrew Unicode block: U+0590–U+05FF
_HEBREW_RANGE = range(0x0590, 0x0600)


# ---------------------------------------------------------------------------
#  RTL / reverse helpers
# ---------------------------------------------------------------------------

def _has_hebrew(s: str) -> bool:
    """True if the string contains any Hebrew-script character."""
    return any(ord(ch) in _HEBREW_RANGE for ch in s)


def _reverse_hebrew(text: str) -> str:
    """Reverse a mixed LTR/RTL string so it displays correctly on a
    pure-LTR terminal.

    Strategy: split into contiguous Hebrew vs non-Hebrew runs, reverse the
    order of those runs, AND reverse the characters within each Hebrew run.
    This makes the text fully readable left-to-right on terminals without
    bidi support.

    Example:
       "אבג hello דהו world"
    becomes:
       "world והוד hello גבא"
    """
    runs: list[str] = []
    if not text:
        return text

    current = text[0]
    in_hebrew = ord(current) in _HEBREW_RANGE
    for ch in text[1:]:
        ch_hebrew = ord(ch) in _HEBREW_RANGE
        if ch_hebrew == in_hebrew:
            current += ch
        else:
            runs.append(current)
            current = ch
            in_hebrew = ch_hebrew
    runs.append(current)

    # Reverse the order of runs AND reverse chars within each Hebrew run
    result: list[str] = []
    for run in reversed(runs):
        if any(ord(ch) in _HEBREW_RANGE for ch in run):
            result.append(run[::-1])  # reverse individual characters
        else:
            result.append(run)
    return "".join(result)


def _rev(text: str, *, enabled: bool) -> str:
    """Conditionally reverse Hebrew text."""
    if enabled and _has_hebrew(text):
        return _reverse_hebrew(text)
    return text


def _wrap(text: str, width: int) -> list[str]:
    """Wrap a string into multiple lines, each at most `width` chars.

    Splits on spaces and natural breakpoints to keep text readable.
    If a single token exceeds width it is hard-broken at the boundary.
    """
    if not text:
        return [""]

    # Normalize: insert break hints before common delimiter patterns so
    # the greedy fill can prefer breaking at those points.
    # We use a zero-width hint that the tokenizer will split on.
    # Replace: ]space, .space, :space, ;space  with token boundary
    text = re.sub(r'(\.)(\s)', r'\1\n\2', text)  # period + space
    text = re.sub(r'(\])(\s)', r'\1\n\2', text)  # bracket + space
    text = re.sub(r'(:)(\s)', r'\1\n\2', text)   # colon + space

    tokens = text.split()
    lines: list[str] = []
    current = ""
    for tok in tokens:
        # Hard-break very long tokens
        while len(tok) > width:
            if current:
                lines.append(current)
                current = ""
            lines.append(tok[:width])
            tok = tok[width:]
        candidate = current + (" " if current else "") + tok
        if len(candidate) <= width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = tok
    if current:
        lines.append(current)
    return lines or [""]


# ---------------------------------------------------------------------------
#  helpers
# ---------------------------------------------------------------------------

def _round_time(t: datetime) -> datetime:
    """Round to nearest minute (VBA tround semantics)."""
    return t + timedelta(minutes=1) if t.second > 29 else t


def _is_israel_dst(d: datetime) -> bool:
    """Return True if `d` falls within Israel DST (UTC+3).

    Israel DST: starts Friday before the last Sunday of March,
    ends the last Sunday of October (at 02:00 local time).
    """
    year = d.year

    # End: last Sunday of October
    oct31 = datetime(year, 10, 31)
    end_sunday = oct31 - timedelta(days=(oct31.weekday() + 1) % 7)
    end_dst = end_sunday.replace(hour=2, minute=0)

    # Start: Friday before last Sunday of March
    mar31 = datetime(year, 3, 31)
    last_sunday_mar = mar31 - timedelta(days=(mar31.weekday() + 1) % 7)
    start_dst = (last_sunday_mar - timedelta(days=2)).replace(hour=2, minute=0)

    return start_dst <= d < end_dst


def hebrew_today(*, is_dst: bool | None = None) -> HDate:
    """Return today's Hebrew date (Israel time, EY=True).

    is_dst=None means auto-detect Israel DST; True/False forces it.
    """
    now = datetime.now()
    if is_dst is None:
        is_dst = _is_israel_dst(now)
    h = convert_date(now)
    h.offset = 10800 if is_dst else 7200  # UTC+3 DST, UTC+2 standard
    h.EY = True
    return h


def get_day_info(hd: HDate, *, show_limud: bool = True,
                 location: tuple[Location, int] | None = None) -> list[str]:
    """Return day-info lines for a single date, VBA calendar_utils_get_day_info style.
    If show_limud is False, daily learning lines are omitted.
    If location is given, only candle-lighting/havdalah times for that location
    are shown; otherwise all 4 cities are shown."""
    lines: list[str] = []

    # 1. Hebrew date line
    lines.append(hdate_format(hd))

    # 2. Shabbos / Yom Tov title
    parsh = get_parshah(hd)
    if parsh != Parshah.NOPARSHAH:
        lines.append(f"שבת פרשת {parshah_format(parsh)}")
    else:
        ytov = get_yom_tov(hd)
        if ytov != YomTov.CHOL:
            lines.append(yom_tov_format(ytov))

    # Special shabbos note
    sp = get_special_shabbos(hd)
    if sp != YomTov.CHOL:
        lines.append(f"({yom_tov_format(sp)})")

    # Shabbos Mevorchim
    if is_shabbos_mevorchim(hd):
        lines.append("שבת מברכים")

    # 3. Sefirat HaOmer
    omer = get_omer(hd)
    if omer:
        lines.append(f"({omer} בעומר)")

    # 4. Rosh Chodesh
    hd_next = HDate(**hd.__dict__)
    hdate_add_day(hd_next, 1)
    if get_rosh_chodesh(hd) == YomTov.ROSH_CHODESH:
        if hd.day == 1:
            month_name = num_to_h_month(hd.month, hd.leap)
        else:
            month_name = num_to_h_month(hd_next.month, hd_next.leap)
        lines.append(f"ראש חודש {month_name}")

    # -- Sunrise / sunset (primary location, or all if none selected) --
    primary_loc, primary_offset = location if location else CITIES[0][1:]
    sr = _round_time(hdate_gregorian(getsunrise(hd, primary_loc)))
    ss = _round_time(hdate_gregorian(getsunset(hd, primary_loc)))
    lines.append(f"זריחה: {sr.strftime('%H:%M')}  שקיעה: {ss.strftime('%H:%M')}")

    # 5. Candle lighting / havdalah times
    # When location is selected, only show times for that location
    entries: list[tuple[str, Location, int]]
    if location:
        entries = [("", primary_loc, primary_offset)]
    else:
        entries = CITIES

    cl = is_candle_lighting(hd)
    if cl == 1:  # erev shabbos / yom tov (before shekiah)
        lines.append("הדלקת נרות:")
        for name, loc, mins in entries:
            t = _round_time(hdate_gregorian(getelevationsunset(hd, loc)))
            t -= timedelta(minutes=mins)
            prefix = f"  {name} " if name else "  "
            lines.append(f"{prefix}{t.strftime('%H:%M')}")
    elif cl == 2:  # candle lighting after tzais
        lines.append('הדלקת נרות (צה"כ):')
        for name, loc, _ in entries:
            t = _round_time(hdate_gregorian(gettzais8p5(hd, loc)))
            prefix = f"  {name} " if name else "  "
            lines.append(f"{prefix}{t.strftime('%H:%M')}")
    else:
        # Havdalah (shabbos / yom tov / yom kippur -> next day is chol)
        if is_assur_be_melachah(hd) and not is_assur_be_melachah(hd_next):
            yt = get_yom_tov(hd)
            if yt == YomTov.YOM_KIPPUR:
                label = 'מוצאי יוה"כ:'
            elif hd.wday == 0:
                label = "מוצאי שבת:"
            else:
                label = 'מוצאי יו"ט:'
            lines.append(label)
            for name, loc, _ in entries:
                t = _round_time(hdate_gregorian(gettzais8p5(hd, loc)))
                prefix = f"  {name} " if name else "  "
                lines.append(f"{prefix}{t.strftime('%H:%M')}")

    # 6. Daily learning (compact), separated by a blank line
    if show_limud:
        greg_date = hdate_gregorian(hd).date()
        daf = get_daf_yomi_format(greg_date)
        mishna = get_mishna_yomi_format(greg_date)
        halacha = get_halacha(hd)
        rambam = get_rambam(hd, daily_chapter=False)
        tanya = get_tanya(hd)
        has_limud = any([daf, mishna, halacha, rambam, tanya])
        if has_limud:
            lines.append("")  # blank separator
        if daf:
            lines.append(f"בבלי: {daf}")
        if mishna:
            lines.append(f"משנה: {mishna}")
        if halacha:
            lines.append(f"הלכה: {halacha}")
        if rambam:
            lines.append(f"רמב״ם: {rambam.replace(';', ', ')}")
        if tanya:
            lines.append(f"תניא: {tanya}")

    return lines


def build_month_grid(hd_seed: HDate, *, show_limud: bool = True,
                     location: tuple[Location, int] | None = None,
                     is_dst: bool | None = None,
                     loazi: bool = False) -> list[list[list[str]]]:
    """Return a 6x7 grid. Each cell is a list of info lines for that date.
    The grid always has exactly 42 cells (6 rows x 7 columns), padded with
    days from the previous/next month as needed, mirroring the VBA
    calendar_utils_get_month_matrix logic.

    If loazi=True, month/year are Gregorian and the title is formatted
    as a Gregorian month name; otherwise it's a Hebrew month.

    is_dst=None means auto-detect per-day Israel DST."""
    month = hd_seed.month  # 1-12 Gregorian, or 1-13 Hebrew
    year = hd_seed.year

    if loazi:
        # Gregorian month: simple DateSerial(year, month, 1)
        first_greg = datetime(year, month, 1, 12, 0, 0)
        # Last day of Gregorian month
        if month == 12:
            last_greg_date = datetime(year + 1, 1, 1).date() - timedelta(days=1)
        else:
            last_greg_date = datetime(year, month + 1, 1).date() - timedelta(days=1)
        last_day_num = last_greg_date.day
    else:
        # Hebrew month: build HDate, convert to Gregorian
        hd_first = hdate_new()
        hdate_add_year(hd_first, year - 1)
        if month < 7:
            hdate_add_year(hd_first, 1)
        hdate_add_month(hd_first, month)
        hdate_add_day(hd_first, 1)
        hd_first.offset = 7200
        hd_first.EY = True
        first_greg = hdate_gregorian(hd_first)
        last_day_num = last_day_of_hebrew_month(month, year)
        last_greg_date = first_greg.date() + timedelta(days=last_day_num - 1)

    # Column for the 1st (0=Sunday per VBA Weekday)
    start_col = (first_greg.weekday() + 1) % 7  # Python Mon=0,Sun=6 -> Sun=0..Sat=6

    # Build ordered list of (greg_date, in_month)
    cells: list[tuple[datetime, bool]] = []

    # Previous month tail (backfill from Sunday before 1st)
    prev_day = first_greg
    for i in range(start_col):
        prev_day -= timedelta(days=1)
        cells.insert(0, (prev_day, False))

    # Current month days
    cur = first_greg
    d = 0
    while d < last_day_num:
        cells.append((cur, True))
        cur += timedelta(days=1)
        d += 1

    # Next month head (fill to 42)
    while len(cells) < 42:
        cells.append((cur, False))
        cur += timedelta(days=1)

    # Convert to grid rows
    grid: list[list[list[str]]] = []
    row: list[list[str]] = []
    for i, (gdate, in_month) in enumerate(cells):
        hd = convert_date(datetime(gdate.year, gdate.month, gdate.day,
                                   hour=12, minute=0, second=0))
        hd.offset = 10800 if (is_dst if is_dst is not None
                            else _is_israel_dst(gdate)) else 7200
        hd.EY = True
        info = get_day_info(hd, show_limud=show_limud, location=location)
        # Prepend a marker line: Gregorian date + in/out-of-month
        marker = f"{gdate.strftime('%d/%m/%Y')} {'[IN]' if in_month else '[out]'}"
        info.insert(0, marker)
        row.append(info)
        if len(row) == 7:
            grid.append(row)
            row = []

    return grid


def print_grid(grid: list[list[list[str]]], *, reverse: bool = False,
               tile_width: int = 50) -> None:
    """Pretty-print the month grid with borders.
    Lines that exceed CONTENT_WIDTH are word-wrapped so cells stay aligned.
    When reverse=True, weekday columns run right-to-left (שבת .. ראשון)."""
    PAD = 1  # space on each side of cell content
    CONTENT_WIDTH = tile_width
    COL_WIDTH = CONTENT_WIDTH + 2 * PAD
    WEEKDAYS_HE = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"]

    # Reverse column order when --reverse
    if reverse:
        WEEKDAYS_HE = list(reversed(WEEKDAYS_HE))
        grid = [list(reversed(row)) for row in grid]

    # Pre-wrap all cell content
    wrapped_grid: list[list[list[str]]] = []
    for row in grid:
        wrapped_row: list[list[str]] = []
        for cell in row:
            flat: list[str] = []
            for line in cell:
                rline = _rev(line, enabled=reverse)
                flat.extend(_wrap(rline, CONTENT_WIDTH))
            wrapped_row.append(flat)
        wrapped_grid.append(wrapped_row)

    # Top border
    header_days = [_rev(d, enabled=reverse) for d in WEEKDAYS_HE]
    header = "│" + "│".join(f"{d:^{COL_WIDTH}}" for d in header_days) + "│"
    sep = "─" * len(header)
    print(sep)
    print(header)
    print(sep)

    for row in wrapped_grid:
        max_lines = max((len(cell) for cell in row), default=1)
        for line_idx in range(max_lines):
            parts = []
            for cell in row:
                if line_idx < len(cell):
                    line = cell[line_idx]
                    # Right-align when reversed (simulates RTL), left-align normally
                    if reverse:
                        parts.append(f"{' ' * PAD}{line:>{CONTENT_WIDTH}}{' ' * PAD}")
                    else:
                        parts.append(f"{' ' * PAD}{line:<{CONTENT_WIDTH}}{' ' * PAD}")
                else:
                    parts.append(" " * COL_WIDTH)
            print("│" + "│".join(parts) + "│")
        print(sep)


def print_molad(hd_seed: HDate, *, reverse: bool = False) -> None:
    """Print molad info for the month (and next month if relevant), mirroring
    calendar_utils_get_month_molad_info."""
    year = hd_seed.year
    month = hd_seed.month

    # Get ~15th of month to determine molad month
    hd_mid = HDate(**hd_seed.__dict__)
    hdate_add_day(hd_mid, 14)
    molad1 = get_molad(hd_mid.year, hd_mid.month)
    molad_str1 = molad_format(molad1)
    month1_name = num_to_h_month(hd_mid.month, hd_mid.leap)
    print(_rev(f"\nמולד חודש {month1_name}:", enabled=reverse))
    print(_rev(f"  {molad_str1}", enabled=reverse))

    # Check if next month's molad is relevant
    last_day = last_day_of_hebrew_month(month, year)
    hd_last = HDate(**hd_seed.__dict__)
    hdate_add_day(hd_last, last_day - 1)
    if (hd_last.month != hd_mid.month
            or get_rosh_chodesh(hd_last) == YomTov.ROSH_CHODESH
            or is_shabbos_mevorchim(hd_last)):
        if hd_last.month == hd_mid.month:
            hdate_add_month(hd_last, 1)
        molad2 = get_molad(hd_last.year, hd_last.month)
        molad_str2 = molad_format(molad2)
        month2_name = num_to_h_month(hd_last.month, hd_last.leap)
        print(_rev(f"\nמולד חודש {month2_name}:", enabled=reverse))
        print(_rev(f"  {molad_str2}", enabled=reverse))


def print_csv(grid: list[list[list[str]]], *, reverse: bool = False) -> None:
    """Output the grid as CSV. Each cell's info lines are joined by ";".
    First row is the weekday header."""
    import csv
    writer = csv.writer(sys.stdout, lineterminator="\n")
    WEEKDAYS_HE = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"]
    if reverse:
        WEEKDAYS_HE = list(reversed(WEEKDAYS_HE))
        grid = [list(reversed(row)) for row in grid]
    writer.writerow(WEEKDAYS_HE)
    for row in grid:
        csv_row = []
        for cell in row:
            # First line is the Gregorian marker; join all lines with ";"
            csv_row.append(";".join(cell))
        writer.writerow(csv_row)


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Jewish calendar month view")
    parser.add_argument("--month", type=int, help="Hebrew month (1=Nissan .. 13=Adar II)")
    parser.add_argument("--year", type=int, help="Hebrew year")
    parser.add_argument("--no-molad", action="store_true", help="Skip molad info")
    parser.add_argument("--no-limud", action="store_true", help="Skip daily learning in tiles")
    parser.add_argument("--tile-width", type=int, default=50,
                        help="Content width per tile in chars (default: 50)")
    parser.add_argument("--dst", nargs="?", const=True, default=None,
                        help="Force DST on (UTC+3) or off (UTC+2). "
                             "Default (no flag): auto-detect Israel DST. "
                             "Use --dst yes/on/1 to force on, --dst no/off/0 to force off")
    parser.add_argument("--location", type=str,
                        help="City preset (jerusalem, tel-aviv, haifa, beer-sheva) "
                             "or lat/lon/elv e.g. 31.78,35.22,800 or 31.78x35.22x800. "
                             "Offset (CL minutes before shekiah) defaults to 40; append e.g. ,40")
    parser.add_argument("--reverse", action="store_true",
                        help="Reverse Hebrew text for LTR-only terminals")
    parser.add_argument("--loazi", action="store_true",
                        help="Gregorian (non-Hebrew) calendar mode. "
                             "--month and --year refer to Gregorian dates")
    parser.add_argument("--csv", action="store_true",
                        help="Output as CSV instead of formatted grid")
    args = parser.parse_args()

    rev = args.reverse
    show_limud = not args.no_limud
    tile_width = args.tile_width

    # Parse --location: named preset or lat/lon/elv[,offset]
    location: tuple[Location, int] | None = None
    if args.location:
        preset = LOCATION_PRESETS.get(args.location)
        if preset:
            location = preset
        else:
            parts = re.split(r'[,x]', args.location)
            if len(parts) < 3:
                parser.error("--location must be a preset or lat,lon,elev[,offset]")
            try:
                lat, lon, elev = map(float, parts[:3])
            except ValueError:
                parser.error("lat/lon/elev must be numeric")
            offset = int(parts[3]) if len(parts) > 3 else 40
            location = (Location(latitude=lat, longitude=lon, elevation=elev), offset)

    # Parse --dst: auto (None), forced True, or forced False
    is_dst: bool | None = None
    if args.dst is not None:
        if args.dst is True:
            is_dst = True
        elif isinstance(args.dst, str) and args.dst.lower() in ("yes", "on", "1", "true"):
            is_dst = True
        elif isinstance(args.dst, str) and args.dst.lower() in ("no", "off", "0", "false"):
            is_dst = False
        else:
            is_dst = True  # bare --dst means force on

    loazi = args.loazi

    if loazi:
        # Gregorian mode: use today's Gregorian date as default
        now = datetime.now()
        if not args.year:
            args.year = now.year
        if not args.month:
            args.month = now.month
        from calendar import month_name as _gm
        title = f"  {_gm[args.month]} {args.year}"
        hd = HDate(year=args.year, month=args.month, day=1)  # placeholder for grid building
        title_he = _rev(title, enabled=rev) if rev else title
    else:
        hd = hebrew_today(is_dst=is_dst)
        if args.month:
            hd.month = args.month
        if args.year:
            hd.year = args.year
        month_name = num_to_h_month(hd.month, hd.leap)
        title_he = _rev(f"  {month_name} {hd.year}  (leap year: {bool(hd.leap)})",
                        enabled=rev)

    sep = "=" * 60
    if not args.csv:
        print(f"\n{sep}")
        print(title_he)
        print(sep)

    grid = build_month_grid(hd, show_limud=show_limud, location=location,
                            is_dst=is_dst, loazi=loazi)

    if args.csv:
        print_csv(grid, reverse=rev)
    else:
        print_grid(grid, reverse=rev, tile_width=tile_width)

    if not args.no_molad and not loazi and not args.csv:
        print_molad(hd, reverse=rev)