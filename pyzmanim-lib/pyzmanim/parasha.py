"""
Weekly Torah portion (parshah) lookup table.

Ported from mod_parasha.bas in https://github.com/diaphone1/vbzmanim.

Static table mapping each of the 16 possible Hebrew year "types" (see
`hebrewcalendar.get_year_type`) to which Torah portion is read on each of the
~55 Shabbosos of the year (handling combined double-portions vs split single
portions, and Israel vs Diaspora differences implicitly via which row is
selected).

The raw data (parashaList, a 17x56 grid of enum-name strings) was extracted
from the original VBA source by tools/generate_data.py into
pyzmanim/_data/parasha_data.py; this module just resolves those name strings
into `Parshah` enum members once, at import time.
"""
from __future__ import annotations

from pyzmanim._data.parasha_data import PARASHA_LIST as _PARASHA_LIST_NAMES


def _build_parasha_list() -> list[list["Parshah"]]:
    from pyzmanim.hebrewcalendar import Parshah

    return [[Parshah[name] for name in row] for row in _PARASHA_LIST_NAMES]


PARASHA_LIST: list[list["Parshah"]] = _build_parasha_list()
