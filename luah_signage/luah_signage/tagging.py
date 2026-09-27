"""
Hashtag scanning & rendering.

PowerPoint/Impress shapes are authored with placeholder text containing tags
like ``#SUNRISE``, ``#PARSHA``, ``#DAFYOMI``, optionally followed by a
time offset suffix (e.g. ``#SUNSET-18`` for "18 minutes before sunset", or
``#SUNSET+00:30`` for "30 minutes after sunset"). This module scans shape
text for known tags and re-renders the original template text with
computed values substituted in, given a ZmanContext for "now".

Offset syntax: ``+``/``-`` followed by either a plain integer number of
minutes (``+18``, ``-45``), or an ``HH:MM`` duration (``+00:30``,
``-01:15``) for offsets larger than 59 minutes or when hours+minutes is
just clearer to author. Both forms are equivalent ways of specifying the
same total-minutes offset - ``+00:30`` and ``+30`` mean the same thing.

Unlike the original VBA (which mutated the shape's displayed text in place,
losing the template unless a separate ORG tag/property was stashed), this
implementation always keeps the original template text in memory (in
TrackedShape.template) and re-renders from that copy every time, so no
information is ever destroyed and re-scanning is never needed after the
initial scan.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import timedelta
from typing import Callable, Protocol

from pyzmanim import dafyomi, hdateformat, shiur, zmanim
from pyzmanim.hebrewcalendar import YomTov, Parshah, get_parshah, get_special_shabbos, get_yom_tov, hdate_gregorian

from .zman_context import ZmanContext

# Group 2 matches either "+18"/"-45" (plain minutes) or "+00:30"/"-01:15"
# (HH:MM duration) - see module docstring.
TAG_PATTERN = re.compile(r"#([A-Z][A-Z0-9]*)([+-]\d+:\d{2}|[+-]\d+)?")


def _parse_offset_minutes(offset_str: str | None) -> int:
    """Parses a tag's offset suffix (as captured by TAG_PATTERN's second
    group) into a signed total-minutes integer. Returns 0 if there's no
    offset. Accepts both "+18"/"-45" (plain minutes) and "+00:30"/"-01:15"
    (HH:MM duration) forms."""
    if not offset_str:
        return 0
    sign = -1 if offset_str[0] == "-" else 1
    body = offset_str[1:]
    if ":" in body:
        hours_str, minutes_str = body.split(":", 1)
        return sign * (int(hours_str) * 60 + int(minutes_str))
    return sign * int(body)


class TextShapeLike(Protocol):
    """Minimal interface a shape needs to support to be tag-scanned; both
    real UNO shapes (via uno_shapes.UnoTextShape) and test doubles satisfy
    this without any UNO dependency."""

    def get_text(self) -> str: ...
    def set_text(self, text: str) -> None: ...


def _fmt_time(hd, offset_minutes: int = 0) -> str:
    dt = hdate_gregorian(hd)
    if offset_minutes:
        dt = dt + timedelta(minutes=offset_minutes)
    return dt.strftime("%H:%M")


def _tag_parsha(ctx: ZmanContext, offset_minutes: int) -> str:
    parsh = get_parshah(ctx.shabbos)
    if parsh != Parshah.NOPARSHAH:
        result = f"פרשת {hdateformat.parshah_format(parsh)}"
    else:
        ytov = get_yom_tov(ctx.shabbos)
        result = hdateformat.yom_tov_format(ytov) if ytov != YomTov.CHOL else ""

    special = get_special_shabbos(ctx.shabbos)
    if special != YomTov.CHOL:
        result += f"\n({hdateformat.yom_tov_format(special)})"
    return result


def _tag_dayzmanim(ctx: ZmanContext, offset_minutes: int) -> str:
    return (
        f"זריחה: {_fmt_time(zmanim.getsunrise(ctx.now, ctx.location))}\n"
        f"שקיעה: {_fmt_time(zmanim.getelevationsunset(ctx.now, ctx.location))}\n"
        f"צאת הכוכבים: {_fmt_time(zmanim.gettzais8p5(ctx.now, ctx.location))}"
    )


def _tag_fullzmanim(ctx: ZmanContext, offset_minutes: int) -> str:
    lines = [
        ("עלות השחר (72 דק'):", zmanim.getalos72),
        ("זריחה:", zmanim.getsunrise),
        ('סוף זמן ק"ש (מג"א):', zmanim.getshmamga),
        ('סוף זמן ק"ש (גר"א):', zmanim.getshmagra),
        ('סוף זמן תפילה (מג"א):', zmanim.gettefilamga),
        ('סוף זמן תפילה (גר"א):', zmanim.gettefilagra),
        ("חצות היום:", zmanim.getchatzosgra),
        ("מנחה גדולה:", zmanim.getminchagedolagra),
        ("מנחה קטנה:", zmanim.getminchaketanagra),
        ("שקיעה:", zmanim.getelevationsunset),
        ("צאת הכוכבים:", zmanim.gettzais8p5),
    ]
    return "\n".join(f"{label} {_fmt_time(func(ctx.now, ctx.location))}" for label, func in lines)


def _tag_dafyomi(ctx: ZmanContext, offset_minutes: int) -> str:
    return "דף היום:\n" + dafyomi.get_daf_yomi_format(hdate_gregorian(ctx.now).date())


def _tag_limudyomi(ctx: ZmanContext, offset_minutes: int) -> str:
    rambam = shiur.get_rambam(ctx.now, True)
    parts = rambam.split(";", 2)
    if len(parts) == 3:
        rambam = f"{parts[0]} - {parts[1]}\n{parts[2]}"
    daf = dafyomi.get_daf_yomi_format(hdate_gregorian(ctx.now).date())
    return (
        f"דף היום:\n{daf}\n\n"
        f'רמב"ם היום:\n{rambam}\n\n'
        f"תהלים:\n{shiur.tehillim(ctx.now)}"
    )


# tag name -> callable(ctx, offset_minutes) -> rendered string
TAG_REGISTRY: dict[str, Callable[[ZmanContext, int], str]] = {
    "HEBDATE": lambda ctx, off: hdateformat.hdate_or_format(ctx.now, ctx.location),
    "DAYZMANIM": _tag_dayzmanim,
    "FULLZMANIM": _tag_fullzmanim,
    "DAFYOMI": _tag_dafyomi,
    "PARSHA": _tag_parsha,
    "SHABBOS": lambda ctx, off: _fmt_time(zmanim.getelevationsunset(ctx.erev_shabbos, ctx.location), off),
    "MOZASH": lambda ctx, off: _fmt_time(zmanim.gettzais8p5(ctx.shabbos, ctx.location), off),
    "LIMUDYOMI": _tag_limudyomi,
    "DICLOCK": lambda ctx, off: hdate_gregorian(ctx.now).strftime("%H:%M:%S"),
    "ALOS72": lambda ctx, off: _fmt_time(zmanim.getalos72(ctx.now, ctx.location), off),
    "SUNRISE": lambda ctx, off: _fmt_time(zmanim.getsunrise(ctx.now, ctx.location), off),
    "SHMAMGA": lambda ctx, off: _fmt_time(zmanim.getshmamga(ctx.now, ctx.location), off),
    "SHMAGRA": lambda ctx, off: _fmt_time(zmanim.getshmagra(ctx.now, ctx.location), off),
    "TFILAMGA": lambda ctx, off: _fmt_time(zmanim.gettefilamga(ctx.now, ctx.location), off),
    "TFILAGRA": lambda ctx, off: _fmt_time(zmanim.gettefilagra(ctx.now, ctx.location), off),
    "CHAZOT": lambda ctx, off: _fmt_time(zmanim.getchatzosgra(ctx.now, ctx.location), off),
    "BGMINHA": lambda ctx, off: _fmt_time(zmanim.getminchagedolagra(ctx.now, ctx.location), off),
    "LTMINHA": lambda ctx, off: _fmt_time(zmanim.getminchaketanagra(ctx.now, ctx.location), off),
    "SUNSET": lambda ctx, off: _fmt_time(zmanim.getelevationsunset(ctx.now, ctx.location), off),
    "TZAIS": lambda ctx, off: _fmt_time(zmanim.gettzais8p5(ctx.now, ctx.location), off),
}

# Tags that are handled elsewhere (analog clock shape generation, deferred
# multi-mode slide switching) rather than by generic text substitution.
NON_TEXT_TAGS = {"ANCLOCK", "MULTIMODE"}


def contains_known_tag(text: str) -> bool:
    return any(m.group(1) in TAG_REGISTRY for m in TAG_PATTERN.finditer(text))


def render_template(template: str, ctx: ZmanContext) -> str:
    def _replace(match: re.Match) -> str:
        tag_name = match.group(1)
        offset = _parse_offset_minutes(match.group(2))
        renderer = TAG_REGISTRY.get(tag_name)
        if renderer is None:
            return match.group(0)
        return renderer(ctx, offset)

    return TAG_PATTERN.sub(_replace, template)


@dataclass
class TrackedShape:
    """A shape whose text template contains one or more known tags, plus
    the original (unmodified) template text to re-render from on every
    refresh."""

    shape: TextShapeLike
    template: str

    def refresh(self, ctx: ZmanContext) -> None:
        self.shape.set_text(render_template(self.template, ctx))


def scan_shapes_for_tags(shapes: list[TextShapeLike]) -> list[TrackedShape]:
    tracked = []
    for shape in shapes:
        text = shape.get_text()
        if contains_known_tag(text):
            tracked.append(TrackedShape(shape=shape, template=text))
    return tracked
