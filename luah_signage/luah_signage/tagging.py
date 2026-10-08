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

``#CMD:<args>`` is a special tag (handled separately from the generic
TAG_REGISTRY-based tags above): it runs an external script/command
(configured once, in ``config.json``'s ``cmd_tag_executable``) and
substitutes the tag with that command's stdout. ``<args>`` (optional) is a
shell-like argument string - whitespace separates args, a single/double-
quoted span is one arg (may itself contain spaces) - parsed by
``_parse_shell_like_args`` and appended as extra argv entries after
``cmd_tag_executable``'s own (also shell-like-split) prefix. See that
config field's docstring in config.py for the full behavior/security notes.

``#CMD`` is deliberately NOT executed inline during ``render_template``
(unlike every other tag, which is cheap pure computation). Running an
external command/network call on every content-refresh tick
(``content_tick_seconds``, default every 30s) risked tripping rate
limits/blocks on whatever remote service the configured command might call
(e.g. ``mg_sync``'s mygabay.com fetch) - repeated requests every 30s,
forever, is exactly the kind of pattern that gets flagged/blocked. Instead,
``#CMD`` results are CACHED in a plain ``dict[str, str]`` keyed by the raw
``:<args>`` string (normalized: no-args and ``:<>``/missing both map to the
same ``""`` key) - ``render_template`` only ever READS this cache (never
runs the subprocess itself), defaulting to ``""`` for a key that hasn't
been resolved yet (e.g. right after startup, before the first cmd-refresh
tick has run). A SEPARATE, much-less-frequent tick
(``PresentationController.refresh_cmd_tags``, driven by
``config.cmd_tick_seconds`` - default 300s = 5 minutes - from app.py's main
loop) is what actually invokes ``refresh_cmd_cache`` to run the command
once per distinct args-key found in the deck and populate/update the
cache. This cleanly decouples "how often does text get redrawn" from "how
often does a slow/network-bound external command get invoked".

Unlike the original VBA (which mutated the shape's displayed text in place,
losing the template unless a separate ORG tag/property was stashed), this
implementation always keeps the original template text in memory (in
TrackedShape.template) and re-renders from that copy every time, so no
information is ever destroyed and re-scanning is never needed after the
initial scan.
"""
from __future__ import annotations

import os
import re
import shlex
import subprocess
from dataclasses import dataclass, field
from datetime import timedelta
from typing import Callable, Iterable, Protocol

from pyzmanim import dafyomi, hdateformat, shiur, zmanim
from pyzmanim.hdateformat import num_to_h_month
from pyzmanim.hebrewcalendar import (
    HDate,
    YomTov,
    Parshah,
    get_omer,
    get_parshah,
    get_rosh_chodesh,
    get_special_shabbos,
    get_yom_tov,
    hdate_add_day,
    hdate_gregorian,
    is_assur_be_melachah,
    is_candle_lighting,
    is_shabbos_mevorchim,
)

from .zman_context import ZmanContext

# Group 2 matches either "+18"/"-45" (plain minutes) or "+00:30"/"-01:15"
# (HH:MM duration) - see module docstring. Group 3 (only meaningful for
# #CMD) matches a ":<...>" argument-string suffix, e.g. "#CMD:<-city tzfat>"
# captures "-city tzfat" in group 3. Group 4 is a newline-formatting
# suffix: ":n" inserts a newline BEFORE the tag's rendered value, ":na"
# inserts one AFTER it (e.g. "#DAYINFO:n" / "#DAYINFO:na") - see
# _apply_newline_format. All three suffix forms never collide: group 2
# only matches if the character right after the tag name is +/-, group 3
# only matches if it's a literal ":<" instead, and group 4 only matches a
# literal ":n"/":na" followed by a non-word character or end-of-string
# (the trailing ``\b`` prevents it from misfiring on an unrelated
# ":name"-like suffix that merely happens to start with "n").
TAG_PATTERN = re.compile(
    r"#([A-Z][A-Z0-9]*)([+-]\d+:\d{2}|[+-]\d+)?(?::<([^>]*)>)?(?::(na|n)\b)?"
)

# Tag name handled specially in render_template (does not go through
# TAG_REGISTRY's uniform (ctx, offset_minutes) -> str signature, since it
# needs the raw :<args> string and the configured executable instead).
CMD_TAG_NAME = "CMD"

# Safety cap on how much of a #CMD command's stdout gets substituted into a
# shape - guards against a misbehaving/unexpectedly chatty script bloating
# a slide's text indefinitely.
_MAX_CMD_OUTPUT_CHARS = 4000

# Marker prefix for the "#CONFIG:{json}" tag (see
# presentation.py's _prepare_document docstring/module comments for the
# full design) - a shape whose text contains this prefix is a config-
# override marker: always hidden (like #HIDDEN), with everything after
# the prefix treated as a JSON object that overrides one or more
# config.json values for the current load/reload. NOT handled via
# TAG_PATTERN's regex at all (unlike #CMD's ":<...>" suffix) since JSON
# can itself contain any characters including "<"/">"/colons/braces -
# a simple substring search for this prefix, with everything after it
# taken as the JSON payload, is simpler and more robust here.
CONFIG_TAG_PREFIX = "#CONFIG:"


def find_first_config_json(texts: Iterable[str]) -> str | None:
    """Scans shape texts (in the given iteration order - callers should
    pass them in deck/document order) and returns the JSON payload (the
    substring after ``#CONFIG:``, stripped) of the FIRST text containing
    a ``#CONFIG:`` marker, or ``None`` if none exists.

    Per the feature's design, only the FIRST ``#CONFIG`` tag found across
    the whole deck is used - every subsequent ``#CONFIG``-tagged shape
    (if any) is silently ignored entirely (not even parsed) - a
    deliberate simplification so a deck author can't accidentally end up
    with ambiguous/conflicting config overrides from two different
    slides."""
    for text in texts:
        idx = text.find(CONFIG_TAG_PREFIX)
        if idx != -1:
            return text[idx + len(CONFIG_TAG_PREFIX):].strip()
    return None


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
    def set_text_range(self, start: int, end: int, text: str) -> None: ...


def _parse_shell_like_args(raw: str) -> list[str]:
    """Parses a ``#CMD:<...>`` argument string into a list of argv entries,
    shell-like: whitespace outside quotes separates args; a single- or
    double-quoted span (which may itself contain spaces) is ONE argument,
    with the quotes themselves stripped. A quote only starts a quoted span
    when it appears at the START of a token (not mid-token).

    Deliberately simpler than Python's stdlib ``shlex`` (no backslash-
    escape handling at all) - this is intentional: shlex's default POSIX
    mode treats ``\\`` as an escape character, which would silently mangle
    a literal Windows path typed between the ``<...>`` brackets (e.g.
    ``C:\\Users\\x``). Nothing here needs escaping for the kinds of simple
    args this tag is meant for; if an argument needs to contain a literal
    quote character, this parser does not support that."""
    args: list[str] = []
    i = 0
    n = len(raw)
    while i < n:
        ch = raw[i]
        if ch.isspace():
            i += 1
            continue
        if ch in ("'", '"'):
            quote = ch
            i += 1
            start = i
            while i < n and raw[i] != quote:
                i += 1
            args.append(raw[start:i])
            i += 1  # skip the closing quote (or just reach end-of-string)
            continue
        start = i
        while i < n and not raw[i].isspace() and raw[i] not in ("'", '"'):
            i += 1
        args.append(raw[start:i])
    return args


def _normalize_global_args(global_args) -> list[str]:
    """Normalizes ``cmd_tag_global_args`` (from ``config.json``) into a
    plain list of literal argv entries.

    Two accepted forms:
    - A JSON LIST of strings (the recommended form) - each item is used
      VERBATIM as one argv entry, with NO shell-like quote/whitespace
      parsing applied at all - e.g. ``["--sep", " - ", "--swap"]`` passes
      exactly those 3 argv entries, including the literal " - " with its
      spaces. This is the only form that can express an argument value
      containing a literal quote character, and avoids any ambiguity
      about whether quotes in a string would be stripped or not.
    - A plain STRING (kept for convenience/backward compatibility) -
      parsed the same shell-like way as a tag's own ``:<args>`` string
      (via ``_parse_shell_like_args``): whitespace-separated, with
      quoted spans treated as one arg and their quotes stripped.

    ``None``/``""``/an empty list all normalize to ``[]`` (no extra
    args)."""
    if not global_args:
        return []
    if isinstance(global_args, str):
        return _parse_shell_like_args(global_args)
    return [str(item) for item in global_args]


def _run_cmd_tag(
    cmd_executable: str | None,
    raw_args: str | None,
    timeout_seconds: float,
    global_args="",
) -> str:
    """Runs the configured ``cmd_tag_executable`` (see LuahConfig's
    docstring in config.py) with ``global_args`` (from ``config.json``'s
    ``cmd_tag_global_args`` - applied to EVERY #CMD invocation regardless
    of which tag/args-key triggered it; see ``_normalize_global_args`` for
    the accepted JSON-list-vs-string forms) followed by the
    ``#CMD:<...>`` tag's OWN parsed args appended, and returns the
    command's stdout (used as the cached value for that args-key - see
    module docstring for the cache-based design; this function is called
    from ``refresh_cmd_cache``, NOT from ``render_template`` directly).
    Final argv order: cmd_executable's own prefix, then global_args, then
    the tag's own args - so a tag's own args are appended LAST and can
    effectively override/extend a global flag if the target script's own
    argument parser treats a later occurrence of the same flag as
    winning (e.g. argparse does, for most flag types).

    Security note: cmd_executable's own shell-like-split prefix, the
    configured global_args, and the per-call tag args are ALL passed to
    subprocess.run as a single plain argv LIST (never ``shell=True``,
    never string-concatenated into a shell command line) - so shell
    metacharacters in a shape's ``#CMD:<...>`` text (which, per this
    project's whole design, ANY shul member can author via their phone's
    PowerPoint app) can never be interpreted as shell syntax/command
    injection. The configured executable itself is still responsible for
    safely handling whatever argv values it receives - this only
    protects against shell injection at the OS-process-launch boundary,
    not against anything the target script itself might do with its own
    arguments.

    Never raises - any failure (missing config, bad executable, timeout,
    non-zero exit, OS error) is caught and turned into a short, visible
    ``[#CMD: ...]`` placeholder string instead, so a misconfigured/failing
    command shows up clearly on the actual signage display rather than
    crashing the refresh tick or silently leaving stale text."""
    if not cmd_executable or not cmd_executable.strip():
        return "[#CMD: no cmd_tag_executable configured in config.json]"
    try:
        base_argv = shlex.split(cmd_executable)
    except ValueError as exc:
        return f"[#CMD: invalid cmd_tag_executable config: {exc}]"
    if not base_argv:
        return "[#CMD: cmd_tag_executable is empty]"

    argv = base_argv + _normalize_global_args(global_args) + _parse_shell_like_args(raw_args or "")
    # Force the child process's own stdout/stderr text encoding to UTF-8
    # regardless of the OS console code page - without this, a Python (or
    # similar) script's own print() calls can be silently mis-encoded on
    # Windows (observed: Hebrew text coming back as mojibake "?"/replacement
    # characters even though we decode the captured bytes as UTF-8 below -
    # the CHILD process itself wrote them in a different encoding to begin
    # with). Same fix already used for this exact class of problem in
    # launch.py/app.py elsewhere in this codebase. Harmless no-op for
    # non-Python commands (PYTHONIOENCODING is Python-interpreter-specific).
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8:backslashreplace"
    try:
        result = subprocess.run(
            argv,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
    except FileNotFoundError:
        return f"[#CMD: executable not found: {argv[0]!r}]"
    except subprocess.TimeoutExpired:
        return f"[#CMD: timed out after {timeout_seconds:.0f}s]"
    except OSError as exc:
        return f"[#CMD: error launching command: {exc}]"

    if result.returncode != 0:
        stderr_preview = (result.stderr or "").strip()[:200]
        suffix = f": {stderr_preview}" if stderr_preview else ""
        return f"[#CMD: exit code {result.returncode}{suffix}]"

    output = (result.stdout or "").strip()
    if len(output) > _MAX_CMD_OUTPUT_CHARS:
        output = output[:_MAX_CMD_OUTPUT_CHARS] + "\u2026(truncated)"
    return output


def find_cmd_arg_keys(templates: Iterable[str]) -> set[str]:
    """Scans a set of shape templates for every ``#CMD:<...>`` occurrence
    and returns the set of distinct, normalized args-keys found (bare
    ``#CMD`` and ``#CMD:<>`` both normalize to the key ``""``). Used by
    ``PresentationController._prepare_document`` to know which args-keys
    ``refresh_cmd_cache`` needs to resolve for the currently-loaded deck."""
    keys: set[str] = set()
    for template in templates:
        for match in TAG_PATTERN.finditer(template):
            if match.group(1) == CMD_TAG_NAME:
                keys.add(match.group(3) or "")
    return keys


def _is_cmd_error_placeholder(value: str) -> bool:
    """True if ``value`` is one of ``_run_cmd_tag``'s own error-placeholder
    strings (e.g. ``"[#CMD: timed out after 10s]"``) rather than genuine
    command output - all of them share this same ``"[#CMD: "`` prefix (see
    every ``return`` in ``_run_cmd_tag``). Used by ``refresh_cmd_cache`` to
    decide whether a failed run should overwrite a previously-cached GOOD
    value or not."""
    return value.startswith("[#CMD: ")


def refresh_cmd_cache(
    cache: dict[str, str],
    cmd_executable: str | None,
    arg_keys: Iterable[str],
    timeout_seconds: float = 10.0,
    global_args="",
) -> None:
    """Runs the configured command ONCE per distinct args-key in
    ``arg_keys`` (as found by ``find_cmd_arg_keys``) and stores each
    result into ``cache`` (mutated in place, keyed by the same args-key
    strings that ``render_template``'s ``#CMD`` handling looks up).
    ``global_args`` (a JSON list of literal argv entries, or a shell-like
    string - see ``_normalize_global_args``) is applied identically to
    every key. Meant to be called on its own, infrequent tick (see module
    docstring) - NOT from the main per-tick content refresh.

    If a run FAILS (returns one of ``_run_cmd_tag``'s own error-placeholder
    strings - see ``_is_cmd_error_placeholder``) AND ``cache`` already has
    a previously-cached GOOD (non-error) value for that same key, the OLD
    value is kept as-is rather than being overwritten with the new error
    placeholder - a transient failure (e.g. the target server being
    briefly unreachable, a one-off timeout) then just means the display
    keeps showing the last successfully-fetched data until the NEXT
    successful run, instead of replacing valid-looking content with a
    visible ``"[#CMD: ...]"`` error message. A key that has NEVER
    succeeded yet (not in ``cache``, or itself already holding an error
    placeholder from a previous failed attempt) still gets the new error
    placeholder stored as before, so a persistently-broken configuration
    remains visibly diagnosable on the display rather than silently
    staying blank forever."""
    for key in arg_keys:
        result = _run_cmd_tag(cmd_executable, key, timeout_seconds, global_args)
        if _is_cmd_error_placeholder(result):
            previous = cache.get(key)
            if previous is not None and not _is_cmd_error_placeholder(previous):
                continue  # keep the last known-good value instead of overwriting it
        cache[key] = result


def _fmt_time(hd, offset_minutes: int = 0, minute_ceil: bool = False) -> str:
    dt = hdate_gregorian(hd)
    if offset_minutes:
        dt = dt + timedelta(minutes=offset_minutes)
    if minute_ceil and dt.second >= 2:
        dt = dt.replace(second=0, microsecond=0) + timedelta(minutes=1)
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
        result += f" - {hdateformat.yom_tov_format(special)}"
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
        rambam = f"{parts[0]} - {parts[1]} - {parts[2]}"
    daf_bavli = dafyomi.get_daf_yomi_format(hdate_gregorian(ctx.now).date())
    daf_yerushalmi = dafyomi.get_daf_yomi_format(hdate_gregorian(ctx.now).date(),True)
    mishna_yomi = dafyomi.get_mishna_yomi_format(hdate_gregorian(ctx.now).date())
    halacha = shiur.get_halacha(ctx.now)

    return (
        f"בבלי: {daf_bavli}\n"
        f"ירושלמי: {daf_yerushalmi}\n"
        f"משניות: {mishna_yomi}\n"
        f"הלכה: {halacha}\n"
        f'רמב"ם: {rambam}\n'
        f"תהילים: {shiur.tehillim(ctx.now)}"
    )


def _tag_mozashtitle(ctx: ZmanContext, offset_minutes: int) -> str:
    """Returns the motzaei-label (e.g. "מוצאי שבת"/"מוצאי יו\"ט"/"מוצאי
    יוה\"כ") for whichever day it actually applies to - if TODAY is a
    shabbos/yom-tov/yom-kippur whose next day is chol, the label is based
    on today; otherwise it falls back to ``ctx.shabbos`` (the upcoming/
    current shabbos) instead. Unlike its name might suggest, this NEVER
    returns an empty string - one of the three labels is always returned."""
    hd = ctx.now
    hd_next = HDate(**hd.__dict__)
    hdate_add_day(hd_next, 1)
    if not (is_assur_be_melachah(hd) and not is_assur_be_melachah(hd_next)):
        hd = ctx.shabbos
    yt = get_yom_tov(hd)
    if yt == YomTov.YOM_KIPPUR:
        return 'מוצאי יוה"כ'
    elif hd.wday == 0:
        return "מוצאי שבת"
    else:
        return 'מוצאי יו"ט'


def _tag_dayinfo(ctx: ZmanContext, offset_minutes: int) -> str:
    """Return extra day details (yom-tov title on non-parshah days, omer,
    rosh chodesh, candle-lighting/havdalah, daily learning) — everything
    ``get_day_info`` provides EXCEPT the Hebrew date and the parshah
    title itself (which have their own dedicated tags ``#HEBDATE`` and
    ``#PARSHA``)."""
    lines: list[str] = []

    hd = ctx.now
    # Build a copy of hd for "next day" (used by Rosh Chodesh & havdalah)
    hd_next = HDate(**hd.__dict__)
    hdate_add_day(hd_next, 1)

    # -- Yom Tov / moed title (only on a NON-parshah day - i.e. a weekday
    # yom tov like Pesach/Sukkos/Rosh Hashanah, NOT an ordinary Shabbos,
    # which already gets its own title from #PARSHA) --
    if get_parshah(hd) == Parshah.NOPARSHAH:
        ytov = get_yom_tov(hd)
        if ytov != YomTov.CHOL:
            lines.append(hdateformat.yom_tov_format(ytov))

    # -- Special shabbos note --
    sp = get_special_shabbos(hd)
    if sp != YomTov.CHOL:
        lines.append(f"{hdateformat.yom_tov_format(sp)}")

    # -- Shabbos Mevorchim --
    if is_shabbos_mevorchim(hd):
        lines.append("שבת מברכים")

    # -- Sefirat HaOmer --
    omer = get_omer(hd)
    if omer:
        lines.append(f"{omer} בעומר")

    # -- Rosh Chodesh --
    if get_rosh_chodesh(hd) == YomTov.ROSH_CHODESH:
        if hd.day == 1:
            month_name = num_to_h_month(hd.month, hd.leap)
        else:
            month_name = num_to_h_month(hd_next.month, hd_next.leap)
        lines.append(f"ראש חודש {month_name}")

    return " - ".join(lines)


# tag name -> callable(ctx, offset_minutes) -> rendered string
TAG_REGISTRY: dict[str, Callable[[ZmanContext, int], str]] = {
    "HEBDATE": lambda ctx, off: hdateformat.hdate_or_format(ctx.now, ctx.location),
    "DAYZMANIM": _tag_dayzmanim,
    "FULLZMANIM": _tag_fullzmanim,
    "DAFYOMI": _tag_dafyomi,
    "DAYINFO": _tag_dayinfo,
    "PARSHA": _tag_parsha,
    "SHABBOS": lambda ctx, off: _fmt_time(zmanim.getelevationsunset(ctx.erev_shabbos, ctx.location), off),
    "MOZASH": lambda ctx, off: _fmt_time(zmanim.gettzais8p5(ctx.shabbos, ctx.location), off),
    "MOZASHTITLE": _tag_mozashtitle,
    "LIMUDYOMI": _tag_limudyomi,
    "DAFYOMIBV": lambda ctx, off: dafyomi.get_daf_yomi_format(hdate_gregorian(ctx.now).date()),
    "DAFYOMIYR": lambda ctx, off: dafyomi.get_daf_yomi_format(hdate_gregorian(ctx.now).date(), True),
    "MISHNA": lambda ctx, off: dafyomi.get_mishna_yomi_format(hdate_gregorian(ctx.now).date()),
    "HALACHA": lambda ctx, off: shiur.get_halacha(ctx.now),
    "RAMBAM": lambda ctx, off: shiur.get_rambam(ctx.now, daily_chapter=False).replace(";", ", "),
    "TEHILIM": lambda ctx, off: shiur.tehillim(ctx.now),
    "TANYA": lambda ctx, off: shiur.get_tanya(ctx.now),
    "DICLOCK": lambda ctx, off: hdate_gregorian(ctx.now).strftime("%H:%M"),
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
    "TZAISYESHIVA": lambda ctx, off: _fmt_time(zmanim.gettzaisyeshiva(ctx.now, ctx.location, True), off, True),
}

# Tags that are handled elsewhere (analog clock shape generation, deferred
# multi-mode slide switching) rather than by generic text substitution.
NON_TEXT_TAGS = {"ANCLOCK", "MULTIMODE"}

# Tags handled specially inside render_template's _replace closure, rather
# than via TAG_REGISTRY's uniform (ctx, offset_minutes) -> str renderers.
_SPECIAL_TAGS = {CMD_TAG_NAME}


def contains_known_tag(text: str) -> bool:
    return any(
        m.group(1) in TAG_REGISTRY or m.group(1) in _SPECIAL_TAGS for m in TAG_PATTERN.finditer(text)
    )


def _apply_newline_format(value: str, newline_flag: str | None) -> str:
    """Applies a tag's optional ":n"/":na" newline-formatting suffix (see
    TAG_PATTERN's group 4) to its already-rendered ``value``: ``"n"``
    prepends a newline, ``"na"`` appends one, and ``None`` (no suffix)
    leaves it unchanged. A no-op whenever ``value`` is empty, so a
    conditionally-empty tag (e.g. ``#DAYINFO`` on a plain weekday with
    nothing to report) never leaves a stray blank line behind."""
    if not value or not newline_flag:
        return value
    if newline_flag == "n":
        return "\n" + value
    return value + "\n"  # "na"


def _render_tag_match(match: re.Match, ctx: ZmanContext, cmd_cache: dict[str, str] | None) -> str:
    """Computes the rendered replacement for a single TAG_PATTERN match -
    the shared per-tag logic used by both ``render_template`` (whole-string
    substitution) and ``TrackedShape.refresh`` (formatting-preserving
    partial update - see its docstring)."""
    tag_name = match.group(1)
    if tag_name == CMD_TAG_NAME:
        key = match.group(3) or ""
        value = (cmd_cache or {}).get(key, "")
    else:
        offset = _parse_offset_minutes(match.group(2))
        renderer = TAG_REGISTRY.get(tag_name)
        if renderer is None:
            return match.group(0)
        value = renderer(ctx, offset)
    return _apply_newline_format(value, match.group(4))


def render_template(
    template: str,
    ctx: ZmanContext,
    cmd_cache: dict[str, str] | None = None,
) -> str:
    """Re-renders ``template`` (a shape's original tag-containing text)
    against ``ctx``, returning the resulting FULL string. ``#CMD:<...>``
    tags are looked up in ``cmd_cache`` (populated separately/infrequently
    by ``refresh_cmd_cache`` - see module docstring) rather than executed
    here - a key not yet present in the cache (e.g. before the first
    cmd-refresh tick has run) substitutes as an empty string, not an error
    placeholder.

    Note: this returns a plain string with no formatting information -
    it's used for tests/standalone rendering and by ``find_cmd_arg_keys``-
    style callers. ``TrackedShape.refresh`` does NOT call this for live
    shape updates (see its own docstring for why) - the two must stay
    behaviorally equivalent in terms of WHAT each tag renders to, via the
    shared ``_render_tag_match`` helper above."""
    def _replace(match: re.Match) -> str:
        return _render_tag_match(match, ctx, cmd_cache)

    return TAG_PATTERN.sub(_replace, template)


def _split_template(template: str) -> tuple[list[str], list[re.Match]]:
    """Splits ``template`` into alternating literal/static segments and
    tag matches: ``statics[0] + matches[0].group(0) + statics[1] +
    matches[1].group(0) + ... + statics[-1]`` reconstructs ``template``
    exactly (``len(statics) == len(matches) + 1``)."""
    statics: list[str] = []
    matches: list[re.Match] = []
    pos = 0
    for m in TAG_PATTERN.finditer(template):
        statics.append(template[pos:m.start()])
        matches.append(m)
        pos = m.end()
    statics.append(template[pos:])
    return statics, matches


@dataclass
class TrackedShape:
    """A shape whose text template contains one or more known tags, plus
    the original (unmodified) template text to re-render from on every
    refresh.

    ``refresh`` deliberately does NOT call ``shape.set_text(whole_string)``
    (a bulk ``setString`` on the whole shape) - doing so replaces every
    text run/portion with a single new run using one uniform formatting
    context, silently discarding any per-run formatting (bold/italic/
    color/etc.) a deck author applied to parts of the shape's text that
    aren't themselves a tag (e.g. a bold "בבלי:" label next to a
    non-bold "#DAFYOMIBV" tag). Instead, each tag's rendered value is
    updated IN PLACE via ``shape.set_text_range(start, end, text)`` -
    touching ONLY the characters that actually need to change, which for
    a real UNO shape means replacing the content of an existing text
    cursor/range rather than the whole shape, preserving that range's
    (and every other untouched run's) own character formatting.

    Position bookkeeping: the live shape's actual text is NEVER re-read
    during ``refresh`` - positions are tracked purely arithmetically,
    since the template's static segments never change and each tag's
    CURRENT rendered length is already known from the previous refresh
    (or, on the very first refresh, equals the length of the tag's own
    literal matched text, e.g. ``"#SUNRISE-18"``, since the live shape
    text still literally equals ``template`` at that point)."""

    shape: TextShapeLike
    template: str
    _static_segments: list[str] = field(init=False, repr=False, compare=False)
    _tag_matches: list[re.Match] = field(init=False, repr=False, compare=False)
    _current_values: list[str] | None = field(default=None, init=False, repr=False, compare=False)

    def __post_init__(self) -> None:
        self._static_segments, self._tag_matches = _split_template(self.template)

    def refresh(self, ctx: ZmanContext, cmd_cache: dict[str, str] | None = None, nudge: bool = False) -> None:
        """Re-renders every tag in this shape's template against ``ctx``
        (and ``cmd_cache``, for any #CMD tag), updating only the ranges
        whose rendered value actually changed since the last refresh (see
        this class's own docstring for why only changed ranges are
        touched at all).

        ``nudge=True`` (see LuahConfig.debug_repaint_nudge_text_shapes and
        uno_shapes._nudge_position's docstring) additionally nudges the
        shape's Position by a sub-pixel amount and immediately back right
        after a changed range's text is written - an experimental
        workaround for a LibreOffice slideshow rendering quirk where a
        shape's text genuinely changes in the document model but the live
        slideshow view never visibly repaints it (confirmed to
        specifically affect a #DICLOCK tag on a slide with no analog
        clock shape providing an incidental repaint via its own
        per-second property writes). A no-op (never calls
        ``self.shape.nudge()``) on a refresh where NOTHING in this shape
        actually changed, and silently does nothing at all if
        ``self.shape`` has no ``nudge`` method (e.g. a plain test double)
        - only real UNO shapes (uno_shapes.UnoTextShape) are expected to
        provide one."""
        if not self._tag_matches:
            return
        if self._current_values is None:
            # First refresh: the live shape text still literally equals
            # `self.template` (nothing substituted yet), so each tag's
            # CURRENT text is its own raw matched substring.
            self._current_values = [m.group(0) for m in self._tag_matches]

        changed = False
        offset = 0
        for i, match in enumerate(self._tag_matches):
            offset += len(self._static_segments[i])
            old_value = self._current_values[i]
            new_value = _render_tag_match(match, ctx, cmd_cache)
            if new_value != old_value:
                self.shape.set_text_range(offset, offset + len(old_value), new_value)
                self._current_values[i] = new_value
                changed = True
            offset += len(self._current_values[i])

        if changed and nudge:
            nudge_method = getattr(self.shape, "nudge", None)
            if nudge_method is not None:
                nudge_method()


def scan_shapes_for_tags(shapes: list[TextShapeLike]) -> list[TrackedShape]:
    tracked = []
    for shape in shapes:
        text = shape.get_text()
        if contains_known_tag(text):
            tracked.append(TrackedShape(shape=shape, template=text))
    return tracked
