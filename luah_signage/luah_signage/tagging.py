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
from dataclasses import dataclass
from datetime import timedelta
from typing import Callable, Iterable, Protocol

from pyzmanim import dafyomi, hdateformat, shiur, zmanim
from pyzmanim.hebrewcalendar import YomTov, Parshah, get_parshah, get_special_shabbos, get_yom_tov, hdate_gregorian

from .zman_context import ZmanContext

# Group 2 matches either "+18"/"-45" (plain minutes) or "+00:30"/"-01:15"
# (HH:MM duration) - see module docstring. Group 3 (only meaningful for
# #CMD) matches a ":<...>" argument-string suffix, e.g. "#CMD:<-city tzfat>"
# captures "-city tzfat" in group 3. The two suffix forms never collide:
# group 2 only matches if the character right after the tag name is +/-,
# group 3 only matches if it's a literal ":<" instead.
TAG_PATTERN = re.compile(r"#([A-Z][A-Z0-9]*)([+-]\d+:\d{2}|[+-]\d+)?(?::<([^>]*)>)?")

# Tag name handled specially in render_template (does not go through
# TAG_REGISTRY's uniform (ctx, offset_minutes) -> str signature, since it
# needs the raw :<args> string and the configured executable instead).
CMD_TAG_NAME = "CMD"

# Safety cap on how much of a #CMD command's stdout gets substituted into a
# shape - guards against a misbehaving/unexpectedly chatty script bloating
# a slide's text indefinitely.
_MAX_CMD_OUTPUT_CHARS = 4000


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
    docstring) - NOT from the main per-tick content refresh."""
    for key in arg_keys:
        cache[key] = _run_cmd_tag(cmd_executable, key, timeout_seconds, global_args)


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

# Tags handled specially inside render_template's _replace closure, rather
# than via TAG_REGISTRY's uniform (ctx, offset_minutes) -> str renderers.
_SPECIAL_TAGS = {CMD_TAG_NAME}


def contains_known_tag(text: str) -> bool:
    return any(
        m.group(1) in TAG_REGISTRY or m.group(1) in _SPECIAL_TAGS for m in TAG_PATTERN.finditer(text)
    )


def render_template(
    template: str,
    ctx: ZmanContext,
    cmd_cache: dict[str, str] | None = None,
) -> str:
    """Re-renders ``template`` (a shape's original tag-containing text)
    against ``ctx``. ``#CMD:<...>`` tags are looked up in ``cmd_cache``
    (populated separately/infrequently by ``refresh_cmd_cache`` - see
    module docstring) rather than executed here - a key not yet present
    in the cache (e.g. before the first cmd-refresh tick has run)
    substitutes as an empty string, not an error placeholder."""
    def _replace(match: re.Match) -> str:
        tag_name = match.group(1)
        if tag_name == CMD_TAG_NAME:
            key = match.group(3) or ""
            return (cmd_cache or {}).get(key, "")
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

    def refresh(self, ctx: ZmanContext, cmd_cache: dict[str, str] | None = None) -> None:
        self.shape.set_text(render_template(self.template, ctx, cmd_cache))


def scan_shapes_for_tags(shapes: list[TextShapeLike]) -> list[TrackedShape]:
    tracked = []
    for shape in shapes:
        text = shape.get_text()
        if contains_known_tag(text):
            tracked.append(TrackedShape(shape=shape, template=text))
    return tracked
