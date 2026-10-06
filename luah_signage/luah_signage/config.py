"""
Configuration for the luah_signage application.

Settings are loaded from a JSON file (see config.example.json) rather than
being hardcoded, unlike the original VBA (which hardcoded location, UTC
offset, and the pptx path directly in LuahMain.bas).
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field, replace
from pathlib import Path

from pyzmanim.noaa_calculator import Location

# config.py lives at <repo_root>/luah_signage/luah_signage/config.py, so
# config.json (repo-root-level, alongside pyzmanim-lib/ and reference/) is
# two directories up.
REPO_ROOT = Path(__file__).resolve().parent.parent.parent
DEFAULT_CONFIG_PATH = REPO_ROOT / "config.json"
# Default presentation file used when config.json's "pptx_path" is missing
# or empty - the sample/test deck shipped alongside launch.py at the repo
# root, so a bare-minimum config (or one with pptx_path accidentally left
# blank) still has something to load rather than failing outright.
DEFAULT_PPTX_PATH = REPO_ROOT / "luah.pptx"


@dataclass
class ClockStyle:
    """Proportions/appearance of the generated analog clock hands, expressed
    as fractions of the clock face radius (for length) or in millimeters
    (for line widths/pivot size).

    The *_length_ratio fields are the TIP length only (distance from the
    pivot/center out to the end of the hand) - a normal analog clock hand
    does not extend backward past the center. tail_length_ratio is a small
    shared stub extending in the opposite direction for visual balance
    (mostly hidden under the pivot dot); set to 0 for no tail at all."""

    hour_length_ratio: float = 0.5
    minute_length_ratio: float = 0.85
    second_length_ratio: float = 0.9
    tail_length_ratio: float = 0.12
    hour_width_mm: float = 2.5
    minute_width_mm: float = 1.8
    second_width_mm: float = 1.0
    pivot_radius_mm: float = 2.0
    hour_color: int = 0x000000
    minute_color: int = 0x000000
    second_color: int = 0xC00000
    # If False, no second hand is created at all (hour/minute hands only).
    # Useful for a calmer/less-distracting display, and it also means one
    # less shape needing a per-second UNO update (refresh_clock only needs
    # to run at minute granularity in that case - see app.py).
    show_second_hand: bool = True


@dataclass
class LuahConfig:
    pptx_path: Path
    location: Location
    eretz_yisroel: bool = True
    # IANA timezone name used to compute the local UTC offset (DST-aware,
    # unlike the original VBA which hardcoded a fixed +3h offset).
    timezone: str = "Asia/Jerusalem"
    clock_tick_seconds: float = 1.0
    content_tick_seconds: float = 30.0
    reload_poll_seconds: float = 5.0
    uno_host: str = "localhost"
    uno_port: int = 2002
    # Which display/monitor to run the fullscreen slideshow on (1-based,
    # matches com.sun.star.presentation.XPresentation2's Display property).
    display_number: int = 1
    clock_style: ClockStyle = field(default_factory=ClockStyle)
    # If True (the default, intended for unattended signage hardware), the
    # app tries to reconnect/relaunch LibreOffice and reload the
    # presentation whenever the display connection is lost for any reason
    # (crash, someone closing the window, etc.) instead of exiting - see
    # app.py's recovery loop. Set to False for debug/dev sessions (e.g.
    # when repeatedly closing LibreOffice manually to get back to a chat
    # window) so the app just exits cleanly instead of endlessly retrying.
    auto_recover: bool = True
    # After a slide's original auto-advance duration elapses, we restore its
    # transition properties and then wait this many seconds before actually
    # advancing - giving Impress's rendering pipeline time to pick up the
    # restored transition so it visibly plays on that advance (restoring and
    # advancing in the same tick was observed to skip the transition
    # entirely). See slide_advance.DynamicSlideAdvancer.
    transition_settle_seconds: float = 1.0
    # For TESTING ONLY: an ISO-8601 datetime string (e.g.
    # "2026-09-25T14:30:00") to use as a simulated wall-clock start time
    # instead of the real system clock - once set, the app's notion of
    # "now" starts at this value and advances at normal real-time speed
    # (1 real second = 1 simulated second) from the moment the app
    # launches. Lets you test date/time-dependent behavior (zmanim,
    # #CHOLONLY/#NONCHOLONLY slide switching, etc.) for any day/time
    # without changing the actual OS clock. Leave as None (the default)
    # for normal/production use - see time_source.py.
    mock_start_datetime: str | None = None
    # For TESTING/DEBUGGING ONLY: when True, prints verbose diagnostics to
    # stdout every time poll_slide_advance takes an action (entered/
    # prepare_advance/advance) and every time refresh_clock's active
    # slide index changes - showing the exact CurrentSlideIndex values,
    # computed next-slide indices, and which clocks were considered
    # active. Intended to help diagnose slide-transition/clock-timing bugs
    # without needing to guess blindly - leave False for normal/production
    # use (the extra prints add noise and minor overhead).
    debug_slide_advance: bool = False
    # For DEBUGGING ONLY: when True (the default, intended for normal/
    # production signage use), the Impress editing window is moved
    # off-screen and shrunk to 1x1px right after load (see
    # presentation._move_window_offscreen) so only the fullscreen
    # slideshow is visible. Set to False to leave that window in its
    # normal on-screen position/size - useful while debugging, since it
    # lets you see the editing view (e.g. to visually confirm shape
    # scanning/tag substitution) without hunting for an off-screen window.
    hide_editor_window: bool = True
    # For DEBUGGING ONLY: when True, the fullscreen slideshow is NEVER
    # started at all - only the Impress editing window is opened, so you
    # can inspect/scroll through the actual slides and shapes (hashtag
    # substitution, analog clock shapes, etc.) in a normal windowed view
    # without a fullscreen presentation covering the screen. Pair this
    # with hide_editor_window=False (otherwise there would be nothing
    # visible at all, since the editing window would ALSO be hidden
    # off-screen with no slideshow to show instead). Every other refresh
    # mechanism (clock ticks, hashtag/content refresh, #CMD cache,
    # file-change reload) keeps running completely normally - only the
    # slideshow window itself is skipped - so shape updates are still
    # visible live in the editing window as they happen. Leave False for
    # normal/production use.
    debug_skip_slideshow: bool = False
    # --- debug_repaint_* : DEBUGGING ONLY -----------------------------
    # Independent on/off toggles for each of the individual tweaks that
    # were added over time to avoid visible flashing/flickering in the
    # analog clock and tagged-text refreshes. Added because on SOME
    # machines (observed: certain x86 and ARM systems, but not others) the
    # clock hands stop visibly moving in the live slideshow window -
    # updates only become visible after navigating away from and back to
    # the slide - while other machines running the identical deck/config
    # work flawlessly. This strongly suggests a LibreOffice/graphics-
    # driver rendering-pipeline quirk that one of these anti-flicker
    # tweaks interacts badly with on certain systems, but which one is
    # unknown without being able to test directly on the affected
    # hardware. Each flag below lets you disable ONE tweak at a time (by
    # editing config.json and relaunching) to isolate which one is
    # responsible on a given machine - leave all at their defaults for
    # normal/production use, where the tweaks are known to help far more
    # often than they hurt.
    #
    # Controls whether refresh_clock/refresh_content wrap their shape
    # updates in document.lockControllers()/unlockControllers() (batches
    # multiple shape updates into a single repaint - see refresh_clock's
    # comment). Set to False to update shapes with no batching at all, one
    # UNO call at a time - if THIS is the culprit, disabling it should at
    # least restore live hand movement (likely reintroducing the flash
    # the batching was meant to prevent, but that's the whole point of
    # isolating the variable).
    debug_repaint_lock_controllers: bool = True
    # Controls the "only rewrite a clock hand's PolyPolygon if its angle
    # actually changed since last tick" optimization (see uno_shapes.py's
    # _refresh_hand). Set to False to unconditionally rewrite every
    # hand's PolyPolygon on every tick, even when the computed position
    # is identical to what's already there.
    debug_repaint_skip_unchanged_clock_writes: bool = True
    # Controls whether clocks on VISIBLE slides other than the one
    # currently being displayed are updated every tick at all (see
    # refresh_clock's comment on why this was added - freezing them
    # caused a visible "jump" the next time that slide became current).
    # Set to False to go back to updating ONLY the current slide's
    # clock(s), matching how an earlier version of this code behaved.
    debug_repaint_update_offscreen_clocks: bool = True
    # EXPERIMENTAL, default OFF: after a clock hand's PolyPolygon is
    # rewritten, also toggles that hand shape's Visible property off then
    # back on immediately - forcing it to be removed and re-inserted into
    # the render tree, which may force a stubborn rendering backend to
    # actually repaint it even if a plain property write alone doesn't
    # seem to. This is a heavier-handed workaround than the other toggles
    # above (not merely disabling an optimization, but adding a new one)
    # and may itself cause a very brief visible blink of the hand - only
    # try this if disabling the other debug_repaint_* toggles individually
    # doesn't resolve the frozen-hands symptom on a given machine.
    debug_repaint_nudge_shape: bool = False
    # Path/command-line for the "#CMD:<args>" tag (see tagging.py's module
    # docstring): a shell-like string (shlex-split) giving the command to
    # run - e.g. a direct executable path, OR an interpreter + script path
    # for script-based tools (e.g. "python C:/Intel/luahnew/mg_sync/
    # fetch_prayer_times.py"). Use FORWARD slashes in paths here, same
    # convention as pptx_path above - this string is parsed with Python's
    # shlex (shell-like quoting/whitespace rules), which treats a backslash
    # as an escape character and would otherwise mangle a literal Windows
    # path. Quote any path segment containing spaces (e.g. a quoted
    # "C:/Program Files/..." segment). Any args captured after the tag's
    # own ":<...>" suffix (e.g. "#CMD:<-city tzfat>" -> ["-city", "tzfat"])
    # are appended AFTER this prefix's own argv entries. Leave None (the
    # default) to leave #CMD unconfigured - its tag(s) will then render as
    # an empty string (not an error) until configured - see
    # tagging.py's module docstring for why #CMD always renders from a
    # cache rather than running inline.
    #
    # SECURITY NOTE: the configured command is run via subprocess with a
    # plain argv list (never shell=True / string concatenation), so shell
    # metacharacters typed into a shape's "#CMD:<...>" text (which, per
    # this whole project's design, ANY shul member can author via their
    # phone's PowerPoint app) can never be interpreted as shell syntax -
    # but whatever you point this at is still fully trusted to run with
    # whatever argv values a shape's author chooses to type; only point
    # this at a script/tool you're comfortable having anyone who can edit
    # the deck effectively invoke with arbitrary arguments.
    #
    # PERFORMANCE NOTE: #CMD does NOT run on the main content-refresh tick
    # (content_tick_seconds) at all - it only ever reads from an
    # in-memory cache populated by a SEPARATE, much less frequent tick
    # (see cmd_tick_seconds below). This is specifically to avoid
    # invoking a possibly slow/network-bound command every
    # content_tick_seconds (default 30s) - repeated requests that often,
    # forever, risks tripping rate limits/blocks on whatever remote
    # service the configured command might call. The cmd-refresh tick
    # ITSELF still runs synchronously/blocking (no background threads
    # anywhere in this app) for up to cmd_tag_timeout_seconds PER distinct
    # args-key found in the deck, so it still briefly stalls clock/slide-
    # advance updates when it fires - just far less often than every
    # content refresh. Keep the configured command reasonably fast
    # regardless.
    cmd_tag_executable: str | None = None
    # Timeout (seconds) for each #CMD invocation - see cmd_tag_executable's
    # docstring above. Exceeding it stores a "[#CMD: timed out ...]"
    # placeholder into the cache instead of hanging indefinitely.
    cmd_tag_timeout_seconds: float = 10.0
    # Extra argv entries applied to EVERY #CMD invocation, regardless of
    # which tag/args-key triggered it - e.g. a shared flag every call
    # should always get (an API key, a --cache-file override, etc.)
    # without having to repeat it in every single shape's "#CMD:<...>"
    # text. Final argv order is cmd_tag_executable's own prefix, THEN
    # these global args, THEN the tag's own args - so a tag's own args
    # are appended last and can override a global flag if the target
    # script's argument parser treats a later occurrence as winning (true
    # for most argparse-based tools).
    #
    # Accepts two forms in config.json:
    # - A JSON LIST of strings (RECOMMENDED) - each item is used VERBATIM
    #   as one argv entry, e.g. ["--sep", " - ", "--swap"] passes exactly
    #   those 3 argv entries, including the literal " - " with its
    #   spaces. NO shell-like quote/whitespace parsing is applied to list
    #   items at all - this is the only form that can express an
    #   argument value containing a literal quote character, and it
    #   avoids any ambiguity about whether quotes typed inside a JSON
    #   string would be stripped (they would NOT be - e.g. the JSON list
    #   item "--sep=' - '" would be passed to the target script with
    #   those single-quote characters literally included, almost
    #   certainly NOT what you want - use ["--sep", " - "] or
    #   ["--sep= - "] instead).
    # - A plain STRING (kept for convenience/backward compatibility) -
    #   parsed the same shell-like way as a tag's own ":<...>" args
    #   (whitespace-separated, quoted spans count as one arg, quotes
    #   stripped) - see tagging._parse_shell_like_args.
    # Defaults to [] (no extra args at all) either way.
    cmd_tag_global_args: list[str] | str = field(default_factory=list)
    # How often (seconds) the #CMD cache is actually refreshed by
    # re-running cmd_tag_executable for every distinct args-key found in
    # the current deck - see cmd_tag_executable's PERFORMANCE NOTE above
    # for why this is deliberately decoupled from content_tick_seconds
    # and defaults to a much longer interval (5 minutes). Has no effect
    # at all if no shape in the deck uses a #CMD tag.
    cmd_tick_seconds: float = 300.0


def _resolve_pptx_path(raw: str | None) -> Path:
    """Resolves config.json's "pptx_path" value to an actual Path, falling
    back to DEFAULT_PPTX_PATH (luah.pptx next to launch.py) if the raw
    value is missing/empty OR if it doesn't point to an actually-existing,
    readable file (e.g. a typo'd path, a path on a drive that isn't
    mounted on this particular machine, or a permissions issue) - rather
    than failing later with a much less obvious error from deep inside
    shutil.copy2/LibreOffice's own file-open call in presentation.py's
    load(). Path.is_file() returns False (rather than raising) for a
    missing path, a path that exists but isn't a regular file, AND for a
    permission error during the check itself - all three collapse to the
    same "fall back to the default" behavior here."""
    if not raw:
        return DEFAULT_PPTX_PATH
    path = Path(raw)
    if not path.is_file():
        return DEFAULT_PPTX_PATH
    return path


def load_config(path: Path | None = None) -> LuahConfig:
    path = path or DEFAULT_CONFIG_PATH
    if not path.exists():
        raise FileNotFoundError(
            f"Config file not found at {path}. Copy config.example.json to "
            "config.json next to it and edit it for your site."
        )
    data = json.loads(path.read_text(encoding="utf-8"))
    location = Location(
        latitude=data["location"]["latitude"],
        longitude=data["location"]["longitude"],
        elevation=data["location"]["elevation"],
    )
    clock_style_data = data.get("clock_style", {})
    return LuahConfig(
        pptx_path=_resolve_pptx_path(data.get("pptx_path")),
        location=location,
        eretz_yisroel=data.get("eretz_yisroel", True),
        timezone=data.get("timezone", "Asia/Jerusalem"),
        clock_tick_seconds=data.get("clock_tick_seconds", 1.0),
        content_tick_seconds=data.get("content_tick_seconds", 30.0),
        reload_poll_seconds=data.get("reload_poll_seconds", 5.0),
        uno_host=data.get("uno_host", "localhost"),
        uno_port=data.get("uno_port", 2002),
        display_number=data.get("display_number", 1),
        clock_style=ClockStyle(**clock_style_data),
        auto_recover=data.get("auto_recover", True),
        transition_settle_seconds=data.get("transition_settle_seconds", 1.0),
        mock_start_datetime=data.get("mock_start_datetime"),
        debug_slide_advance=data.get("debug_slide_advance", False),
        hide_editor_window=data.get("hide_editor_window", True),
        debug_skip_slideshow=data.get("debug_skip_slideshow", False),
        debug_repaint_lock_controllers=data.get("debug_repaint_lock_controllers", True),
        debug_repaint_skip_unchanged_clock_writes=data.get(
            "debug_repaint_skip_unchanged_clock_writes", True
        ),
        debug_repaint_update_offscreen_clocks=data.get(
            "debug_repaint_update_offscreen_clocks", True
        ),
        debug_repaint_nudge_shape=data.get("debug_repaint_nudge_shape", False),
        cmd_tag_executable=data.get("cmd_tag_executable"),
        cmd_tag_timeout_seconds=data.get("cmd_tag_timeout_seconds", 10.0),
        cmd_tag_global_args=data.get("cmd_tag_global_args", []),
        cmd_tick_seconds=data.get("cmd_tick_seconds", 300.0),
    )


def _strip_comment_keys(data: dict) -> dict:
    """Returns a copy of ``data`` with any key starting with ``_`` removed
    - lets a ``#CONFIG:{...}`` tag's JSON payload include comment-only
    entries (e.g. ``"_comment": "..."``) the same way config.json/
    config.example.json already do (see their own ``_comment_*`` keys),
    without those keys being mistaken for real override fields."""
    return {key: value for key, value in data.items() if not key.startswith("_")}


def parse_config_override(json_text: str) -> dict:
    """Parses a ``#CONFIG:{...}`` tag's JSON payload (see tagging.py's
    ``find_first_config_json``) into a plain dict of override keys/
    values, for use with ``apply_config_overrides`` below.

    Any top-level key starting with ``_`` (e.g. ``"_comment": "..."``) is
    silently dropped - see ``_strip_comment_keys`` - letting a deck
    author annotate their override JSON with comments without those keys
    being treated as (unknown, error-raising) override fields.

    Raises ``json.JSONDecodeError`` on malformed JSON, or ``ValueError``
    if the parsed value isn't a JSON object (e.g. a bare list or string) -
    callers should catch both and decide how to report the failure (the
    #CONFIG marker shape is always hidden and has no visible text of its
    own to show an inline error in, unlike #CMD - presentation.py reports
    failures via a console print instead)."""
    data = json.loads(json_text)
    if not isinstance(data, dict):
        raise ValueError(f"#CONFIG JSON must be an object, got {type(data).__name__}")
    return _strip_comment_keys(data)


def apply_config_overrides(config: LuahConfig, overrides: dict) -> LuahConfig:
    """Returns a NEW LuahConfig with ``overrides`` applied on top of
    ``config`` - used for a deck's ``#CONFIG:{...}`` tag (see
    presentation.py's ``_prepare_document``), letting a single
    shul-authored slide override one or more config.json values for that
    specific deck/session, WITHOUT ever touching the actual config.json
    file on disk. Works for both a single-key override (e.g.
    ``{"cmd_tick_seconds": 60}``) and a full-config override (every
    field at once) - there's no distinction in how either is handled,
    since this just applies whatever keys happen to be present.

    The ``location`` and ``clock_style`` sub-objects, if present in
    ``overrides``, are MERGED onto the EXISTING config's own nested
    dataclass (via ``dataclasses.replace``) rather than replaced outright
    - so e.g. ``{"clock_style": {"show_second_hand": false}}`` only
    changes that one clock_style field, keeping every other clock_style
    field (hour_color, pivot_radius_mm, etc.) at its current value,
    rather than resetting the whole sub-object to ClockStyle's bare
    defaults. Any ``_``-prefixed comment key WITHIN these nested dicts is
    also stripped before merging (same as top-level keys - see
    ``parse_config_override``).

    A ``pptx_path`` override is resolved through the same
    ``_resolve_pptx_path`` fallback logic used when loading config.json
    normally (falls back to ``DEFAULT_PPTX_PATH`` if the override path
    doesn't point to an existing file) - though note a mid-session
    pptx_path override only takes effect on the NEXT reload() (the
    currently-loading document is already open by the time #CONFIG tags
    are scanned).

    An unknown top-level key raises ``TypeError`` (via
    ``dataclasses.replace``, which rejects unexpected keyword arguments)
    - this is an intentional "fail loud" choice so a typo'd override key
    doesn't silently do nothing; the whole override is applied atomically
    (all keys or none) so a single bad key can't leave the config in a
    half-overridden state."""
    overrides = dict(overrides)  # don't mutate the caller's dict
    if "location" in overrides and isinstance(overrides["location"], dict):
        overrides["location"] = replace(config.location, **_strip_comment_keys(overrides["location"]))
    if "clock_style" in overrides and isinstance(overrides["clock_style"], dict):
        overrides["clock_style"] = replace(
            config.clock_style, **_strip_comment_keys(overrides["clock_style"])
        )
    if "pptx_path" in overrides:
        overrides["pptx_path"] = _resolve_pptx_path(overrides["pptx_path"])
    return replace(config, **overrides)
