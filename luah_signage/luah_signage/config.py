"""
Configuration for the luah_signage application.

Settings are loaded from a JSON file (see config.example.json) rather than
being hardcoded, unlike the original VBA (which hardcoded location, UTC
offset, and the pptx path directly in LuahMain.bas).
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

from pyzmanim.noaa_calculator import Location

# config.py lives at <repo_root>/luah_signage/luah_signage/config.py, so
# config.json (repo-root-level, alongside pyzmanim-lib/ and reference/) is
# two directories up.
DEFAULT_CONFIG_PATH = Path(__file__).resolve().parent.parent.parent / "config.json"


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
        pptx_path=Path(data["pptx_path"]),
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
    )
