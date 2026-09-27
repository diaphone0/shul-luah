"""
UNO-specific shape adapters: wraps raw UNO draw shapes so tagging.py's
generic (UNO-agnostic) logic can operate on them, plus the analog clock
shape builder/updater.

Not testable in this dev environment (no LibreOffice installed here) -
written carefully against the documented UNO Draw/Impress API but should be
smoke-tested on the actual signage machine.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Iterator

from . import clock_geometry
from .clock_geometry import Point, Rect
from .config import ClockStyle

MM_PER_100 = 100  # UNO measurements are in 1/100 mm


class UnoTextShape:
    """Adapts a raw UNO shape to the TextShapeLike protocol expected by
    tagging.py (get_text/set_text)."""

    def __init__(self, shape):
        self._shape = shape

    @property
    def raw(self):
        return self._shape

    def get_text(self) -> str:
        return self._shape.getString()

    def set_text(self, text: str) -> None:
        self._shape.setString(text)


def is_text_shape(shape) -> bool:
    try:
        return shape.supportsService("com.sun.star.drawing.Text")
    except Exception:
        return False


def iter_all_shapes(document) -> Iterator[tuple[int, object, object]]:
    """Yields (slide_index, slide, shape) for every shape on every
    (non-master) slide of the document, in slide order."""
    slides = document.DrawPages
    for i in range(slides.Count):
        slide = slides.getByIndex(i)
        for j in range(slide.Count):
            yield i, slide, slide.getByIndex(j)


@dataclass
class AnalogClock:
    center: Point
    radius: float
    hour: object
    minute: object
    second: object | None
    pivot: object
    placeholder: object
    hour_tip_length: float
    minute_tip_length: float
    second_tip_length: float
    tail_length: float
    # Which slide (by index) this clock lives on - used by
    # presentation.py's refresh_clock to update ONLY the clock(s) on the
    # CURRENTLY DISPLAYED slide each tick, rather than every clock in the
    # deck regardless of slide. Updating an off-screen/not-currently-shown
    # slide's clock hand shapes was found to visibly interfere with that
    # slide's pending/just-restored transition effect (see repo memory) -
    # restricting updates to the active slide only avoids touching any
    # shape on a slide whose transition state might currently be "live"
    # ahead of an upcoming advance.
    slide_index: int = -1
    # Last angle actually rendered for each hand (degrees, clockwise from
    # 12), used by clock_geometry.stabilize_angle to suppress a hand
    # visibly regressing due to a small backward wall-clock time step -
    # see that function's docstring. None until the first render.
    last_hour_angle: float | None = None
    last_minute_angle: float | None = None
    last_second_angle: float | None = None


def _make_point(x: float, y: float):
    import uno
    from com.sun.star.awt import Point as AwtPoint

    return AwtPoint(int(round(x)), int(round(y)))


def _set_hand_points(hand, tip: Point, tail: Point) -> None:
    hand.PolyPolygon = ((_make_point(tip.x, tip.y), _make_point(tail.x, tail.y)),)


def _make_line_shape(document, slide, tip: Point, tail: Point, width_mm: float, color: int):
    hand = document.createInstance("com.sun.star.drawing.PolyLineShape")
    slide.add(hand)
    _set_hand_points(hand, tip, tail)
    hand.LineWidth = int(width_mm * MM_PER_100)
    hand.LineColor = color
    return hand


def create_analog_clock(
    document, slide, placeholder_shape, style: ClockStyle, now: datetime, slide_index: int
) -> AnalogClock:
    """Builds hour/minute/second hand shapes (plus a small pivot dot) sized
    and centered to match `placeholder_shape`'s bounding box, then hides the
    placeholder (which held the literal "#ANCLOCK" text). Hands are given
    their correct initial position for `now` immediately (via
    update_analog_clock) rather than starting at 12:00:00 - `now` is passed
    in explicitly (rather than calling datetime.now() here) so the caller's
    time source (real or mocked - see time_source.py) is respected.
    `slide_index` records which slide this clock belongs to (see
    AnalogClock.slide_index's docstring for why this matters). The second
    hand is omitted entirely if `style.show_second_hand` is False."""
    pos = placeholder_shape.Position
    size = placeholder_shape.Size
    rect = Rect(pos.X, pos.Y, size.Width, size.Height)
    center = rect.center
    radius = rect.radius

    hour_tip_length = radius * style.hour_length_ratio
    minute_tip_length = radius * style.minute_length_ratio
    second_tip_length = radius * style.second_length_ratio
    tail_length = radius * style.tail_length_ratio

    hour_tip, hour_tail = clock_geometry.hand_endpoints(center, 0.0, hour_tip_length, tail_length)
    minute_tip, minute_tail = clock_geometry.hand_endpoints(center, 0.0, minute_tip_length, tail_length)

    hour_hand = _make_line_shape(document, slide, hour_tip, hour_tail, style.hour_width_mm, style.hour_color)
    minute_hand = _make_line_shape(document, slide, minute_tip, minute_tail, style.minute_width_mm, style.minute_color)

    second_hand = None
    if style.show_second_hand:
        second_tip, second_tail = clock_geometry.hand_endpoints(center, 0.0, second_tip_length, tail_length)
        second_hand = _make_line_shape(document, slide, second_tip, second_tail, style.second_width_mm, style.second_color)

    pivot = document.createInstance("com.sun.star.drawing.EllipseShape")
    slide.add(pivot)
    pivot_r = style.pivot_radius_mm * MM_PER_100
    pivot.Position = _make_point(center.x - pivot_r, center.y - pivot_r)
    import uno
    from com.sun.star.awt import Size as AwtSize

    pivot.Size = AwtSize(int(pivot_r * 2), int(pivot_r * 2))
    pivot.FillColor = style.second_color
    # LineStyle is a UNO enum; `from com.sun.star.drawing import LineStyle`
    # is unreliable via pyuno's import hook (enums must be resolved via
    # uno.Enum, not the plain package-attribute import used for structs).
    pivot.LineStyle = uno.Enum("com.sun.star.drawing.LineStyle", "NONE")
    # Shapes are drawn in the order they're added to the slide (last added
    # = on top), and the pivot is added after the hands, so it naturally
    # covers their tail stubs - matching a real clock's look.

    placeholder_shape.setString("")
    placeholder_shape.Visible = False

    clock = AnalogClock(
        center=center,
        radius=radius,
        hour=hour_hand,
        minute=minute_hand,
        second=second_hand,
        pivot=pivot,
        placeholder=placeholder_shape,
        hour_tip_length=hour_tip_length,
        minute_tip_length=minute_tip_length,
        second_tip_length=second_tip_length,
        tail_length=tail_length,
        slide_index=slide_index,
    )
    update_analog_clock(clock, now)
    return clock


def update_analog_clock(clock: AnalogClock, now: datetime) -> None:
    """Recomputes each hand's tip/tail coordinates directly for the current
    time and re-sets the shape's polygon points (see clock_geometry module
    docstring for why this is done instead of setting RotateAngle). Each
    hand's raw computed angle is passed through
    clock_geometry.stabilize_angle against its own last-rendered angle
    first, so a small backward wall-clock time step (e.g. an NTP
    correction) never makes a hand visibly twitch backward. Skips the
    second hand entirely if the clock was built without one."""
    _refresh_hand(clock, clock.hour, clock_geometry.hour_angle_deg(now), clock.hour_tip_length, "last_hour_angle")
    _refresh_hand(clock, clock.minute, clock_geometry.minute_angle_deg(now), clock.minute_tip_length, "last_minute_angle")
    if clock.second is not None:
        _refresh_hand(clock, clock.second, clock_geometry.second_angle_deg(now), clock.second_tip_length, "last_second_angle")


def _refresh_hand(clock: AnalogClock, shape, raw_angle_deg: float, tip_length: float, last_angle_attr: str) -> None:
    """Stabilizes `raw_angle_deg` against the hand's own last-rendered
    angle (stored on `clock` under `last_angle_attr`), and only actually
    rewrites the shape's PolyPolygon if the stabilized angle differs from
    what was last rendered.

    The "only write if changed" part matters on top of stabilize_angle
    itself: on ticks where stabilize_angle decides to HOLD the hand at its
    previous position (suppressing a small backward wall-clock jitter -
    see that function's docstring), the computed tip/tail points would be
    IDENTICAL to what's already on screen. Re-setting a shape's
    PolyPolygon property to the exact same value it already had was
    observed/suspected to still cause LibreOffice's slideshow view to
    redraw/repaint that shape - which can itself look like a subtle
    flicker/twitch, precisely on the ticks meant to look perfectly static.
    Skipping the property write entirely when nothing actually needs to
    change avoids this residual artifact."""
    last_angle = getattr(clock, last_angle_attr)
    stabilized = clock_geometry.stabilize_angle(raw_angle_deg, last_angle)
    changed = last_angle is None or stabilized != last_angle
    setattr(clock, last_angle_attr, stabilized)
    if not changed:
        return
    tip, tail = clock_geometry.hand_endpoints(clock.center, stabilized, tip_length, clock.tail_length)
    _set_hand_points(shape, tip, tail)

