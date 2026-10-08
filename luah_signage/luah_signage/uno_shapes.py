"""
UNO-specific shape adapters: wraps raw UNO draw shapes so tagging.py's
generic (UNO-agnostic) logic can operate on them, plus the analog clock
shape builder/updater.

Most of this module is written carefully against the documented UNO Draw/
Impress API but is otherwise untestable in this dev environment (no real
signage deck/slideshow to drive). `UnoTextShape.set_text_range`'s specific
formatting-preservation behavior WAS empirically verified against a real
LibreOffice instance (headless, via its own bundled Python) - see that
method's docstring for the root cause it works around.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Iterator

from . import clock_geometry
from .clock_geometry import Point, Rect
from .config import ClockStyle

MM_PER_100 = 100  # UNO measurements are in 1/100 mm

# Character formatting properties captured/reapplied by
# UnoTextShape.set_text_range - covers the common run-level formatting a
# deck author might apply (bold/italic/size/color/underline/strikeout/
# font), including the separate Asian/Complex (CTL) property variants that
# actually govern Hebrew/RTL text rendering in LibreOffice (Western
# CharWeight/CharHeight alone are not always what's visually applied to
# Hebrew text - both sets are captured defensively).
_CHAR_FORMAT_PROPERTIES = (
    "CharWeight", "CharWeightAsian", "CharWeightComplex",
    "CharHeight", "CharHeightAsian", "CharHeightComplex",
    "CharPosture", "CharPostureAsian", "CharPostureComplex",
    "CharColor", "CharUnderline", "CharStrikeout",
    "CharFontName", "CharFontNameAsian", "CharFontNameComplex",
)


def _nudge_position(shape) -> None:
    """Moves ``shape`` right by 1 (1/100 mm) and immediately back to its
    original position - an experimental workaround for a LibreOffice
    slideshow rendering quirk where a shape's property write (a clock
    hand's PolyPolygon, or a tagged text shape's run content) updates the
    document model correctly but the live slideshow view doesn't always
    visibly repaint that shape on its own. Used by BOTH ``_refresh_hand``
    (clock hands - see LuahConfig.debug_repaint_nudge_shape) and
    ``UnoTextShape.nudge`` (tagged text shapes - see LuahConfig.
    debug_repaint_nudge_text_shapes).

    Why a Position nudge rather than toggling Visible (an earlier version
    of this function did exactly that for both use cases): empirically
    confirmed - live, on the real fullscreen slideshow, not just headless
    - that toggling a tagged TEXT shape's Visible property off then
    immediately back on does NOT reliably restore its visibility in the
    RUNNING slideshow view - the shape genuinely vanished after its first
    post-nudge content change and only reappeared once the slideshow
    navigated away from and back to that slide (i.e. a full slide
    re-entry, which independently resyncs every shape's visibility from
    the document model). This suggests Impress's live slideshow engine
    tracks a running show's per-shape "is this shown right now" state
    somewhat independently of the document model's own Visible property,
    and a rapid off/on toggle can race that internal state rather than
    reliably forcing an immediate repaint. A Position nudge sidesteps
    this risk entirely by never touching visibility-related state at all
    - and was adopted for clock hands too (even though no equivalent
    vanishing was ever reported for them) simply to keep both nudge code
    paths using the same, confirmed-safe mechanism rather than keeping
    two different ones around.

    The shape's Position UNO struct is a plain value type (not a live
    reference), so mutating a local copy and re-assigning it back to the
    ORIGINAL value afterward leaves the shape's final, settled position
    bit-for-bit identical to before this function ran - confirmed via a
    real-document round-trip test (capture Position before, nudge
    repeatedly, re-read Position after: byte-for-byte identical X/Y every
    time, no drift whatsoever). Swallows any exception (e.g. shape
    disposed mid-tick) rather than letting a cosmetic nudge crash an
    otherwise-successful refresh."""
    try:
        pos = shape.Position
        original_x = pos.X
        pos.X = original_x + 1
        shape.Position = pos
        pos.X = original_x
        shape.Position = pos
    except Exception:
        pass


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

    def nudge(self) -> None:
        """See module-level _nudge_position's docstring for why BOTH
        this (tagged text shapes) and clock hands use a position-based
        nudge rather than toggling Visible - called by tagging.
        TrackedShape.refresh after a tag's rendered value actually
        changes, ONLY when LuahConfig.debug_repaint_nudge_text_shapes is
        True."""
        _nudge_position(self._shape)

    def set_text_range(self, start: int, end: int, text: str) -> None:
        """Replaces characters [start, end) of the shape's CURRENT text
        with ``text``, preserving the character formatting (bold/italic/
        size/color/font/etc.) that range ALREADY had - unlike
        ``set_text``/``setString`` on the whole shape, which collapses
        every existing text run/portion into a single new run using one
        uniform formatting context (the ORIGINAL observed bug: a shape
        like "בבלי: #DAFYOMIBV" with "בבלי:" bold and the tag not bold
        would lose the bold formatting entirely once the tag was
        substituted via a whole-shape ``setString``).

        Formatting is explicitly CAPTURED from the target range and
        RE-APPLIED after the text is replaced, rather than relying on the
        replaced text to simply "inherit" the surrounding formatting
        implicitly. This was found to be necessary (not just a defensive
        extra step) via direct testing against a real LibreOffice
        instance: when a shape has two adjacent runs with DIFFERENT
        formatting (e.g. a bold run immediately followed by a non-bold
        run) and this range exactly starts at that run boundary, a plain
        ``cursor.setString(text)`` (and likewise
        ``xtext.insertString(cursor, text, True)`` with
        ``bAbsorb=True``) was observed to make the REPLACED text
        incorrectly inherit the PRECEDING run's formatting instead of its
        own - i.e. a second tag's rendered value picked up the FIRST
        tag's bold/size formatting purely because it happened to sit
        immediately after it in the same shape (reproduced with a real
        "#HEBDATE #DAYINFO:n" shape from this project's own deck,
        #HEBDATE bold/28 and #DAYINFO:n regular/24 - after substitution,
        BOTH became bold/28 without this fix). Capturing the OLD range's
        own properties BEFORE replacing it and explicitly re-applying them
        to the NEW range afterward sidesteps this cursor/run-boundary
        quirk entirely, since the new text's formatting no longer depends
        on any ambient/inherited cursor state at all.

        No-op (skips the capture/reapply step) when the OLD range is
        already empty (``start == end``) - there is no existing range to
        read formatting FROM in that case (e.g. a tag whose rendered
        value was previously "" and is becoming non-empty for the first
        time); the new text falls back to whatever formatting the
        (collapsed) insertion point otherwise provides, matching this
        method's pre-fix behavior for that specific edge case only."""
        xtext = self._shape.getText()
        cursor = xtext.createTextCursor()
        cursor.gotoStart(False)
        if start:
            cursor.goRight(start, False)
        if end > start:
            cursor.goRight(end - start, True)

        captured_format: dict = {}
        if end > start:
            for prop in _CHAR_FORMAT_PROPERTIES:
                try:
                    captured_format[prop] = cursor.getPropertyValue(prop)
                except Exception:
                    pass

        cursor.setString(text)

        if text and captured_format:
            format_cursor = xtext.createTextCursor()
            format_cursor.gotoStart(False)
            format_cursor.goRight(start, False)
            format_cursor.goRight(len(text), True)
            for prop, value in captured_format.items():
                try:
                    format_cursor.setPropertyValue(prop, value)
                except Exception:
                    pass



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

    def reset_stabilization(self) -> None:
        """Clears all 3 hands' last-rendered-angle state back to ``None``
        - see presentation.py's ``refresh_clock`` for why/when this is
        called (any tick where this clock is SKIPPED - e.g. its slide is
        currently hidden - rather than actually updated). ``None`` makes
        ``clock_geometry.stabilize_angle`` accept the next raw computed
        angle unconditionally (same as a hand's very first-ever render),
        bypassing its backward-jitter suppression for exactly one
        subsequent update."""
        self.last_hour_angle = None
        self.last_minute_angle = None
        self.last_second_angle = None


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


def update_analog_clock(
    clock: AnalogClock,
    now: datetime,
    *,
    skip_unchanged_writes: bool = True,
    nudge_shape: bool = False,
) -> None:
    """Recomputes each hand's tip/tail coordinates directly for the current
    time and re-sets the shape's polygon points (see clock_geometry module
    docstring for why this is done instead of setting RotateAngle). Each
    hand's raw computed angle is passed through
    clock_geometry.stabilize_angle against its own last-rendered angle
    first, so a small backward wall-clock time step (e.g. an NTP
    correction) never makes a hand visibly twitch backward. Skips the
    second hand entirely if the clock was built without one.

    `skip_unchanged_writes`/`nudge_shape` correspond directly to
    LuahConfig.debug_repaint_skip_unchanged_clock_writes/
    debug_repaint_nudge_shape - see that dataclass's docstrings for what
    each does and why they exist (diagnosing a machine-specific
    frozen-clock-hands rendering bug)."""
    _refresh_hand(
        clock, clock.hour, clock_geometry.hour_angle_deg(now), clock.hour_tip_length, "last_hour_angle",
        skip_unchanged_writes=skip_unchanged_writes, nudge_shape=nudge_shape,
    )
    _refresh_hand(
        clock, clock.minute, clock_geometry.minute_angle_deg(now), clock.minute_tip_length, "last_minute_angle",
        skip_unchanged_writes=skip_unchanged_writes, nudge_shape=nudge_shape,
    )
    if clock.second is not None:
        _refresh_hand(
            clock, clock.second, clock_geometry.second_angle_deg(now), clock.second_tip_length, "last_second_angle",
            skip_unchanged_writes=skip_unchanged_writes, nudge_shape=nudge_shape,
        )


def _refresh_hand(
    clock: AnalogClock,
    shape,
    raw_angle_deg: float,
    tip_length: float,
    last_angle_attr: str,
    *,
    skip_unchanged_writes: bool = True,
    nudge_shape: bool = False,
) -> None:
    """Stabilizes `raw_angle_deg` against the hand's own last-rendered
    angle (stored on `clock` under `last_angle_attr`), and rewrites the
    shape's PolyPolygon if the stabilized angle differs from what was
    last rendered (or unconditionally, if `skip_unchanged_writes=False` -
    see LuahConfig.debug_repaint_skip_unchanged_clock_writes).

    The "only write if changed" behavior matters on top of
    stabilize_angle itself: on ticks where stabilize_angle decides to
    HOLD the hand at its previous position (suppressing a small backward
    wall-clock jitter - see that function's docstring), the computed
    tip/tail points would be IDENTICAL to what's already on screen.
    Re-setting a shape's PolyPolygon property to the exact same value it
    already had was observed/suspected to still cause LibreOffice's
    slideshow view to redraw/repaint that shape - which can itself look
    like a subtle flicker/twitch, precisely on the ticks meant to look
    perfectly static. Skipping the property write entirely when nothing
    actually needs to change avoids this residual artifact.

    `nudge_shape=True` (see LuahConfig.debug_repaint_nudge_shape) nudges
    the shape's Position by 1/100 mm and immediately back immediately
    after writing its PolyPolygon, whenever a write actually happens - an
    experimental attempt to force a stubborn rendering backend to notice
    the change (see _nudge_position's docstring for why a Position nudge
    is used here rather than toggling Visible)."""
    last_angle = getattr(clock, last_angle_attr)
    stabilized = clock_geometry.stabilize_angle(raw_angle_deg, last_angle)
    changed = last_angle is None or stabilized != last_angle
    setattr(clock, last_angle_attr, stabilized)
    if not changed and skip_unchanged_writes:
        return
    tip, tail = clock_geometry.hand_endpoints(clock.center, stabilized, tip_length, clock.tail_length)
    _set_hand_points(shape, tip, tail)
    if nudge_shape:
        _nudge_position(shape)

