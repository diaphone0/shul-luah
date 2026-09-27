"""
Pure math for the analog clock hands - no UNO dependency, fully unit
testable. See uno_shapes.py for the UNO-specific shape creation that
consumes this.

Angles are measured clockwise from 12 o'clock (standard clock-face
convention), in degrees, range [0, 360).

Hands are redrawn from scratch every tick by recomputing their tip/tail
coordinates directly (see hand_endpoints) rather than by setting UNO's
RotateAngle property. RotateAngle always rotates a shape around the CENTER
OF ITS OWN (unrotated) BOUNDING BOX - for a straight-line shape, that center
is simply the midpoint between its two endpoints. This means a line can
only be rotated correctly "in place" around the clock's pivot if its tip and
tail are equidistant from the pivot (a symmetric hand) - which looks wrong
(each hand appears to span almost the full clock diameter, sticking out
equally on both sides of the pivot, rather than looking like a normal
asymmetric clock hand). Recomputing coordinates directly sidesteps this
limitation entirely and allows realistic hands (long tip, short/no tail).
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import datetime


def hour_angle_deg(dt: datetime) -> float:
    hour12 = dt.hour % 12
    return (hour12 + dt.minute / 60 + dt.second / 3600) * 30.0


def minute_angle_deg(dt: datetime) -> float:
    return (dt.minute + dt.second / 60) * 6.0


def second_angle_deg(dt: datetime) -> float:
    return (dt.second + dt.microsecond / 1_000_000) * 6.0


# Any backward step smaller in magnitude than this (degrees) is treated as
# clock-source jitter (see stabilize_angle) rather than a legitimate new-lap
# wraparound - a real wraparound's delta is close to -360 (e.g. 359.9 ->
# 0.1), which is unambiguously far below this threshold.
_WRAPAROUND_THRESHOLD_DEG = 180.0


def stabilize_angle(new_angle_deg: float, last_angle_deg: float | None) -> float:
    """Suppresses small backward regressions of a clock hand's angle,
    while still allowing the legitimate once-per-lap wraparound through.

    Why this exists: the wall-clock time source (datetime.now()) can very
    occasionally step slightly BACKWARD - e.g. from an OS/NTP time
    synchronization correction - even though it normally only moves
    forward. Since a hand's angle is recomputed directly from the current
    time every tick (see module docstring - no incremental
    rotation/RotateAngle is used), a backward time step translates
    directly into the hand's angle also stepping backward by a small
    amount for that one tick, which is visible as the hand seeming to
    "twitch backward slightly" before resuming its normal forward motion.
    Real analog clock hands never move backward, so this is purely a
    rendering artifact of directly-recomputed-every-tick angles reacting
    to a non-monotonic time source, not a meaningful state change.

    If `new_angle_deg` is behind `last_angle_deg` by less than
    `_WRAPAROUND_THRESHOLD_DEG`, the OLD angle is returned instead (the
    hand simply holds position rather than visibly regressing) - it will
    resume advancing normally once the wall clock catches back up past
    where it was already displayed. A backward "step" close to a full lap
    (e.g. last=359.9, new=0.1) is recognized as a genuine wraparound (not
    jitter) and passed through unchanged. `last_angle_deg=None` (first
    render) always accepts `new_angle_deg` as-is."""
    if last_angle_deg is None:
        return new_angle_deg % 360.0
    delta = new_angle_deg - last_angle_deg
    if -_WRAPAROUND_THRESHOLD_DEG < delta < 0.0:
        return last_angle_deg
    return new_angle_deg % 360.0


def to_uno_rotate_angle(clockwise_deg: float) -> int:
    """Convert a clockwise-from-12 angle (degrees) into a
    com.sun.star.drawing.XShape.RotateAngle value (1/100 degree,
    counter-clockwise from 3 o'clock/positive-X-axis, the UNO convention).

    Not used for the analog clock hands (see module docstring) - kept
    available for shapes that genuinely should rotate around their own
    bounding-box center (e.g. a decorative rotating dial ring)."""
    normalized = clockwise_deg % 360.0
    # UNO rotation is counter-clockwise; clock angles are clockwise, and
    # also offset by 90 degrees (12-o'clock vs 3-o'clock reference).
    uno_deg = (90.0 - normalized) % 360.0
    return int(round(uno_deg * 100.0)) % 36000


@dataclass(frozen=True)
class Point:
    x: float
    y: float


def hand_endpoints(
    center: Point, angle_deg: float, tip_length: float, tail_length: float = 0.0
) -> tuple[Point, Point]:
    """Returns (tip, tail) endpoints of a hand pivoting at `center`, at the
    given clockwise-from-12 angle: the tip is `tip_length` away from center
    in the direction of the angle, and the tail is `tail_length` away in the
    opposite direction (a short stub, `tail_length=0` for no tail - the hand
    starts exactly at the pivot, the normal look for a clock hand)."""
    rad = math.radians(angle_deg)
    sin_a, cos_a = math.sin(rad), -math.cos(rad)
    tip = Point(center.x + sin_a * tip_length, center.y + cos_a * tip_length)
    tail = Point(center.x - sin_a * tail_length, center.y - cos_a * tail_length)
    return tip, tail


@dataclass(frozen=True)
class Rect:
    x: float
    y: float
    width: float
    height: float

    @property
    def center(self) -> Point:
        return Point(self.x + self.width / 2, self.y + self.height / 2)

    @property
    def radius(self) -> float:
        return min(self.width, self.height) / 2
