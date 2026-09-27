"""
Sanity tests for luah_signage.clock_geometry (pure math, no UNO dependency).
Run with: python tests/test_clock_geometry.py (also pytest-compatible)
"""
from __future__ import annotations

from datetime import datetime

from luah_signage.clock_geometry import (
    Point,
    Rect,
    hand_endpoints,
    hour_angle_deg,
    minute_angle_deg,
    second_angle_deg,
    stabilize_angle,
    to_uno_rotate_angle,
)


def test_hour_angle_noon():
    assert hour_angle_deg(datetime(2024, 1, 1, 0, 0, 0)) == 0.0
    assert hour_angle_deg(datetime(2024, 1, 1, 3, 0, 0)) == 90.0
    assert hour_angle_deg(datetime(2024, 1, 1, 6, 0, 0)) == 180.0
    assert hour_angle_deg(datetime(2024, 1, 1, 9, 0, 0)) == 270.0


def test_hour_angle_wraps_pm():
    # 15:00 -> hour12 = 3 -> same angle as 3:00
    assert hour_angle_deg(datetime(2024, 1, 1, 15, 0, 0)) == 90.0


def test_minute_angle():
    assert minute_angle_deg(datetime(2024, 1, 1, 0, 0, 0)) == 0.0
    assert minute_angle_deg(datetime(2024, 1, 1, 0, 15, 0)) == 90.0
    assert minute_angle_deg(datetime(2024, 1, 1, 0, 30, 0)) == 180.0
    assert minute_angle_deg(datetime(2024, 1, 1, 0, 45, 0)) == 270.0


def test_second_angle():
    assert second_angle_deg(datetime(2024, 1, 1, 0, 0, 0)) == 0.0
    assert second_angle_deg(datetime(2024, 1, 1, 0, 0, 30)) == 180.0


def test_to_uno_rotate_angle_reference_points():
    # clock angle 0 (12 o'clock, pointing up) -> UNO 90deg (9000 in 1/100deg)
    assert to_uno_rotate_angle(0.0) == 9000
    # clock angle 90 (3 o'clock, pointing right) -> UNO 0deg
    assert to_uno_rotate_angle(90.0) == 0
    # clock angle 180 (6 o'clock, pointing down) -> UNO 270deg
    assert to_uno_rotate_angle(180.0) == 27000
    # clock angle 270 (9 o'clock, pointing left) -> UNO 180deg
    assert to_uno_rotate_angle(270.0) == 18000


def test_hand_endpoints_no_tail_starts_at_pivot():
    center = Point(100.0, 100.0)
    tip, tail = hand_endpoints(center, angle_deg=0.0, tip_length=50.0, tail_length=0.0)
    # tail defaults to sitting exactly at the pivot when tail_length=0
    assert tail == center
    # pointing straight up: tip has smaller y (screen coords, y grows down)
    assert tip.y < center.y
    assert tip.x == center.x


def test_hand_endpoints_tip_farther_than_tail():
    center = Point(50.0, 20.0)
    tip, tail = hand_endpoints(center, angle_deg=37.0, tip_length=40.0, tail_length=8.0)

    def dist(p):
        return ((p.x - center.x) ** 2 + (p.y - center.y) ** 2) ** 0.5

    assert abs(dist(tip) - 40.0) < 1e-9
    assert abs(dist(tail) - 8.0) < 1e-9
    # tail is in the opposite direction from tip (not symmetric-length, but
    # still collinear through the pivot on the opposite side)
    assert dist(tip) > dist(tail)


def test_hand_endpoints_equal_lengths_are_symmetric_about_center():
    # Sanity check: passing equal tip/tail lengths recovers symmetry (the
    # old default behavior), useful e.g. for shapes that DO want to rotate
    # via RotateAngle around their own bbox center.
    center = Point(50.0, 20.0)
    for angle in (0, 37, 90, 123.4, 270, 359):
        tip, tail = hand_endpoints(center, angle_deg=angle, tip_length=12.0, tail_length=12.0)
        assert abs((tip.x + tail.x) / 2 - center.x) < 1e-9
        assert abs((tip.y + tail.y) / 2 - center.y) < 1e-9


def test_rect_center_and_radius():
    rect = Rect(x=10, y=20, width=100, height=60)
    assert rect.center == Point(60, 50)
    assert rect.radius == 30  # min(width, height) / 2


def test_stabilize_angle_first_call_accepts_new_angle():
    assert stabilize_angle(123.45, None) == 123.45


def test_stabilize_angle_forward_movement_passes_through():
    assert stabilize_angle(125.3, 125.0) == 125.3
    assert stabilize_angle(10.0, 5.0) == 10.0


def test_stabilize_angle_small_backward_jitter_is_suppressed():
    # e.g. an NTP time-sync correction stepping the wall clock back a
    # fraction of a second - must NOT visibly regress the hand.
    assert stabilize_angle(125.25, 125.30) == 125.30
    assert stabilize_angle(10.0, 15.0) == 15.0


def test_stabilize_angle_wraparound_passes_through():
    # A genuine new-lap wraparound (e.g. minute hand going from just under
    # 360 back to just over 0) must NOT be mistaken for backward jitter.
    assert stabilize_angle(0.5, 359.8) == 0.5
    assert stabilize_angle(0.0, 359.999) == 0.0


def test_stabilize_angle_exact_same_angle_is_a_noop():
    assert stabilize_angle(200.0, 200.0) == 200.0


if __name__ == "__main__":
    import sys
    import traceback

    tests = [obj for name, obj in list(globals().items()) if name.startswith("test_")]
    failures = 0
    for t in tests:
        try:
            t()
            print(f"PASS: {t.__name__}")
        except Exception:
            failures += 1
            print(f"FAIL: {t.__name__}")
            traceback.print_exc()
    print(f"\n{len(tests) - failures}/{len(tests)} passed")
    sys.exit(1 if failures else 0)
