"""
Sanity tests for luah_signage.slide_advance (pure state machine, no UNO
dependency). Run with: python tests/test_slide_advance.py (also
pytest-compatible)
"""
from __future__ import annotations

from luah_signage.slide_advance import DynamicSlideAdvancer


def test_entering_managed_slide_returns_entered():
    adv = DynamicSlideAdvancer()
    adv.register(0, duration_seconds=10.0, transition_props={"TransitionType": 5})
    assert adv.on_tick(0, now=100.0) == "entered"


def test_entering_unmanaged_slide_returns_none():
    adv = DynamicSlideAdvancer()
    adv.register(0, duration_seconds=10.0, transition_props={})
    assert adv.on_tick(1, now=100.0) == "none"


def test_stays_none_before_duration_elapses():
    adv = DynamicSlideAdvancer()
    adv.register(0, duration_seconds=10.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=105.0) == "none"
    assert adv.on_tick(0, now=109.9) == "none"


def test_advance_triggers_once_duration_elapses():
    adv = DynamicSlideAdvancer(transition_settle_seconds=1.0)
    adv.register(0, duration_seconds=10.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=109.0) == "none"
    assert adv.on_tick(0, now=110.0) == "prepare_advance"
    # Still settling - not yet time to actually advance.
    assert adv.on_tick(0, now=110.5) == "none"
    assert adv.on_tick(0, now=111.0) == "advance"


def test_prepare_advance_only_returned_once_while_settling():
    # Once "prepare_advance" fires, subsequent ticks before the settle
    # period elapses must return "none" (not re-trigger "prepare_advance"
    # repeatedly, which would re-apply the transition restore pointlessly
    # every tick).
    adv = DynamicSlideAdvancer(transition_settle_seconds=2.0)
    adv.register(0, duration_seconds=5.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=105.0) == "prepare_advance"
    assert adv.on_tick(0, now=105.5) == "none"
    assert adv.on_tick(0, now=106.5) == "none"
    assert adv.on_tick(0, now=107.0) == "advance"


def test_zero_duration_never_advances():
    # Authored as "advance on click only" - must never be auto-advanced.
    adv = DynamicSlideAdvancer()
    adv.register(0, duration_seconds=0.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    for t in (101.0, 200.0, 100000.0):
        assert adv.on_tick(0, now=t) == "none"


def test_left_slide_resets_entry_timing_for_next_visit():
    adv = DynamicSlideAdvancer(transition_settle_seconds=1.0)
    adv.register(0, duration_seconds=5.0, transition_props={})
    adv.register(1, duration_seconds=5.0, transition_props={})

    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=106.0) == "prepare_advance"
    assert adv.on_tick(0, now=107.0) == "advance"
    adv.left_slide()

    # Now the caller commands the slideshow to slide index 1.
    assert adv.on_tick(1, now=107.5) == "entered"
    assert adv.on_tick(1, now=108.0) == "none"
    assert adv.on_tick(1, now=112.5) == "prepare_advance"
    assert adv.on_tick(1, now=113.5) == "advance"


def test_looping_back_to_a_previously_visited_slide_reenters():
    adv = DynamicSlideAdvancer(transition_settle_seconds=1.0)
    adv.register(0, duration_seconds=5.0, transition_props={"TransitionType": 3})

    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=106.0) == "prepare_advance"
    assert adv.on_tick(0, now=107.0) == "advance"
    adv.left_slide()
    assert adv.on_tick(1, now=107.5) == "none"  # slide 1 unmanaged
    # Show loops back around to slide 0 later.
    assert adv.on_tick(0, now=200.0) == "entered"


def test_get_transition_props_returns_a_copy():
    adv = DynamicSlideAdvancer()
    original = {"TransitionType": 7, "TransitionSubtype": 2}
    adv.register(0, duration_seconds=5.0, transition_props=original)
    props = adv.get_transition_props(0)
    props["TransitionType"] = 999
    assert adv.get_transition_props(0)["TransitionType"] == 7


def test_reset_clears_all_state():
    adv = DynamicSlideAdvancer()
    adv.register(0, duration_seconds=5.0, transition_props={})
    adv.on_tick(0, now=100.0)
    adv.reset()
    assert not adv.is_managed(0)
    # Re-entering slide 0 after reset (unregistered) is "none" until
    # re-registered.
    assert adv.on_tick(0, now=200.0) == "none"


def test_force_advance_triggers_prepare_advance_immediately():
    # Even though duration_seconds=100 (far from elapsed), force_advance=True
    # must still trigger "prepare_advance" right away - mirrors a
    # #CHOLONLY/#NONCHOLONLY day-mode flip hiding the currently-shown slide.
    adv = DynamicSlideAdvancer(transition_settle_seconds=1.0)
    adv.register(0, duration_seconds=100.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=100.5, force_advance=True) == "prepare_advance"
    assert adv.on_tick(0, now=100.9, force_advance=True) == "none"
    assert adv.on_tick(0, now=101.5, force_advance=True) == "advance"


def test_force_advance_overrides_manual_advance_only_slide():
    # A slide authored as "advance on click only" (duration_seconds=0)
    # would normally NEVER auto-advance - force_advance must override this
    # too, since a hidden slide must not get stuck forever.
    adv = DynamicSlideAdvancer(transition_settle_seconds=1.0)
    adv.register(0, duration_seconds=0.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=200.0) == "none"  # confirm it would NOT advance on its own
    assert adv.on_tick(0, now=300.0, force_advance=True) == "prepare_advance"
    assert adv.on_tick(0, now=301.0, force_advance=True) == "advance"


def test_force_advance_does_not_restart_an_already_settling_sequence():
    # Once "prepare_advance" has fired (whether triggered naturally or by
    # force_advance), subsequent force_advance=True ticks while still
    # settling must not re-trigger "prepare_advance" again.
    adv = DynamicSlideAdvancer(transition_settle_seconds=2.0)
    adv.register(0, duration_seconds=100.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=101.0, force_advance=True) == "prepare_advance"
    assert adv.on_tick(0, now=101.5, force_advance=True) == "none"
    assert adv.on_tick(0, now=102.5, force_advance=True) == "none"
    assert adv.on_tick(0, now=103.0, force_advance=True) == "advance"


def test_force_advance_false_behaves_exactly_like_default():
    adv = DynamicSlideAdvancer()
    adv.register(0, duration_seconds=10.0, transition_props={})
    assert adv.on_tick(0, now=100.0, force_advance=False) == "entered"
    assert adv.on_tick(0, now=105.0, force_advance=False) == "none"


def test_cancel_prepared_advance_stops_advance_from_firing_again():
    # Mirrors presentation.py's handling when there's nowhere to actually
    # advance to (e.g. the current slide is the ONLY visible slide in the
    # deck) - after cancel_prepared_advance, on_tick must NOT immediately
    # return "advance" on the next tick (it would, forever, if the caller
    # had done nothing).
    adv = DynamicSlideAdvancer(transition_settle_seconds=1.0)
    adv.register(0, duration_seconds=10.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=110.0) == "prepare_advance"
    adv.cancel_prepared_advance(now=110.0)
    # Must NOT fire "advance" even after the settle period would have
    # elapsed - the attempt was cancelled. (The duration timer DOES
    # restart from the cancellation point - see the next test - so this
    # only checks shortly after cancelling, before a full new
    # duration_seconds has elapsed again.)
    assert adv.on_tick(0, now=111.5) == "none"
    assert adv.on_tick(0, now=115.0) == "none"


def test_cancel_prepared_advance_restarts_the_duration_timer():
    # After cancelling, the slide's own duration-elapsed timer must
    # restart from the cancellation point - so a genuine future advance
    # attempt (e.g. once a day-mode flip makes a real destination slide
    # visible) is retried after another FULL duration_seconds, not
    # instantly and not never.
    adv = DynamicSlideAdvancer(transition_settle_seconds=1.0)
    adv.register(0, duration_seconds=10.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    assert adv.on_tick(0, now=110.0) == "prepare_advance"
    adv.cancel_prepared_advance(now=110.0)
    assert adv.on_tick(0, now=115.0) == "none"  # only 5s since cancellation - not yet
    assert adv.on_tick(0, now=119.9) == "none"
    assert adv.on_tick(0, now=120.0) == "prepare_advance"  # a full 10s after cancellation


def test_cancel_prepared_advance_is_a_noop_if_called_without_a_pending_prepare():
    adv = DynamicSlideAdvancer(transition_settle_seconds=1.0)
    adv.register(0, duration_seconds=10.0, transition_props={})
    assert adv.on_tick(0, now=100.0) == "entered"
    adv.cancel_prepared_advance(now=105.0)
    # Timer restarts from the cancel call regardless - same as above, just
    # confirming this doesn't raise/misbehave when called speculatively
    # with no actual pending "prepare_advance" in progress.
    assert adv.on_tick(0, now=114.9) == "none"
    assert adv.on_tick(0, now=115.0) == "prepare_advance"


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
