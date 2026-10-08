"""
Pure state machine for manually driving slide-to-slide advancement and
transition restoration - no UNO dependency, fully unit testable.

Why this exists: LibreOffice Impress's slideshow engine appears to tie a
slide's native "advance automatically after N seconds" behavior to its
transition effect internally - disabling the transition (done in
presentation.py to stop a flash/fade replaying on every content refresh)
also silently disables the native auto-advance timer. Rather than rely on
Impress's built-in scheduling at all, we take over slide advancement
entirely ourselves (closely mirroring the original VBA's myTimer/
advancecount mechanism in reference/LuahMain.bas, which manually called
View.Next based on each slide's own AdvanceTime):

For every slide, at prepare-time we capture its original auto-advance
duration and transition properties, then suppress both (transition -> none,
auto-advance -> manual/off). While a slide is being shown, we track how
long it's been current; once that reaches the slide's ORIGINAL duration, we
restore transition properties, then wait a short "settle" period
(transition_settle_seconds) before actually commanding the slideshow to
advance.

IMPORTANT caller contract - which slide's properties to touch: a slide's
transition effect is the animation played when ARRIVING at that slide, not
when leaving it. So on "prepare_advance", the caller must restore the
transition properties of the slide about to become current NEXT (not the
slide that is current right now) - restoring the current/outgoing slide's
own properties here would be a no-op for this advance (it would only
matter a full lap later, the next time that same slide is revisited,
manifesting as transitions appearing to only work "one loop late"). Then,
right after actually calling the advance (on "advance"), the caller should
immediately re-suppress the slide just arrived at, before any other content
refresh runs, so the transition doesn't get replayed by unrelated shape
updates on subsequent ticks.

Splitting restoration and advancing into two separate ticks (rather than
doing both in the same instant) is required because Impress's rendering
pipeline needs a brief real-time gap to pick up a just-changed transition
property before it's read at advance-time - restoring and advancing
immediately back-to-back was observed to silently skip the transition
effect entirely. A slide with an original duration of 0 (authored as
"advance on click only") is never auto-advanced, matching the original
VBA's own "AdvanceTime = 0 -> don't auto-advance" behavior.
"""
from __future__ import annotations

from dataclasses import dataclass

DEFAULT_TRANSITION_SETTLE_SECONDS = 1.0


@dataclass
class SlideTiming:
    duration_seconds: float
    transition_props: dict


class DynamicSlideAdvancer:
    """Tracks per-slide-index timing state and decides, on each tick,
    whether the current slide needs its suppression (re-)applied, needs its
    transition restored in preparation to advance, or needs to actually be
    advanced now. Slide identity is tracked by index (0-based, matching
    UNO's XSlideShowController.CurrentSlideIndex) rather than by object,
    since that's what the slideshow controller reports.
    """

    def __init__(self, transition_settle_seconds: float = DEFAULT_TRANSITION_SETTLE_SECONDS) -> None:
        self.transition_settle_seconds = transition_settle_seconds
        self._timings: dict[int, SlideTiming] = {}
        self._current_index: int | None = None
        self._entered_at: float | None = None
        # Set once we've restored the current slide's transition and are
        # just waiting out transition_settle_seconds before advancing.
        self._prepared_at: float | None = None

    def register(self, slide_index: int, duration_seconds: float, transition_props: dict) -> None:
        self._timings[slide_index] = SlideTiming(duration_seconds, dict(transition_props))

    def is_managed(self, slide_index: int) -> bool:
        return slide_index in self._timings

    def get_transition_props(self, slide_index: int) -> dict:
        return dict(self._timings[slide_index].transition_props)

    def reset(self) -> None:
        """Clears all tracked timing/registration state - call this after a
        document reload, since slide indices/shapes from the previous
        document are no longer valid."""
        self._timings.clear()
        self._current_index = None
        self._entered_at = None
        self._prepared_at = None

    def on_tick(self, current_slide_index: int, now: float, force_advance: bool = False) -> str:
        """Call once per tick with the slideshow's current slide index and
        the current monotonic time. Returns one of:
        - "none": nothing to do this tick.
        - "entered": `current_slide_index` just became current (either the
          show just started, or we/someone advanced to it); the caller
          should (re-)apply suppression to this slide (idempotent - safe to
          call even if already suppressed, e.g. if the show loops back to a
          previously-visited slide).
        - "prepare_advance": this slide's original duration has elapsed (OR
          `force_advance=True` was passed - see below); the caller should
          restore the NEXT slide's (the one about to become current, NOT
          this one - see module docstring) original transition properties
          now, but NOT advance yet - `on_tick` will return "advance" once
          `transition_settle_seconds` has passed, giving the transition
          property change time to take effect before the slideshow
          actually moves on.
        - "advance": the settle period has elapsed; the caller should now
          command the slideshow to advance to the next slide, immediately
          re-suppress that next slide's transition properties (they were
          only just restored, and are no longer needed live once the
          arrival transition has been triggered), then call
          `left_slide()`.
        Unmanaged slide indices (not registered via `register`) always
        return "none" - the caller should leave them to Impress's native
        behavior entirely (untouched, since we never modified their
        properties).

        `force_advance=True` (the caller should pass this when the
        CURRENTLY current slide has just become HIDDEN - e.g. a
        #CHOLONLY/#NONCHOLONLY day-mode flip caught it mid-display - and
        the show therefore needs to move off of it immediately,
        regardless of its own authored Duration/"advance on click only"
        setting) makes this slide immediately eligible to enter the
        "prepare_advance" -> (settle) -> "advance" sequence on THIS tick,
        bypassing the normal `timing.duration_seconds` elapsed-time check
        entirely (and even overriding a `duration_seconds <= 0` "manual
        advance only" slide, which would otherwise never trigger this
        sequence on its own). Still goes through the SAME settle-delay
        dance as a normal timed advance (not an instant cut) - reusing
        this one state machine for both cases is what lets a forced
        advance due to day-mode visibility changing still play its
        destination slide's authored transition effect correctly, rather
        than needing a separate, cruder "just cut to the next slide"
        code path with no transition at all. Has no effect once already
        mid-"prepare_advance"/settling (that sequence always runs to
        completion via the normal settle-based check below, whether it
        was entered via a forced or a natural duration-elapsed trigger)."""
        if current_slide_index != self._current_index:
            self._current_index = current_slide_index
            self._entered_at = now
            self._prepared_at = None
            return "entered" if self.is_managed(current_slide_index) else "none"

        if not self.is_managed(current_slide_index):
            return "none"

        if self._prepared_at is not None:
            if now - self._prepared_at >= self.transition_settle_seconds:
                return "advance"
            return "none"

        if force_advance:
            self._prepared_at = now
            return "prepare_advance"

        timing = self._timings[current_slide_index]
        if timing.duration_seconds <= 0:
            return "none"  # authored as "advance on click only" - never auto-advance
        if self._entered_at is not None and now - self._entered_at >= timing.duration_seconds:
            self._prepared_at = now
            return "prepare_advance"
        return "none"

    def left_slide(self) -> None:
        """Call right after successfully commanding the slideshow to
        advance away from the current slide, so the next `on_tick()` call
        (which will observe a new current index) re-initializes entry
        timing correctly rather than immediately re-triggering "advance"."""
        self._entered_at = None
        self._prepared_at = None

    def cancel_prepared_advance(self, now: float) -> None:
        """Call instead of actually advancing, when `on_tick` returned
        `"prepare_advance"` but the caller determines there is nowhere to
        actually advance TO (e.g. the only currently-visible slide in the
        deck is this one itself, so the computed "next visible slide"
        would just be this same slide looping back to itself - see
        presentation.py's `poll_slide_advance`, which is the only caller
        of this method).

        Resets `_prepared_at` back to `None` (so `on_tick` stops
        immediately returning `"advance"` on every subsequent tick - which
        it otherwise would, forever, since nothing ever clears a set
        `_prepared_at` except an actual advance or this method) AND resets
        `_entered_at` to `now` (restarting this slide's own
        duration-elapsed timer from scratch, exactly as if it had just
        been freshly (re-)entered) - this lets a LATER, genuine advance
        attempt be retried normally after another full `duration_seconds`
        elapses (e.g. once a day-mode flip elsewhere in the deck makes a
        real destination slide visible again), rather than either getting
        stuck retrying every single tick forever, or never being retried
        again at all."""
        self._prepared_at = None
        self._entered_at = now
