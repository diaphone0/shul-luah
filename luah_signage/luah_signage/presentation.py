"""
Owns the loaded Impress document: initial load, reload-on-change, tag/clock
shape scanning, and per-tick refresh. This replaces
init_shapes/reset_shapes/update_shapes/luah/reload_ssw from LuahMain.bas
with a much simpler design, since we never write computed values back into
the source .pptx (see module docstring in tagging.py) - there is no need for
a reset_shapes-style "undo" pass, nor a save+relaunch dance.

File-locking / cloud-sync safety: the source .pptx is NEVER opened directly
by LibreOffice. Instead, on every load/reload we copy it to a throwaway temp
file (via shutil.copy2, which opens the source only briefly with a normal
shared read handle and closes it immediately) and open THAT copy in
LibreOffice instead. This guarantees the file your cloud-sync client
manages is never held open by LibreOffice for the whole display duration -
mirroring the original VBA's own workaround (it also copied luah.pptx to a
temp file before opening it, for exactly this reason - see reload_ssw in
reference/LuahMain.bas).

Slide advancement is entirely manual (see slide_advance.py): Impress's own
"advance automatically after N seconds" + transition-effect scheduling is
suppressed on every slide, and a DynamicSlideAdvancer decides, once per
tick, when to restore a slide's original transition and command the
slideshow to move on - mirroring the original VBA's myTimer/advancecount
mechanism (reference/LuahMain.bas), which also manually drove slide
advancement rather than relying on PowerPoint's native scheduling. This is
necessary because Impress's slideshow renderer appears to replay a slide's
transition effect not just when advancing, but also whenever it detects the
displayed slide's content changed underneath it (e.g. our per-second
hashtag/clock refreshes) - producing a visible flash on every refresh. And
disabling the transition effect alone (without also taking over advancing)
was observed to silently disable the native auto-advance timer too.

Not testable in this dev environment (no LibreOffice installed here) -
written carefully against the documented UNO API but should be
smoke-tested on the actual signage machine.
"""
from __future__ import annotations

import math
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path

from . import tagging
from .config import LuahConfig
from .slide_advance import DynamicSlideAdvancer
from .uno_shapes import (
    AnalogClock,
    UnoTextShape,
    create_analog_clock,
    is_text_shape,
    iter_all_shapes,
    update_analog_clock,
)
from .time_source import TimeSource
from .zman_context import build_zman_context

_TRANSITION_PROP_NAMES = ("TransitionType", "TransitionSubtype", "TransitionDuration")


def _path_to_url(path: Path) -> str:
    import uno

    return uno.systemPathToFileUrl(str(Path(path).resolve()))


def _make_props(**kwargs):
    from com.sun.star.beans import PropertyValue

    props = []
    for name, value in kwargs.items():
        p = PropertyValue()
        p.Name = name
        p.Value = value
        props.append(p)
    return tuple(props)


def _move_window_offscreen(document) -> None:
    """Moves the underlying Impress document's window far off-screen and
    shrinks it to a minimal size, so nothing of it is visible to the user,
    while leaving it a genuine, valid window (unlike loading the document
    with Hidden=True, which prevents a real window/frame from being
    created at all and was found to break XSlideShowController.start() -
    see load()'s comment). The fullscreen slideshow window, once started,
    is a separate top-level window and is unaffected by this."""
    try:
        window = document.CurrentController.Frame.ContainerWindow
        # com.sun.star.awt.PosSize.POSSIZE = X | Y | WIDTH | HEIGHT (1|2|4|8)
        # combined - moves AND resizes in one call. Plain int constant (not
        # a UNO struct or enum), safe to hardcode.
        window.setPosSize(-32000, -32000, 1, 1, 15)
    except Exception:
        pass


def _restore_window_normal_size(document) -> None:
    """Counterpart to _move_window_offscreen, used when
    config.hide_editor_window is False (debugging). A document loaded via
    loadComponentFromURL (rather than opened by a user click) is NOT given
    a normal-looking window by LibreOffice on its own - it defaults to a
    tiny window (observed effectively at/near 0x0 or a few pixels), which
    looks identical to "hidden" even though hide_editor_window=False only
    means "don't explicitly shrink/move it off-screen" - the window still
    needs to be EXPLICITLY given a reasonable on-screen size/position to
    actually be visible/usable for debugging. Sets a generous default size
    positioned near the screen's top-left and brings it to the front."""
    try:
        window = document.CurrentController.Frame.ContainerWindow
        window.setPosSize(50, 50, 1200, 800, 15)
        window.setVisible(True)
        window.toFront()
    except Exception:
        pass


def _capture_original_timing(slide) -> dict:
    """Reads `slide`'s original Duration (auto-advance-after-N-seconds) and
    transition properties WITHOUT mutating anything. Used at prepare-time
    to register every slide's original values with a DynamicSlideAdvancer,
    while leaving every slide's actual properties completely untouched
    until the moment it's actually about to be shown (see
    `_suppress_slide_timing`, called lazily instead) - see that function's
    docstring for why suppressing every slide's timing eagerly, right at
    document-load time, was found to be harmful."""
    original: dict = {}
    try:
        original["Duration"] = slide.Duration
    except Exception:
        original["Duration"] = 0
    for name in _TRANSITION_PROP_NAMES:
        try:
            original[name] = getattr(slide, name)
        except Exception:
            original[name] = 0
    return original


def _suppress_slide_timing(slide) -> None:
    """Suppresses `slide`'s native auto-advance (Change -> manual/0) and
    transition (-> none) so Impress's own scheduling never fires for it -
    we take over advancing it entirely ourselves (see module docstring and
    slide_advance.py).

    IMPORTANT: call this LAZILY - only right when `slide` is actually about
    to become the current slide (i.e. from poll_slide_advance's "entered"/
    "advance" handling) - NEVER preemptively for every slide right after
    document load. An earlier version of this code called the
    read-and-mutate combo (formerly `_capture_and_suppress_slide_timing`)
    on every slide immediately in `_prepare_document`, well before most
    slides were ever shown - this was observed to cause a shape with an
    authored entrance/build animation on it to render invisible the FIRST
    time its slide was shown (but fine on every subsequent visit),
    presumably because Impress's slideshow engine pre-computes/caches
    per-slide animation timelines using each slide's Duration/Transition
    properties at some point during slideshow startup, and zeroing those
    out for every slide ahead of time corrupted that computation for
    slides not yet visited. Suppressing only the slide that's actually
    about to be shown, at the last possible moment, avoids disturbing
    slides that haven't been prepared/visited yet.

    This only touches the disposable in-memory temp-copy of the
    presentation (see module docstring) - never the original .pptx on
    disk."""
    try:
        slide.Change = 0  # manual advance only - we drive it ourselves
    except Exception:
        pass
    for name in _TRANSITION_PROP_NAMES:
        try:
            setattr(slide, name, 0)
        except Exception:
            pass


def _restore_slide_transition(slide, original: dict) -> None:
    """Restores `slide`'s original transition properties (captured by
    `_capture_original_timing`) so the visual transition plays as the
    author configured when we command the slideshow to advance away from
    this slide. `Change`/`Duration` are deliberately NOT restored - we
    permanently own advancing this slide via DynamicSlideAdvancer, so
    Impress's own auto-advance timer must stay off to avoid a double
    advance."""
    for name in _TRANSITION_PROP_NAMES:
        try:
            setattr(slide, name, original.get(name, 0))
        except Exception:
            pass


# Values used in PresentationController._slide_day_mode.
DAY_MODE_CHOL_ONLY = "CHOL_ONLY"
DAY_MODE_NONCHOL_ONLY = "NONCHOL_ONLY"


def _reconcile_day_modes(
    cholonly_indices: set, noncholonly_indices: set, slide_count: int
) -> dict:
    """Pure (no UNO dependency) reconciliation of which slides are marked
    #CHOLONLY / #NONCHOLONLY (as collected during shape scanning) into a
    final per-slide-index day-mode dict.

    Per the design: a slide tagged #CHOLONLY (and NOT also #NONCHOLONLY)
    shows only on chol (regular weekday) days; a slide tagged #NONCHOLONLY
    (and NOT also #CHOLONLY) shows only on non-chol (Shabbos/Yom Tov -
    assur-be-melachah) days; a slide with NEITHER tag, or with BOTH tags
    (an authoring edge case, e.g. two separate marker shapes on the same
    slide), shows every day - such slides are simply left out of the
    returned dict entirely (the caller should treat "not present in this
    dict" as "always visible, never touched")."""
    result: dict = {}
    for index in range(slide_count):
        is_cholonly = index in cholonly_indices
        is_noncholonly = index in noncholonly_indices
        if is_cholonly and not is_noncholonly:
            result[index] = DAY_MODE_CHOL_ONLY
        elif is_noncholonly and not is_cholonly:
            result[index] = DAY_MODE_NONCHOL_ONLY
        # else (neither, or both): always visible - not added to the dict.
    return result


def _next_visible_slide_index(current_index: int, visible_flags: list) -> tuple:
    """Pure (no UNO dependency) helper: given the current slide index and
    a list of booleans (`visible_flags[i]` = whether slide i is currently
    visible - whether hidden via #CHOLONLY/#NONCHOLONLY, the deprecated
    #MULTIMODE, or simply manually hidden by the deck author via normal
    "Hide Slide" authoring), returns `(next_index, wrapped)`:
    - `next_index`: the index of the next VISIBLE slide the live
      slideshow will actually navigate to when advancing forward from
      `current_index` - skipping over any hidden slides in between,
      exactly matching what Impress's own `gotoNextSlide()` does. `None`
      if NO slide in the deck is currently visible at all (a degenerate
      all-hidden deck).
    - `wrapped`: `True` if reaching `next_index` required wrapping around
      past the end of the deck back toward slide 0 (i.e. the caller
      should use `gotoFirstSlide()` instead of `gotoNextSlide()` to
      actually get there, matching Impress's own end-of-show wraparound
      behavior).

    This is essential because naive "current_index + 1, wrapping at
    slide_count" arithmetic assumes every slide is always visible - once
    ANY slide in the deck is hidden, that assumption breaks: Impress's
    live `gotoNextSlide()`/`gotoFirstSlide()` calls silently SKIP hidden
    slides, but naive arithmetic does not - causing transition properties
    to be restored/suppressed on the WRONG (hidden, never-actually-shown)
    slide instead of the slide the show will really land on next. This
    was found to corrupt transition-on-arrival behavior in exactly this
    way (observed symptom: transitions/flicker behaving correctly only on
    some laps, or not at all, once a hidden slide exists anywhere in the
    deck) - see repo memory for the full user-reported bug report and
    diagnosis that led to this fix."""
    count = len(visible_flags)
    if count == 0:
        return None, False
    for offset in range(1, count + 1):
        raw = current_index + offset
        idx = raw % count
        if visible_flags[idx]:
            return idx, raw >= count
    return None, False


class PresentationController:
    def __init__(self, config: LuahConfig, uno_ctx, time_source: TimeSource | None = None):
        self.config = config
        self.uno_ctx = uno_ctx
        # Defaults to a real (datetime.now()-based) TimeSource if not
        # given - app.py always passes one explicitly (built from
        # config.mock_start_datetime via time_source.create_time_source),
        # but this default keeps direct construction (e.g. in a REPL or a
        # future test) convenient. See time_source.py.
        self.time_source = time_source or TimeSource()
        self.desktop = uno_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", uno_ctx
        )
        self.document = None
        self.tracked_shapes: list[tagging.TrackedShape] = []
        self.clocks: list[AnalogClock] = []
        # Shapes tagged #MULTIMODE - deprecated/inert (superseded by the
        # #CHOLONLY/#NONCHOLONLY per-slide mechanism below), still just
        # hidden for backward compatibility with decks authored before
        # that mechanism existed.
        self.multimode_shapes: list = []
        # slide index -> DAY_MODE_CHOL_ONLY/DAY_MODE_NONCHOL_ONLY, for
        # slides tagged with a #CHOLONLY or #NONCHOLONLY marker shape (see
        # _reconcile_day_modes). Slides not present in this dict are
        # always shown, untouched.
        self._slide_day_mode: dict = {}
        self._tmp_path: Path | None = None
        self._slides_by_index: list = []
        self._advancer = DynamicSlideAdvancer(
            transition_settle_seconds=config.transition_settle_seconds
        )
        # (slide_index, monotonic-timestamp) up to which refresh_clock
        # should skip updating THAT SPECIFIC slide's clock hands - set
        # right after advancing to a new slide, for the duration of that
        # slide's OWN transition effect (ceiled to whole seconds), so its
        # clock-hand update never lands mid-transition-animation and
        # causes a visible flicker/flash. Scoped to a single slide index
        # (not a global pause) so OTHER visible slides' clocks keep
        # ticking normally in the background - see refresh_clock and
        # poll_slide_advance's "advance" branch.
        self._clock_suppressed: tuple | None = None
        # The slide index (if any) whose transition properties are
        # currently "live"/restored ahead of an upcoming advance (set
        # during poll_slide_advance's "prepare_advance" phase, cleared
        # once that advance actually happens or a new slide is entered) -
        # refresh_clock must NOT update this slide's clock while its
        # pending transition is live, even though it's still a visible,
        # not-yet-current slide (see round 2's bug A in repo memory for
        # why this specific case must remain excluded, even now that
        # other not-currently-shown slides ARE updated again).
        self._pending_transition_slide_index: int | None = None
        # The current slide index, as read ONCE per tick by
        # poll_slide_advance (always called before refresh_clock in
        # app.py's tick, by design) and then reused by refresh_clock -
        # rather than each independently querying
        # controller.CurrentSlideIndex via its own separate UNO call.
        # Querying it twice per tick from two different call sites was
        # found to risk transient inconsistency between the two reads
        # (e.g. one observing the index from just before an advance, the
        # other just after) - which, given refresh_clock filters which
        # clock(s) to update by comparing against this value, could
        # permanently freeze a slide's clock if that slide's index never
        # happened to match whatever inconsistent value refresh_clock
        # observed. Sharing a single per-tick read eliminates the
        # possibility of any such mismatch entirely.
        self._last_known_slide_index: int | None = None
        # See LuahConfig.debug_slide_advance's docstring.
        self._debug = config.debug_slide_advance
        self._debug_last_clock_log_index: object = "unset"
        # #CMD tag support (see tagging.py's module docstring for the
        # full cache-based design): maps a #CMD tag's normalized args-key
        # string -> that command's last-resolved stdout. Populated/
        # refreshed by refresh_cmd_tags (its own, infrequent tick - see
        # app.py), NEVER by refresh_content/render_template directly.
        # Deliberately NOT reset on reload (see _close_document) so a
        # previously-resolved value stays visible across a reload instead
        # of the tag going blank again while waiting for the next
        # (possibly many-minutes-away) cmd-refresh tick.
        self._cmd_cache: dict[str, str] = {}
        # Distinct #CMD args-keys found in the CURRENTLY loaded deck's
        # tracked shapes (recomputed fresh on every _prepare_document) -
        # this is what refresh_cmd_tags actually iterates over.
        self._cmd_arg_keys: set[str] = set()

    # -- loading ------------------------------------------------------
    def load(self) -> None:
        self._close_document()

        tmp_path = Path(tempfile.gettempdir()) / "luah_signage_live.pptx"
        shutil.copy2(self.config.pptx_path, tmp_path)
        self._tmp_path = tmp_path

        url = _path_to_url(tmp_path)
        # Hidden=True was tried here (to suppress the Impress editing
        # window) but had to be reverted - it appears to prevent Impress
        # from creating a real window/frame for the document at all, and
        # XSlideShowController.start() needs a genuine window to attach
        # its rendering context to; with Hidden=True, start_slideshow()
        # crashed with DisposedException: Binary URP bridge disposed
        # during call. Do not re-add Hidden=True.
        # ReadOnly=True: belt-and-suspenders - even though this is a
        # disposable temp copy (not the real synced file), we still never
        # want LibreOffice prompting about save formats or writing back.
        props = _make_props(Hidden=False, ReadOnly=True)
        self.document = self.desktop.loadComponentFromURL(url, "_blank", 0, props)
        # Instead of hiding the document outright, move its (real, valid)
        # window far off-screen and shrink it to a minimal size - this
        # keeps a genuine window for the slideshow to use as its parent,
        # while nothing of the editing window itself is visible to the
        # user. See _move_window_offscreen for details. Debug off-switch:
        # config.hide_editor_window=False leaves it in its normal on-screen
        # position/size, useful while debugging.
        if self.config.hide_editor_window:
            _move_window_offscreen(self.document)
        else:
            # LibreOffice gives API-opened documents a tiny default window
            # size on its own (looks "hidden" either way) - explicitly
            # give it a normal, visible size/position instead. See
            # _restore_window_normal_size's docstring.
            _restore_window_normal_size(self.document)
        self._prepare_document()

    def reload(self) -> None:
        was_showing = self._is_slideshow_running()
        self.load()
        if was_showing:
            self.start_slideshow()

    def _close_document(self) -> None:
        if self.document is not None:
            try:
                # Our own shape/property mutations (clock hands, tag text,
                # transition suppression, etc.) all set the document's
                # internal Modified flag, even though it was loaded
                # ReadOnly (that only blocks writing back to the ORIGINAL
                # file location, not the in-memory "has this document
                # changed since it was loaded" flag). Explicitly clearing
                # it before close() guarantees no "save changes before
                # closing?" prompt can ever appear for this disposable
                # temp-copy document, regardless of exactly which code
                # path triggers the close.
                self.document.setModified(False)
            except Exception:
                pass
            try:
                self.document.close(False)
            except Exception:
                pass
            self.document = None
        if self._tmp_path is not None:
            try:
                self._tmp_path.unlink(missing_ok=True)
            except Exception:
                pass
            self._tmp_path = None
        self.tracked_shapes = []
        self.clocks = []
        self.multimode_shapes = []
        self._slide_day_mode = {}
        self._slides_by_index = []
        self._advancer.reset()
        self._clock_suppressed = None
        self._pending_transition_slide_index = None
        self._last_known_slide_index = None
        self._cmd_arg_keys = set()
        # NOTE: self._cmd_cache is intentionally NOT reset here - see its
        # field docstring in __init__ for why previously-resolved #CMD
        # values should survive a reload.

    # -- scanning -------------------------------------------------------
    def _prepare_document(self) -> None:
        cholonly_indices: set = set()
        noncholonly_indices: set = set()

        for index, slide, shape in iter_all_shapes(self.document):
            if not is_text_shape(shape):
                continue
            text = shape.getString()
            if "#ANCLOCK" in text:
                clock = create_analog_clock(
                    self.document, slide, shape, self.config.clock_style, self.time_source.now(), index
                )
                self.clocks.append(clock)
            elif "#MULTIMODE" in text:
                shape.setString("")
                shape.Visible = False
                self.multimode_shapes.append(shape)
            elif "#CHOLONLY" in text or "#NONCHOLONLY" in text:
                # Pure marker shapes (like #MULTIMODE) - not meant to
                # display their own literal tag text, just to flag which
                # day-mode(s) apply to the slide they're on. A slide can
                # have both markers (as two separate shapes) - see
                # _reconcile_day_modes for how that's resolved.
                if "#CHOLONLY" in text:
                    cholonly_indices.add(index)
                if "#NONCHOLONLY" in text:
                    noncholonly_indices.add(index)
                shape.setString("")
                shape.Visible = False
            elif tagging.contains_known_tag(text):
                self.tracked_shapes.append(tagging.TrackedShape(UnoTextShape(shape), text))

        slides = self.document.DrawPages
        self._slides_by_index = [slides.getByIndex(i) for i in range(slides.Count)]
        self._slide_day_mode = _reconcile_day_modes(
            cholonly_indices, noncholonly_indices, len(self._slides_by_index)
        )
        for index, slide in enumerate(self._slides_by_index):
            # Read-only at this point - see _suppress_slide_timing's
            # docstring for why every slide must NOT be mutated here, only
            # registered for later lazy suppression once actually shown.
            original = _capture_original_timing(slide)
            self._advancer.register(
                index,
                duration_seconds=float(original.get("Duration") or 0),
                transition_props={name: original[name] for name in _TRANSITION_PROP_NAMES},
            )

        if self._slide_day_mode:
            ctx = build_zman_context(
                now=self.time_source.now(),
                location=self.config.location,
                eretz_yisroel=self.config.eretz_yisroel,
                timezone=self.config.timezone,
            )
            self._apply_day_mode_visibility(ctx.today_is_chol)

        self._cmd_arg_keys = tagging.find_cmd_arg_keys(t.template for t in self.tracked_shapes)

        self._warm_up_slides()
        self._log_prepare_summary()

    def _apply_day_mode_visibility(self, today_is_chol: bool) -> None:
        """Shows/hides every #CHOLONLY/#NONCHOLONLY-tagged slide according
        to whether today counts as chol - see _reconcile_day_modes for how
        each slide's mode was determined, and zman_context.build_zman_context
        for how today_is_chol itself is computed (mirrors the original
        VBA's own Friday-post-chatzos-or-Saturday check).

        Uses Impress's per-slide "Visible" property (the same one PowerPoint
        calls SlideShowTransition.Hidden / the UI's "Hide Slide") - a
        hidden slide is completely skipped by the running slideshow,
        exactly like the original VBA's SlideShowTransition.Hidden-based
        weekday/Shabbos switching in reference/LuahMain.bas, just
        generalized to any number of tagged slides instead of only the
        first/last slide of the deck.

        Whenever a slide is newly HIDDEN here, its native auto-advance/
        transition timing is also immediately suppressed (same as
        _suppress_slide_timing does for the currently-displayed slide).
        This matters because a hidden slide NEVER becomes "current"
        through poll_slide_advance's normal lazy suppression (it's
        excluded from navigation entirely while hidden - see
        _next_visible_slide_index) - so without this, a hidden slide would
        keep its original author-configured Change/Transition properties
        ticking away, completely unmanaged by us, for as long as it stays
        hidden. That live-but-invisible native scheduling was found to
        interfere with the CURRENTLY DISPLAYED slide's own transition
        rendering (observed as inconsistent/flickering transition effects
        that seemed to work correctly on some laps and not others) -
        suppressing it the moment a slide goes hidden closes that gap. Its
        true original values remain safely available in the advancer's
        registry (captured once, read-only, in _prepare_document) for if/
        when the slide is later revealed again - at which point it goes
        through the exact same lazy "entered" suppress-then-eventually-
        restore path as any other slide becoming current for real."""
        for index, mode in self._slide_day_mode.items():
            if not (0 <= index < len(self._slides_by_index)):
                continue
            slide = self._slides_by_index[index]
            desired_visible = today_is_chol if mode == DAY_MODE_CHOL_ONLY else not today_is_chol
            try:
                if slide.Visible != desired_visible:
                    slide.Visible = desired_visible
                    if not desired_visible:
                        _suppress_slide_timing(slide)
            except Exception:
                pass

    def _warm_up_slides(self) -> None:
        """Briefly switches the (off-screen, invisible) editing view's
        current page to every slide in the deck, once, right after
        loading - a heuristic workaround for a suspected Impress rendering
        quirk where a shape on a slide that has never been "visited" in
        any view renders incorrectly (observed: a plain static textbox
        with no tags/animation was invisible the FIRST time its slide was
        shown in the actual slideshow, but rendered correctly on every
        subsequent visit). Pre-visiting every slide's layout once here,
        before the real slideshow ever starts, is intended to force
        Impress to fully compute/cache each slide's rendering ahead of
        time so the live slideshow never has to do it "cold" for the
        first time on any given slide.

        This is NOT confirmed to fully fix the underlying rendering issue
        (untestable in this dev environment - no LibreOffice here) - it's
        a reasonable, low-risk attempt based on the pattern of the
        reported symptom (a plain, completely untouched-by-our-code shape
        only misbehaving on its slide's very first display). If it
        doesn't help, the underlying cause needs to be investigated
        directly on the signage machine (e.g. checking LibreOffice's own
        bug tracker for "text invisible first slide show" type reports,
        or trying a LibreOffice version update).

        IMPORTANT: the editing view's CurrentPage is explicitly reset back
        to the FIRST VISIBLE slide at the end of this method (NOT
        necessarily literal slide index 0, and not left on the last slide
        visited during warm-up) - `start_slideshow()` was observed to
        begin the live show from whatever page the document's controller
        considered "current" rather than always starting at slide 0, so
        leaving it on the last slide caused the slideshow to visibly start
        there instead of at the beginning. CRITICALLY, resetting to a
        HIDDEN slide (e.g. literal index 0 when that slide happens to be
        hidden by #CHOLONLY/#NONCHOLONLY) was found to bias the live
        slideshow into actually starting ON that hidden slide - contrary
        to the assumption (used elsewhere in this file, e.g. for
        mid-show advancing) that Impress always transparently skips
        hidden slides during navigation; that skip-hidden behavior
        apparently does NOT apply to whatever internal mechanism
        `start()`/`gotoFirstSlide()` use to pick the show's starting
        slide - see start_slideshow()'s matching fix and comment for the
        full diagnosis."""
        try:
            controller = self.document.CurrentController
        except Exception:
            return
        for slide in self._slides_by_index:
            try:
                controller.setCurrentPage(slide)
            except Exception:
                pass
        first_visible, _wrapped = _next_visible_slide_index(-1, self._get_slide_visibility_flags())
        if first_visible is not None:
            try:
                controller.setCurrentPage(self._slides_by_index[first_visible])
            except Exception:
                pass

    def _log_prepare_summary(self) -> None:
        """Prints a short summary of what was found on (re)scan - which
        hashtags/templates were recognized as live-updating, and how many
        analog clocks were built. Intended purely as a debugging aid while
        authoring/testing a deck's tags (e.g. to immediately tell whether a
        tag typo/formatting issue prevented a shape from being tracked, vs.
        it just not having reloaded yet) - not required for normal
        operation."""
        chol_only = sum(1 for m in self._slide_day_mode.values() if m == DAY_MODE_CHOL_ONLY)
        noncholonly = sum(1 for m in self._slide_day_mode.values() if m == DAY_MODE_NONCHOL_ONLY)
        print(
            f"luah_signage: loaded presentation - "
            f"{len(self._slides_by_index)} slide(s), "
            f"{len(self.tracked_shapes)} tagged text shape(s), "
            f"{len(self.clocks)} analog clock(s), "
            f"{chol_only} chol-only slide(s), {noncholonly} non-chol-only slide(s), "
            f"{len(self._cmd_arg_keys)} distinct #CMD arg key(s)."
        )
        for tracked in self.tracked_shapes:
            preview = tracked.template.replace("\n", " \\n ")
            print(f"  tag shape: {preview!r}")

    # -- per-tick refresh ------------------------------------------------
    def refresh_clock(self, now: datetime, loop_time: float | None = None) -> None:
        if not self.clocks:
            return

        # Update every clock EXCEPT: (a) clocks on a currently HIDDEN
        # slide (#CHOLONLY/#NONCHOLONLY/#MULTIMODE/manually hidden -
        # there's no point animating a clock nobody can see, and doing so
        # would also fight with that slide's own transition-suppression
        # bookkeeping for no benefit), and (b) the one slide (if any)
        # whose transition properties are currently "live"/restored ahead
        # of an upcoming advance (self._pending_transition_slide_index,
        # set by poll_slide_advance's "prepare_advance" phase) - updating
        # THAT specific slide's clock while its transition is pending was
        # found to leak the transition effect onto the display before the
        # advance actually happens (see round 2's bug A in repo memory).
        #
        # Clocks on all OTHER visible slides (including ones not
        # currently displayed) ARE updated every tick, same as the
        # currently-displayed slide's clock - this was changed from an
        # earlier version that updated ONLY the current slide's clock,
        # because leaving other slides' clocks frozen while off-screen
        # caused a visible "jump" (the hands leaping forward all at once)
        # the next time that slide became current again, since the hands'
        # angles are always computed directly from the current wall-clock
        # time rather than incrementally advanced - see clock_geometry.py.
        visibility = self._get_slide_visibility_flags()
        pending = self._pending_transition_slide_index
        suppressed_index, suppressed_until = self._clock_suppressed or (None, 0.0)

        active_clocks = []
        for clock in self.clocks:
            idx = clock.slide_index
            if idx == pending:
                continue
            if 0 <= idx < len(visibility) and not visibility[idx]:
                continue
            if idx == suppressed_index and loop_time is not None and loop_time < suppressed_until:
                # This slide just became current and its own transition
                # animation may still be playing - see
                # _suppress_clock_for_transition/poll_slide_advance.
                continue
            if not self.config.debug_repaint_update_offscreen_clocks and idx != self._last_known_slide_index:
                # DEBUG-ONLY: reproduces an earlier version's behavior
                # (only ever update the CURRENT slide's clock) - see
                # LuahConfig.debug_repaint_update_offscreen_clocks.
                continue
            active_clocks.append(clock)

        if self._debug and self._last_known_slide_index != self._debug_last_clock_log_index:
            all_indices = [c.slide_index for c in self.clocks]
            active_indices = [c.slide_index for c in active_clocks]
            print(
                f"luah_signage[debug]: refresh_clock current_index={self._last_known_slide_index} "
                f"clock_slide_indices={all_indices} active_indices={active_indices} "
                f"pending={pending} suppressed={self._clock_suppressed}"
            )
            self._debug_last_clock_log_index = self._last_known_slide_index

        if not active_clocks:
            return

        # lockControllers/unlockControllers batches the shape updates into
        # a single repaint. Without this, each hand's PolyPolygon update is
        # individually repainted by the slideshow view, causing a visible
        # flash/flicker every tick (only noticeable in the actual
        # slideshow/presentation window - the Impress editing window does
        # not repaint per-property-change the same way). Can be disabled
        # via LuahConfig.debug_repaint_lock_controllers for diagnosing a
        # machine-specific frozen-clock-hands rendering bug (see that
        # field's docstring).
        use_lock = self.config.debug_repaint_lock_controllers
        if use_lock:
            self.document.lockControllers()
        try:
            for clock in active_clocks:
                update_analog_clock(
                    clock,
                    now,
                    skip_unchanged_writes=self.config.debug_repaint_skip_unchanged_clock_writes,
                    nudge_shape=self.config.debug_repaint_nudge_shape,
                )
        finally:
            if use_lock:
                self.document.unlockControllers()

    def refresh_content(self, now: datetime) -> None:
        if not self.tracked_shapes and not self._slide_day_mode:
            return
        ctx = build_zman_context(
            now=now,
            location=self.config.location,
            eretz_yisroel=self.config.eretz_yisroel,
            timezone=self.config.timezone,
        )
        if self.tracked_shapes:
            use_lock = self.config.debug_repaint_lock_controllers
            if use_lock:
                self.document.lockControllers()
            try:
                for tracked in self.tracked_shapes:
                    tracked.refresh(ctx, cmd_cache=self._cmd_cache)
            finally:
                if use_lock:
                    self.document.unlockControllers()
        if self._slide_day_mode:
            self._apply_day_mode_visibility(ctx.today_is_chol)

    def refresh_cmd_tags(self) -> None:
        """Runs the configured #CMD command once for each distinct
        args-key found in the currently-loaded deck, updating
        self._cmd_cache in place - called on its OWN, much-less-frequent
        tick (config.cmd_tick_seconds, default 300s) from app.py's main
        loop, deliberately SEPARATE from refresh_content's tick (see
        tagging.py's module docstring and LuahConfig.cmd_tag_executable's
        PERFORMANCE NOTE for why). A no-op if no shape in the deck uses a
        #CMD tag at all."""
        if not self._cmd_arg_keys:
            return
        tagging.refresh_cmd_cache(
            self._cmd_cache,
            self.config.cmd_tag_executable,
            self._cmd_arg_keys,
            self.config.cmd_tag_timeout_seconds,
            self.config.cmd_tag_global_args,
        )

    # -- slideshow --------------------------------------------------------
    def is_alive(self) -> bool:
        """Returns whether the underlying document is still open/usable.

        Why this exists: if the user manually closes BOTH the slideshow
        window and the (off-screen, but still a real window/frame - see
        _move_window_offscreen) Impress editing window, the document gets
        disposed. However, most of this class's UNO calls are individually
        wrapped in narrow try/except blocks that silently swallow the
        resulting DisposedException (e.g. `_slideshow_controller` just
        returns None, `poll_slide_advance` then does an early `return`,
        etc.) - so nothing ever propagates up to app.py's tick-level
        exception handler, which is what normally triggers either a
        recovery-reconnect (auto_recover=True) or a clean exit
        (auto_recover=False). Without an explicit liveness check like
        this one, the whole app would silently sit in an infinite,
        completely inert tick loop forever after the user closes both
        windows - never crashing, never recovering, never exiting - which
        is exactly the reported symptom ("the python launcher didn't
        end"). app.py's tick calls this FIRST, before anything else, and
        raises if it's False so the existing recovery/exit machinery
        handles it correctly either way."""
        if self.document is None:
            return False
        try:
            # Any live, non-disposed document exposes a real DrawPages
            # collection with a valid Count - cheap, harmless, and will
            # raise (typically DisposedException) if the document itself
            # has been closed/disposed.
            self.document.DrawPages.Count
            return True
        except Exception:
            return False

    def _is_slideshow_running(self) -> bool:
        if self.document is None:
            return False
        try:
            return self.document.getPresentation().isRunning()
        except Exception:
            return False

    def _slideshow_controller(self):
        """Returns the running slideshow's XSlideShowController (has
        CurrentSlideIndex / gotoNextSlide()), or None if the slideshow
        isn't running (e.g. still starting up)."""
        try:
            return self.document.getPresentation().Controller
        except Exception:
            return None

    def start_slideshow(self) -> None:
        presentation = self.document.getPresentation()
        # We manually drive advancing for every slide via poll_slide_advance
        # (see slide_advance.py / module docstring) - Impress's own native
        # auto-advance/transition scheduling is suppressed for every slide
        # in _prepare_document, so this call itself only starts the show;
        # it does not hand off any per-slide timing to Impress.
        presentation.Display = self.config.display_number
        presentation.start()
        # Belt-and-suspenders: explicitly force the show to begin on the
        # first VISIBLE slide (NOT necessarily literal slide index 0 -
        # see _warm_up_slides' matching fix), regardless of which page
        # the (off-screen) editing view's controller considered "current"
        # beforehand - start() was observed to sometimes begin the show
        # from that leftover current-page state instead.
        #
        # CRITICAL DIAGNOSIS (confirmed via debug_slide_advance logging):
        # `gotoFirstSlide()` does NOT skip hidden slides the way
        # `gotoNextSlide()`/`gotoSlideIndex()` do during an ALREADY-
        # RUNNING show's normal advance - it goes to literal slide index
        # 0 regardless of its Visible state. When slide 0 happened to be
        # hidden (e.g. a #CHOLONLY slide on a non-chol day), calling
        # `gotoFirstSlide()` here caused the live show to actually start
        # ON that hidden slide (observed staying there, oscillating
        # briefly, then only correcting itself once our own
        # DynamicSlideAdvancer's registered duration for that slide
        # elapsed and manually drove an advance away from it) - NOT
        # skipping forward to the real first visible slide the way
        # `gotoNextSlide()` reliably does mid-show. FIXED: compute the
        # first VISIBLE slide's index ourselves (reusing
        # `_next_visible_slide_index(-1, ...)`, which naturally starts
        # its scan at index 0) and navigate to it via `_goto_slide`
        # (which itself prefers `gotoSlideIndex`, with the old
        # gotoFirstSlide/gotoNextSlide pair only as a fallback).
        controller = self._slideshow_controller()
        if controller is not None:
            first_visible, wrapped = _next_visible_slide_index(-1, self._get_slide_visibility_flags())
            if self._debug:
                print(f"luah_signage[debug]: start_slideshow navigating to first_visible={first_visible}")
            if first_visible is not None:
                self._goto_slide(controller, first_visible, wrapped)
                if self._debug:
                    try:
                        landed = controller.CurrentSlideIndex
                    except Exception:
                        landed = "?"
                    print(f"luah_signage[debug]: start_slideshow landed on CurrentSlideIndex={landed}")

    def stop_slideshow(self) -> None:
        if not self._is_slideshow_running():
            return
        try:
            self.document.getPresentation().end()
        except Exception:
            pass

    def poll_slide_advance(self, now: float) -> None:
        """Call once per tick (with a monotonic time value) to drive manual
        slide advancement. See slide_advance.DynamicSlideAdvancer and this
        module's docstring for why this is necessary."""
        controller = self._slideshow_controller()
        if controller is None:
            return
        try:
            current_index = controller.CurrentSlideIndex
        except Exception:
            return
        # Recorded for refresh_clock to reuse (see
        # self._last_known_slide_index's docstring in __init__ for why
        # this must be the SAME read used by both, rather than each
        # independently re-querying CurrentSlideIndex). app.py always
        # calls poll_slide_advance before refresh_clock every tick, so
        # this is guaranteed fresh by the time refresh_clock reads it.
        self._last_known_slide_index = current_index

        action = self._advancer.on_tick(current_index, now)
        if self._debug and action != "none":
            print(f"luah_signage[debug]: poll_slide_advance current_index={current_index} action={action}")
        if action == "entered":
            # First time this slide becomes current (either the show just
            # started, or it's a previously-visited slide the show looped
            # back to) - suppress its native auto-advance/transition NOW,
            # lazily, right as it's about to be shown (see
            # _suppress_slide_timing's docstring for why this must NOT be
            # done any earlier, e.g. preemptively for every slide at
            # document-load time).
            self._pending_transition_slide_index = None
            if 0 <= current_index < len(self._slides_by_index):
                _suppress_slide_timing(self._slides_by_index[current_index])
        elif action == "prepare_advance":
            # A slide's transition effect belongs to the slide being
            # ARRIVED AT, not the one being left - so restore the NEXT
            # VISIBLE slide's (the one we're about to actually advance
            # into - skipping any hidden slides, exactly like Impress's
            # own gotoNextSlide() does; see _next_visible_slide_index)
            # transition properties now, while we still have
            # transition_settle_seconds before the actual advance to let
            # the property change propagate. Restoring the CURRENT
            # (outgoing) slide's own transition here would be a no-op as
            # far as this advance is concerned - it only pays off a full
            # lap later, the next time we happen to revisit this same
            # slide (which is exactly what produced the earlier
            # "transitions only work from the 2nd loop onward" symptom).
            next_index, _wrapped = _next_visible_slide_index(current_index, self._get_slide_visibility_flags())
            if self._debug:
                print(f"luah_signage[debug]: prepare_advance next_index={next_index} wrapped={_wrapped}")
            self._pending_transition_slide_index = next_index
            if next_index is not None:
                next_slide = self._slides_by_index[next_index]
                _restore_slide_transition(next_slide, self._advancer.get_transition_props(next_index))
        elif action == "advance":
            next_index, wrapped = _next_visible_slide_index(current_index, self._get_slide_visibility_flags())
            if self._debug:
                print(f"luah_signage[debug]: advance from={current_index} next_index={next_index} wrapped={wrapped}")
            try:
                self._goto_slide(controller, next_index, wrapped)
                self._pending_transition_slide_index = None
                if self._debug:
                    try:
                        arrived = controller.CurrentSlideIndex
                    except Exception:
                        arrived = "?"
                    print(f"luah_signage[debug]: after goto, CurrentSlideIndex={arrived} (expected {next_index})")
                if next_index is not None:
                    # Read the transition duration BEFORE re-suppressing
                    # (which zeroes it back out) - this is the real,
                    # authored duration that was restored during
                    # "prepare_advance" and just played out via the
                    # goto call above.
                    transition_duration = self._advancer.get_transition_props(next_index).get(
                        "TransitionDuration", 0.0
                    )
                    self._suppress_clock_for_transition(next_index, now, transition_duration)
                    # Immediately re-suppress the slide we just arrived at
                    # (its transition property was only just restored, and
                    # already "used" by the goto call above) so that any
                    # content refresh later in THIS SAME tick or the next
                    # one doesn't cause Impress to replay the transition
                    # again on that shape update.
                    _suppress_slide_timing(self._slides_by_index[next_index])
            finally:
                self._advancer.left_slide()

    def _goto_slide(self, controller, next_index: int | None, wrapped: bool) -> None:
        """Navigates the running slideshow to `next_index`.

        Prefers `controller.gotoSlideIndex(next_index)` - a direct "jump to
        this exact slide" call - over the `wrapped ? gotoFirstSlide() :
        gotoNextSlide()` pair used in earlier versions of this code. The
        direct-index call sidesteps a subtlety of `gotoFirstSlide()`
        specifically: when navigating forward requires SKIPPING one or
        more hidden slides near the end of the deck (i.e. `wrapped=True`,
        needed there purely to avoid literally ending the show by calling
        gotoNextSlide() past the last slide - see _next_visible_slide_index),
        `gotoFirstSlide()` was found to sometimes land on the correct
        VISIBLE slide (Impress auto-skips hidden slide 0 forward) but
        WITHOUT playing that slide's transition effect at all - as if
        Impress's "start from the beginning, skip forward past hidden
        slides" code path treats it as a silent cut, distinct from a
        single-step advance. `gotoSlideIndex` names the destination
        directly, without going through that "start from slide 0" code
        path, so it should trigger the landing slide's transition the same
        way any ordinary single-step navigation does.

        Falls back to the old `wrapped ? gotoFirstSlide() : gotoNextSlide()`
        pair if `gotoSlideIndex` isn't available on this LibreOffice
        version (older releases may lack it) or if it raises for any
        reason - `wrapped` is still needed for that fallback path, to
        avoid ending the show via gotoNextSlide() past the last slide."""
        if next_index is None:
            return
        try:
            controller.gotoSlideIndex(next_index)
            if self._debug:
                print(f"luah_signage[debug]: _goto_slide used gotoSlideIndex({next_index})")
            return
        except Exception as exc:
            if self._debug:
                print(f"luah_signage[debug]: gotoSlideIndex failed ({exc!r}), falling back")
        if wrapped:
            controller.gotoFirstSlide()
            if self._debug:
                print("luah_signage[debug]: _goto_slide used gotoFirstSlide() fallback")
        else:
            controller.gotoNextSlide()
            if self._debug:
                print("luah_signage[debug]: _goto_slide used gotoNextSlide() fallback")

    def _get_slide_visibility_flags(self) -> list:
        """Reads the LIVE `Visible` property of every slide, right now -
        used by `poll_slide_advance` (via `_next_visible_slide_index`) to
        determine which slide the live show will actually navigate to
        next, correctly accounting for slides hidden by ANY mechanism
        (#CHOLONLY/#NONCHOLONLY day-mode switching, the deprecated
        #MULTIMODE, or simply the deck author manually hiding a slide) -
        not just the ones this code itself is aware of/controls."""
        flags = []
        for slide in self._slides_by_index:
            try:
                flags.append(bool(slide.Visible))
            except Exception:
                flags.append(True)
        return flags

    def _suppress_clock_for_transition(self, slide_index: int, now: float, transition_duration_seconds) -> None:
        """Pauses refresh_clock's hand updates for `slide_index`'s clock (if
        it has one) for `transition_duration_seconds` (rounded UP to the
        next whole second, e.g. 0.5s -> 1s) starting from `now` (a
        time.monotonic()-comparable value) - avoids that specific slide's
        clock-hand shape update landing in the middle of its OWN
        still-playing transition animation, which was observed to cause a
        visible flicker. Scoped to a single slide (not a global pause), so
        other visible slides' clocks keep ticking normally in the
        background - see refresh_clock. A duration of 0 (no transition
        effect, e.g. a plain cut) results in no suppression at all."""
        duration = float(transition_duration_seconds or 0.0)
        if duration <= 0:
            self._clock_suppressed = None
            return
        suppress_until = now + math.ceil(duration)
        self._clock_suppressed = (slide_index, suppress_until)

