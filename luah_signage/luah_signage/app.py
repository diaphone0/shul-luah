"""
Main entrypoint: connects to LibreOffice, loads the presentation, starts the
slideshow, and runs the refresh/reload loop.

Run with: python -m luah_signage.app
"""
from __future__ import annotations

import sys
import time
import traceback
from pathlib import Path

from . import uno_bridge
from .config import LuahConfig, load_config
from .presentation import PresentationController
from .time_source import TimeSource, create_time_source
from .watcher import FileChangeWatcher

# Unattended-signage resilience: if the LibreOffice window/document is
# closed (accidentally by someone at the site, a crash, a Windows update
# restart, etc.), UNO calls start failing with errors like "Binary URP
# bridge already disposed" (com.sun.star.uno.RuntimeException) or
# com.sun.star.lang.DisposedException. Rather than exit, we treat ANY
# unexpected exception from a tick as "the display needs to be recovered",
# and try to fully reconnect/relaunch + reload from scratch, with a backoff
# so a persistent failure doesn't spin the CPU or spam retries too fast.
_RECOVERY_BACKOFF_SECONDS = 5.0
_MAX_RECOVERY_BACKOFF_SECONDS = 60.0


def _ensure_console_can_print_any_unicode() -> None:
    """Reconfigures stdout/stderr to encode as UTF-8, replacing any
    character the terminal can't display rather than raising.

    Why this matters: several debug/diagnostic `print()` calls throughout
    this app (e.g. presentation.py's `_log_prepare_summary`, which prints
    the raw text of every tagged shape found in the deck) can legitimately
    contain non-ASCII text - Hebrew, in this app's actual use case, since
    a deck author will often mix a Hebrew label directly in the same
    textbox as a hashtag (e.g. "הדלקת נרות: #SHABBOS"). Depending on the
    machine's active console/terminal code page (which can vary by
    Windows locale, terminal emulator, whether output is being
    piped/redirected, etc. - NOT reliably "will support Hebrew just
    because the OS locale is Hebrew/Israel"), printing such text can
    raise `UnicodeEncodeError` and crash an otherwise-healthy tick,
    triggering an unnecessary reconnect/recovery cycle (or an outright
    exit if `auto_recover=False`). Reconfiguring to UTF-8 with
    `errors="backslashreplace"` guarantees `print()` never raises for
    this reason on any platform/console, at the cost of some non-encodable
    characters showing as `\\uXXXX` escapes in the console instead of
    their real glyph - an acceptable tradeoff for a diagnostic message.
    `reconfigure` is only available on real file-like stdout/stderr
    objects (Python 3.7+) - guarded by try/except since it's not
    guaranteed to exist in every possible embedding/execution context."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
        except Exception:
            pass


def _connect_and_load(config: LuahConfig, time_source: TimeSource) -> PresentationController:
    ctx = uno_bridge.connect_or_launch(host=config.uno_host, port=config.uno_port)
    controller = PresentationController(config, ctx, time_source)
    controller.load()
    # DEBUGGING ONLY: debug_skip_slideshow=True skips starting the
    # fullscreen slideshow entirely, leaving only the Impress editing
    # window open (pair with hide_editor_window=False, otherwise nothing
    # would be visible at all) - useful for inspecting shapes/tag
    # substitution without a fullscreen presentation covering the screen.
    # Every other refresh mechanism (clock/content ticks, #CMD cache,
    # file-change reload) still runs completely normally either way - see
    # LuahConfig.debug_skip_slideshow's docstring. Reads from
    # controller.config (not the outer `config` parameter) so a deck's
    # own "#CONFIG:{...}" override of this flag takes effect too, same as
    # every other per-tick setting in this module.
    if not controller.config.debug_skip_slideshow:
        controller.start_slideshow()
    # Populate the #CMD cache once, right away - WITHOUT this, any #CMD tag
    # would render as an empty string for the entire first content_tick_
    # seconds interval (the main loop's cmd-refresh tick also fires on its
    # own first iteration - see run()'s last_cmd_refresh=0.0 - but that
    # happens AFTER refresh_content already ran blank on that same first
    # tick, so the resolved value wouldn't actually reach the display
    # until the SECOND content refresh without this upfront call).
    controller.refresh_cmd_tags()
    return controller


def run(config: LuahConfig) -> None:
    # Created ONCE here (not inside _connect_and_load) so a mocked start
    # time (config.mock_start_datetime) keeps advancing continuously
    # across recovery reconnects instead of resetting back to its
    # configured start value every time the display connection is lost
    # and re-established - see time_source.py.
    time_source = create_time_source(config.mock_start_datetime)
    if config.mock_start_datetime:
        print(f"luah_signage: using MOCKED time, starting at {config.mock_start_datetime}")

    controller = _connect_and_load(config, time_source)
    # Watches whichever pptx_path is CURRENTLY effective (controller.config,
    # not the outer `config` parameter) - if a deck's #CONFIG tag overrode
    # pptx_path during that load, we want to watch the file actually being
    # displayed, not the original config.json path. See the matching
    # watcher.retarget() call in the recovery branch below for why this
    # needs to stay in sync across reconnects too.
    watcher = FileChangeWatcher(controller.config.pptx_path)
    last_content_refresh = 0.0
    last_reload_check = 0.0
    # Set to "now" (not 0.0, unlike the other last_*_refresh trackers) -
    # _connect_and_load already ran refresh_cmd_tags() once immediately
    # above, so the in-loop cmd-tick check should wait a FULL
    # cmd_tick_seconds from here before refreshing again, rather than
    # firing again right away on the very first loop iteration.
    last_cmd_refresh = time.monotonic()
    recovery_backoff = _RECOVERY_BACKOFF_SECONDS

    try:
        while True:
            try:
                now = time_source.now()
                loop_time = time.monotonic()

                # Detect the document having been closed/disposed (e.g.
                # the user manually closed both the slideshow window and
                # the off-screen editing window) BEFORE doing anything
                # else this tick. Most of PresentationController's UNO
                # calls individually swallow the resulting exception, so
                # without this explicit check the app would otherwise
                # sit in an inert, never-ending tick loop forever once
                # the document is gone - see PresentationController.
                # is_alive's docstring for the full explanation.
                if not controller.is_alive():
                    raise RuntimeError("presentation document is no longer alive (closed by user?)")

                # poll_slide_advance runs FIRST, before any content refresh:
                # when it just advanced to a new slide, it immediately
                # re-suppresses that slide's transition property (see
                # presentation.py). If a content refresh ran before this on
                # the same tick, it would fire while the just-played
                # arrival transition is still "live" on the property,
                # causing Impress to replay/flash the transition again on
                # that shape update - visible as a flash on the very first
                # clock-hand/hashtag update after a slide change.
                controller.poll_slide_advance(loop_time)
                controller.refresh_clock(now, loop_time)

                # Read every per-tick setting below from controller.config
                # (the EFFECTIVE config for whatever's currently loaded),
                # NOT the outer `config` parameter (the original, unchanging
                # file-loaded config) - a deck's "#CONFIG:{...}" tag (see
                # presentation.py's _prepare_document) can override any of
                # these at load/reload time, and reading the stale outer
                # `config` here would silently ignore that override for the
                # whole tick-scheduling loop even though the controller
                # itself picked it up correctly.
                if loop_time - last_content_refresh >= controller.config.content_tick_seconds:
                    controller.refresh_content(now)
                    last_content_refresh = loop_time

                if loop_time - last_reload_check >= controller.config.reload_poll_seconds:
                    if watcher.check_for_change():
                        print(f"luah_signage: {watcher.path} changed on disk, reloading ...")
                        controller.reload()
                        # The reload may have picked up a NEW (or removed a
                        # previous) #CONFIG pptx_path override - keep the
                        # watcher pointed at whatever's effective now.
                        if Path(controller.config.pptx_path) != watcher.path:
                            watcher.retarget(controller.config.pptx_path)
                    last_reload_check = loop_time

                # Deliberately on its OWN, much-less-frequent tick
                # (cmd_tick_seconds, default 300s) rather than piggy-
                # backing on content_tick_seconds - see tagging.py's
                # module docstring and LuahConfig.cmd_tag_executable's
                # PERFORMANCE NOTE for why: running a possibly slow/
                # network-bound #CMD command every content_tick_seconds
                # (default 30s), forever, risks tripping rate limits on
                # whatever remote service it calls. A no-op if the deck
                # has no #CMD tags at all.
                if loop_time - last_cmd_refresh >= controller.config.cmd_tick_seconds:
                    controller.refresh_cmd_tags()
                    last_cmd_refresh = loop_time

                recovery_backoff = _RECOVERY_BACKOFF_SECONDS  # reset after a healthy tick
                time.sleep(controller.config.clock_tick_seconds)
            except KeyboardInterrupt:
                raise
            except Exception:
                if not controller.config.auto_recover:
                    # Debug/dev mode (auto_recover=False): let the app exit
                    # normally instead of retrying, so closing LibreOffice
                    # or Ctrl+C-ing during a dev session cleanly ends the
                    # process rather than looping recovery attempts.
                    print(
                        "luah_signage: display connection lost and "
                        "auto_recover is disabled - exiting.",
                    )
                    traceback.print_exc()
                    return
                print(
                    "luah_signage: display connection lost, attempting to "
                    f"recover in {recovery_backoff:.0f}s ...",
                )
                traceback.print_exc()
                time.sleep(recovery_backoff)
                recovery_backoff = min(recovery_backoff * 2, _MAX_RECOVERY_BACKOFF_SECONDS)
                try:
                    controller = _connect_and_load(config, time_source)
                    last_content_refresh = 0.0
                    last_reload_check = 0.0
                    last_cmd_refresh = time.monotonic()  # see run()'s matching comment above
                    if Path(controller.config.pptx_path) != watcher.path:
                        watcher.retarget(controller.config.pptx_path)
                except Exception:
                    print("luah_signage: recovery attempt failed, will retry.")
                    traceback.print_exc()
    except KeyboardInterrupt:
        pass
    finally:
        try:
            controller.stop_slideshow()
        except Exception:
            pass


def main() -> None:
    _ensure_console_can_print_any_unicode()
    config = load_config()
    run(config)


if __name__ == "__main__":
    main()

