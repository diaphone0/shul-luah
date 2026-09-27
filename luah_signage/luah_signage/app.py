"""
Main entrypoint: connects to LibreOffice, loads the presentation, starts the
slideshow, and runs the refresh/reload loop.

Run with: python -m luah_signage.app
"""
from __future__ import annotations

import time
import traceback

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


def _connect_and_load(config: LuahConfig, time_source: TimeSource) -> PresentationController:
    ctx = uno_bridge.connect_or_launch(host=config.uno_host, port=config.uno_port)
    controller = PresentationController(config, ctx, time_source)
    controller.load()
    controller.start_slideshow()
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
    watcher = FileChangeWatcher(config.pptx_path)
    last_content_refresh = 0.0
    last_reload_check = 0.0
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

                if loop_time - last_content_refresh >= config.content_tick_seconds:
                    controller.refresh_content(now)
                    last_content_refresh = loop_time

                if loop_time - last_reload_check >= config.reload_poll_seconds:
                    if watcher.check_for_change():
                        print(f"luah_signage: {config.pptx_path} changed on disk, reloading ...")
                        controller.reload()
                    last_reload_check = loop_time

                recovery_backoff = _RECOVERY_BACKOFF_SECONDS  # reset after a healthy tick
                time.sleep(config.clock_tick_seconds)
            except KeyboardInterrupt:
                raise
            except Exception:
                if not config.auto_recover:
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
    config = load_config()
    run(config)


if __name__ == "__main__":
    main()

