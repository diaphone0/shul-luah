# luahnew — Project Continuation Notes

_Last updated: 2026-09-29. Read this first when resuming work in a new chat session._

## What this project is

Migrating a synagogue's zmanim (halachic times) / prayer-schedule **digital signage** system
from Windows + PowerPoint + VBA to a cross-platform **LibreOffice Impress + Python (python-uno)**
solution. Source VBA reference: https://github.com/diaphone1/vbzmanim (LGPL-2.1).

Two parts:
- **Part 1 — `pyzmanim-lib/`**: a standalone Python port of the generic Hebrew calendar / zmanim
  calculation library (no signage/PowerPoint code). **Status: COMPLETE.**
- **Part 2 — `luah_signage/`**: the LibreOffice Impress signage application (hashtag templating,
  analog clock, slide scheduling, weekday/Shabbos slide switching). **Status: WORKING, actively
  being refined/hardened.** This is where almost all recent effort has gone.

## Repo layout

```
c:\Intel\luahnew\
  config.json                 <- active site config (gitignored-style local file, currently in DEBUG mode - see below)
  config.example.json         <- template to copy from for a fresh deployment
  launch.py                   <- THE way to run the app (see "How to run" below)
  luah.pptx                   <- the actual test presentation deck (has #ANCLOCK, tags, etc.)
  pyzmanim-lib\                <- Part 1, self-contained package
    pyzmanim\                  <- the library itself (hebrewcalendar.py, zmanim.py, dafyomi.py, shiur.py, hdateformat.py, noaa_calculator.py, parasha.py, _data\*.py)
    reference\vbzmanim\        <- vendored original VBA source (UTF-8 converted from Windows-1255)
    tests\                     <- pytest-compatible tests, runnable directly with `python tests\test_sanity.py`
    tools\generate_data.py     <- regenerates pyzmanim/_data/*.py from the VBA source (Hebrew text tables)
    pyproject.toml             <- `pip install -e .` from inside pyzmanim-lib\
  luah_signage\                 <- Part 2, project root (NOTE: nested package dir below is important!)
    pyproject.toml             <- `pip install -e .` from inside this outer luah_signage\ folder
    luah_signage\               <- the ACTUAL package (must be nested like this - see "Gotchas" below)
      app.py                    <- main tick loop / entrypoint (python -m luah_signage.app)
      config.py                 <- LuahConfig/ClockStyle dataclasses + load_config()
      presentation.py           <- the big one: PresentationController - load/scan/tag-refresh/clock-refresh/slide-advance/day-mode-visibility
      slide_advance.py          <- DynamicSlideAdvancer - pure state machine for manual slide timing/transitions
      tagging.py                <- hashtag scanning & rendering (TAG_REGISTRY, offset parsing)
      clock_geometry.py         <- pure math for analog clock hands (no UNO dependency)
      uno_shapes.py             <- UNO-specific shape adapters + analog clock builder
      uno_bridge.py             <- locates/launches/connects to LibreOffice via UNO socket
      zman_context.py           <- builds ZmanContext (current HDate, today_is_chol, etc.) per refresh
      time_source.py            <- TimeSource/MockTimeSource - mockable wall clock for testing
      tzdata_support.py         <- vendored IANA tzdata fallback (see below)
      watcher.py                <- FileChangeWatcher (mtime polling for cloud-synced pptx changes)
      _vendor_tzdata\zoneinfo\  <- vendored IANA tz database (648 files, ~500KB) - DO NOT DELETE
    tests\                      <- pytest-compatible, run each with `python tests\test_X.py`
  reference\LuahMain.bas        <- original VBA PowerPoint app source (reference only, not ported literally)
```

## How to run

```powershell
cd c:\Intel\luahnew
python launch.py
```

`launch.py` is a **cross-platform launcher that requires NO pip installs into LibreOffice's
bundled Python.** LibreOffice ships its own separate Python interpreter (different version than
your system Python), and `import uno` only works under that exact interpreter. `launch.py`:
1. Locates LibreOffice's `program` directory (via `LIBREOFFICE_PROGRAM_DIR` env var, `soffice` on
   PATH, or OS-specific default install paths).
2. Builds a `PYTHONPATH` containing `pyzmanim-lib\`, `luah_signage\` (the outer project folder),
   and the repo root as plain source directories.
3. Re-launches `-m luah_signage.app` as a subprocess under LibreOffice's own `python.exe`.

**You do NOT need to `pip install` either `pyzmanim` or `luah_signage` for `launch.py` to work.**
(Editable installs via `pip install -e .` in each package's dir are still useful for IDE
tooling/running tests conveniently from a normal terminal, but not required for the app itself.)

## Current `config.json` state — ⚠️ IS IN DEBUG MODE

The live `c:\Intel\luahnew\config.json` currently has these **non-production** debug settings
active (check/reset before any real deployment):

```jsonc
"auto_recover": false,               // exits cleanly on any error instead of retrying forever
"mock_start_datetime": "2026-09-26T14:30:00",  // MOCKED clock - a Saturday, for testing #CHOLONLY/#NONCHOLONLY
"debug_slide_advance": false,        // (currently off, but was used heavily for diagnosis - see below)
"hide_editor_window": false,         // shows the Impress editor window on-screen (normally hidden off-screen)
```

**For production use**, set: `auto_recover: true`, `mock_start_datetime: null`,
`debug_slide_advance: false`, `hide_editor_window: true`.

## `pptx_path` fallback

If `config.json`'s `pptx_path` is missing, `null`, or an empty/blank string, `config.load_config()`
now defaults it to `luah.pptx` next to `launch.py` (i.e. `config.DEFAULT_PPTX_PATH`, computed the
same way as `DEFAULT_CONFIG_PATH` — three `.parent`s up from `config.py`). See `test_config.py`.

## `debug_repaint_*` toggles — diagnosing machine-specific frozen clock hands

Reported symptom: on SOME machines (a few specific x86/ARM systems out of many tested), the
analog clock hands stop visibly moving in the live slideshow window — updates only become visible
after navigating away from and back to the slide — while most machines work flawlessly with the
identical deck/config. Since this can't be reproduced/debugged directly (no access to the affected
hardware), four independent config toggles were added so the user can disable one anti-flicker
tweak at a time and report back which one (if any) fixes it on a given machine:

- `debug_repaint_lock_controllers` (default `true`) — disables `document.lockControllers()` /
  `unlockControllers()` batching in both `refresh_clock` and `refresh_content` when `false`.
- `debug_repaint_skip_unchanged_clock_writes` (default `true`) — disables the "only rewrite a
  hand's `PolyPolygon` if its angle actually changed" optimization in `uno_shapes._refresh_hand`
  when `false` (forces an unconditional rewrite every tick).
- `debug_repaint_update_offscreen_clocks` (default `true`) — reverts to updating ONLY the
  currently-displayed slide's clock (the older, pre-"jump fix" behavior) when `false`.
- `debug_repaint_nudge_shape` (default `false`, EXPERIMENTAL) — when `true`, toggles each updated
  hand shape's `Visible` off/on immediately after rewriting its polygon, to force it out of and
  back into the render tree. Heavier-handed than the others (adds a workaround rather than just
  disabling one); may itself cause a brief visible blink.

All four are threaded through `LuahConfig` → `PresentationController.refresh_clock`/
`refresh_content` → `uno_shapes.update_analog_clock`/`_refresh_hand`. See each field's docstring
in `config.py` for full rationale. Recommended debugging approach: on an affected machine, flip
ONE flag to its non-default value at a time, relaunch, and observe whether the hands start moving
normally in the slideshow window.

## Key architectural decisions (part 2)

- **The source `.pptx` is NEVER opened directly by LibreOffice.** Every `load()`/`reload()` copies
  it to a disposable temp file first (`shutil.copy2`) and opens only that copy — guarantees a
  cloud-sync client can update the real file at any time without being blocked by a file lock.
- **Slide advancement is entirely manual**, driven by `slide_advance.DynamicSlideAdvancer` (a pure,
  fully-unit-tested state machine) + `PresentationController.poll_slide_advance()`. Impress's own
  native "advance after N seconds" + transition scheduling is suppressed per-slide, **lazily**
  (only right when a slide is about to be shown — never preemptively for the whole deck at load
  time, which was found to break entrance animations on a slide's first-ever appearance).
- **Hashtag templating** (`tagging.py`): shapes containing `#TAG` (optionally `+HH:MM`/`-HH:MM` or
  `+N`/`-N` minutes offset, e.g. `#SUNSET-01:30`) are re-rendered from an in-memory template every
  refresh — never mutating the "true" template text, unlike the original VBA which had to stash/
  restore original text via a marker-prefix hack.
- **Analog clock** (`#ANCLOCK` tag): hand shapes are rebuilt (`PolyPolygon` recomputed) every tick
  from `clock_geometry.py`'s pure trig — NOT via UNO's `RotateAngle` (that rotates around a shape's
  own bounding-box center, which looked wrong for asymmetric hands).
- **`#CHOLONLY` / `#NONCHOLONLY` slide tags** (replaces the old VBA `#MULTIMODE` concept): a slide
  tagged `#CHOLONLY` shows only on regular weekdays; `#NONCHOLONLY` shows only on Shabbos/Yom Tov;
  no tag or both tags = always shown. Implemented via `_reconcile_day_modes()` (pure, tested) +
  `_apply_day_mode_visibility()` (sets Impress's `slide.Visible` property).
- **Mockable time** (`time_source.py`): `MockTimeSource` lets you test any date/time scenario
  without touching the OS clock — 1 real second = 1 simulated second, no acceleration.

## Known-tricky UNO gotchas learned the hard way (see repo memory for full blow-by-blow)

1. **UNO enums** must be resolved via `uno.Enum("com.sun.star.drawing.LineStyle", "NONE")` — a
   plain `from com.sun.star.drawing import LineStyle` (which works for *structs* like `Point`)
   does NOT work for enums.
2. **`Hidden=True` on `loadComponentFromURL` breaks `XSlideShowController.start()`** — it prevents
   Impress from creating a real window/frame at all, and the slideshow needs one to attach its
   rendering context to. Instead, load `Hidden=False` and then move the window off-screen
   (`window.setPosSize(-32000, -32000, 1, 1, 15)`) — a real but invisible window.
3. **`gotoFirstSlide()` does NOT skip hidden slides** the way `gotoNextSlide()`/`gotoSlideIndex()`
   do during an already-running show. This caused the slideshow to actually START on a hidden
   `#CHOLONLY` slide — the root cause of a whole saga of "erratic transition/clock" bugs that took
   ~6 rounds of debugging to trace back to this one call. **Always compute the first/next VISIBLE
   slide yourself** (`_next_visible_slide_index()`, called with `current_index=-1` for "find the
   first visible slide") and navigate via `_goto_slide()` (which prefers `gotoSlideIndex()`).
4. **Never eagerly mutate every slide's properties at load time** — only touch a slide's
   Change/Transition properties lazily, right when it's about to become current. Doing it
   eagerly for the whole deck broke first-time-shown entrance animations on unrelated slides.
5. **Re-setting a UNO shape property to its own unchanged value can still trigger a repaint** —
   always check "did this actually change?" before writing, for any property touched every tick
   (clock hand `PolyPolygon`, slide `Visible`, etc.).
6. **`presentation.py`/`uno_shapes.py` have NO top-level `import uno`** — all UNO imports are
   done lazily inside the functions that need them. This means **pure-logic helper functions in
   these files CAN be unit-tested without any LibreOffice installed** (e.g.
   `_reconcile_day_modes`, `_next_visible_slide_index` — don't assume these files are fully
   untestable just because most of the class needs live UNO).
7. **LibreOffice's bundled Python has no IANA tzdata and no `tzdata` pip package by default**,
   especially on portable/locked-down installs. Fixed by vendoring the whole tzdata database into
   `luah_signage/luah_signage/_vendor_tzdata/zoneinfo/` and registering it as a fallback via
   `zoneinfo.reset_tzpath()` in `tzdata_support.py` — works with zero internet/pip/admin rights.
8. **Console `print()` can crash on Hebrew/non-ASCII text** if the console's code page can't
   represent it — `app.py` reconfigures stdout/stderr to UTF-8 with `backslashreplace` at startup;
   `launch.py` also sets `PYTHONIOENCODING` as an earlier redundant safeguard.

## Package structure gotcha (IMPORTANT if adding new Python packages here)

`luah_signage`'s `pyproject.toml` MUST live in an outer folder with the actual package **nested
one level below** (`luah_signage/luah_signage/*.py`), exactly mirroring `pyzmanim-lib`'s layout
(`pyzmanim-lib/pyzmanim/*.py`). Putting `pyproject.toml` directly alongside the module files (flat)
silently produces a broken editable install (`pip install -e .` reports success, but the
generated finder's `MAPPING` dict is empty and `import luah_signage` fails from any cwd other
than one that happens to have it on `sys.path` by coincidence). This bit us once already — always
verify a new package's install by testing `import <pkg>` from an unrelated directory.

## Supported hashtags (full list — see `tagging.py`'s `TAG_REGISTRY`)

`#HEBDATE`, `#PARSHA`, `#DAFYOMI`, `#LIMUDYOMI`, `#DAYZMANIM`, `#FULLZMANIM`, `#ALOS72`,
`#SUNRISE`, `#SHMAGRA`, `#SHMAMGA`, `#TFILAGRA`, `#TFILAMGA`, `#CHAZOT`, `#BGMINHA`, `#LTMINHA`,
`#SUNSET`, `#TZAIS`, `#SHABBOS`, `#MOZASH`, `#DICLOCK`. Most accept an optional `+N`/`-N` (minutes)
or `+HH:MM`/`-HH:MM` offset suffix, e.g. `#SUNSET-01:30`. Special non-text tags: `#ANCLOCK`
(analog clock), `#CHOLONLY`/`#NONCHOLONLY` (day-mode slide visibility), `#MULTIMODE` (deprecated/
inert, kept only for backward compat).

## Test suite

69 tests total, all pure-Python (no LibreOffice needed), runnable individually:
```powershell
cd c:\Intel\luahnew\luah_signage
python tests\test_clock_geometry.py
python tests\test_tagging.py
python tests\test_watcher.py
python tests\test_slide_advance.py
python tests\test_presentation_daymode.py
python tests\test_presentation_slide_visibility.py
python tests\test_time_source.py
python tests\test_tzdata_support.py
python tests\test_config.py
```
Plus `pyzmanim-lib\tests\test_sanity.py` (9 tests) for the calendar/zmanim library itself.

Anything requiring a live LibreOffice connection (most of `presentation.py`'s UNO-calling methods,
`uno_shapes.py`, `uno_bridge.py`) is **not unit-testable in a typical dev sandbox** — those changes
have historically been validated by the user running `python launch.py` on their real machine and
reporting back observed behavior (often with `debug_slide_advance: true` to get diagnostic prints).

## Current status / what's confirmed working

- ✅ Core slideshow navigation, hidden-slide skipping, transition timing, and clock-hand rendering
  — confirmed working correctly via live debug-log evidence (not just code review) as of the last
  session.
- ✅ `#CHOLONLY`/`#NONCHOLONLY` day-mode slide visibility — confirmed working live.
- ✅ Cross-machine tzdata issue — fixed and verified via direct reproduction in this dev environment
  (not yet re-confirmed by the user on the actual second PC).
- ⚠️ Not yet independently reconfirmed on real hardware since the last two fixes: (1) app process
  exiting cleanly when both windows are manually closed (`PresentationController.is_alive()`), and
  (2) clock hands on non-active-but-visible slides no longer "jumping" after being off-screen.
- ⬜ Not yet started: any further work on `#MULTIMODE`'s full original behavior beyond what
  `#CHOLONLY`/`#NONCHOLONLY` already covers; `mod_shiur.bas`'s function-level logic was ported early
  on but hasn't been revisited since.
- ⬜ NEW, not yet confirmed by user: `debug_repaint_*` toggles just added to help bisect a
  machine-specific frozen-clock-hands bug reported on some x86/ARM systems — awaiting user testing
  with each flag flipped individually on an affected machine.

## If you're picking this up in a new session

1. Read this file first, then check `/memories/repo/luahnew_pyzmanim.md` (agent memory) for the
   full blow-by-blow history of every bug/fix if you need more forensic detail than the summary
   above.
2. Check `config.json`'s current debug flags before assuming "production-ready" behavior.
3. Run the test suite to confirm nothing regressed since last session.
4. If the user reports a new UNO/LibreOffice-specific bug, remember you can't test it directly —
   reason carefully from code, but don't hesitate to add `debug_slide_advance`-style targeted
   logging and ask for a fresh log rather than guessing repeatedly (that's what eventually cracked
   the hardest bug in this project's history).
