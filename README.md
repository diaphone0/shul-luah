# shul-luah (שול-לוח)

A synagogue zmanim (halachic times) / prayer-schedule **digital signage** system — migrated from
Windows + PowerPoint + VBA to a cross-platform **LibreOffice Impress + Python (python-uno)**
solution.

The name is a transliteration of **"לוח בית כנסת"** (shul luach — "synagogue board/calendar"):
`shul` (synagogue) + `luah` (calendar/board/schedule).

## Why this exists

Most digital-signage approaches for a shul either require a dedicated management app/web
service, or lock content editing to a single desk/PC. This project takes a different approach:

**The presentation deck is a real, ordinary PowerPoint (`.pptx`) file.** That's the heart of the
design, not an implementation detail. It means anyone at the shul can open and edit it directly
from the PowerPoint/Microsoft 365 mobile app on their phone, or via OneDrive — tools people
already know, with no custom CMS, login portal, or web dashboard to learn. The signage app itself
just reads that file (via a disposable copy — see "Key architectural decisions" below) and
renders it live through LibreOffice Impress, substituting hashtag placeholders with computed
zmanim/calendar data and driving an analog clock and day-based slide switching.

## Two parts

- **`pyzmanim-lib/`** (Part 1 — **complete**): a standalone Python port of the generic Hebrew
  calendar / zmanim calculation library. No signage/PowerPoint code at all — just the calendar
  math, usable independently in any Python project.
- **`luah_signage/`** (Part 2 — **working, actively refined/hardened**): the LibreOffice Impress
  signage application — hashtag templating, analog clock rendering, automatic slide scheduling,
  and weekday/Shabbos slide switching, all built on top of `pyzmanim-lib`.

## Credits / provenance

This project builds on a chain of earlier open-source work:

- [`diaphone1/vbzmanim`](https://github.com/diaphone1/vbzmanim) (LGPL-2.1) — the VBA/VB6 Hebrew
  calendar & zmanim library this repo's `pyzmanim-lib` is ported from, and the source of the
  original PowerPoint/VBA signage app (`reference/LuahMain.bas`) this project's Part 2 reimagines.
- [`yparitcher/libzmanim`](https://github.com/yparitcher/libzmanim) — which `vbzmanim` itself
  ports from.
- [`KosherJava/Zmanim`](https://github.com/KosherJava/zmanim) ([kosherjava.com](https://kosherjava.com)),
  by Eliyahu Hershfeld — the original Java zmanim calculation library underlying this whole
  lineage.
- [`NykUser/MyZman`](https://github.com/NykUser/MyZman) — source of the ported Daf Yomi cycle
  logic.

## Repo layout

```
shul-luah/
  config.json                 <- active site config (local, not committed — copy from config.example.json)
  config.example.json         <- template to copy from for a fresh deployment
  launch.py                   <- THE way to run the app (see "How to run" below)
  luah.pptx                   <- example/test presentation deck (has #ANCLOCK, tags, etc.)
  pyzmanim-lib/                <- Part 1, self-contained package
    pyzmanim/                  <- the library itself (hebrewcalendar.py, zmanim.py, dafyomi.py, shiur.py, hdateformat.py, noaa_calculator.py, parasha.py, _data/*.py)
    reference/vbzmanim/        <- vendored original VBA source (UTF-8 converted from Windows-1255)
    tests/                     <- pytest-compatible tests, runnable directly with `python tests/test_sanity.py`
    tools/generate_data.py     <- regenerates pyzmanim/_data/*.py from the VBA source (Hebrew text tables)
    pyproject.toml             <- `pip install -e .` from inside pyzmanim-lib/
  luah_signage/                <- Part 2, project root (NOTE: nested package dir below is important!)
    pyproject.toml             <- `pip install -e .` from inside this outer luah_signage/ folder
    luah_signage/               <- the ACTUAL package (must be nested like this - see "Gotchas" below)
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
      _vendor_tzdata/zoneinfo/  <- vendored IANA tz database (648 files, ~500KB) - DO NOT DELETE
    tests/                      <- pytest-compatible, run each with `python tests/test_X.py`
  reference/LuahMain.bas        <- original VBA PowerPoint app source (reference only, not ported literally)
```

## How to run

```powershell
python launch.py
```

`launch.py` is a **cross-platform launcher that requires NO pip installs into LibreOffice's
bundled Python.** LibreOffice ships its own separate Python interpreter (usually a different
version than your system Python), and `import uno` only works under that exact interpreter (or
under a system Python whose ABI happens to match — see `launch.py`'s own fallback for "libs
only, no bundled interpreter" LibreOffice installs). `launch.py`:

1. Locates LibreOffice's `program` directory (via `LIBREOFFICE_PROGRAM_DIR` env var, `soffice` on
   PATH, or OS-specific default install paths).
2. Builds a `PYTHONPATH` containing `pyzmanim-lib/`, `luah_signage/` (the outer project folder),
   and the repo root as plain source directories.
3. Re-launches `-m luah_signage.app` as a subprocess under LibreOffice's own `python.exe`/`python`
   (or, if no bundled interpreter exists, tries running in-process under the system Python
   instead).

**You do NOT need to `pip install` either `pyzmanim` or `luah_signage` for `launch.py` to work.**
(Editable installs via `pip install -e .` in each package's dir are still useful for IDE
tooling/running tests conveniently from a normal terminal, but not required for the app itself.)

## Setup / installation

Detailed, OS-by-OS setup instructions are coming soon. For now, `launch.py` assumes:

1. LibreOffice is already installed (any recent version, Windows or Linux, x86 or ARM).
2. `config.json` exists next to `launch.py` (copy `config.example.json` and edit the
   `pptx_path`, `location`, and `timezone` fields for your site at minimum).
3. Your presentation deck (a normal `.pptx` authored in PowerPoint, containing hashtag
   placeholders — see "Supported hashtags" below) is saved wherever `config.json`'s `pptx_path`
   points (typically a cloud-synced folder like OneDrive, so it can be edited remotely).

Then just run `python launch.py`.

## Key architectural decisions (Part 2)

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
- **`#HIDDEN` slide tag**: a slide tagged `#HIDDEN` is always hidden, on both chol and non-chol
  days — takes precedence over `#CHOLONLY`/`#NONCHOLONLY` if a slide somehow has both, so a
  permanently-hidden slide doesn't need its other day-mode tags removed first.
- **`#CONFIG:{json}` slide tag**: a slide tagged `#CONFIG:{...}` is always hidden (like `#HIDDEN`),
  and the JSON object after the colon overrides one or more `config.json` values for the current
  session — a single key (e.g. `#CONFIG:{"cmd_tick_seconds": 60}`) or the whole config at once.
  Nested `location`/`clock_style` overrides are merged onto the existing values, not replaced
  outright. Only the FIRST `#CONFIG` tag found in the deck is used; any others are ignored.
  Overrides never touch `config.json` on disk and are reset to the file's own values on every
  reload before the current deck's tag (if any) is re-applied — see `apply_config_overrides()`
  in `config.py`.
- **Mockable time** (`time_source.py`): `MockTimeSource` lets you test any date/time scenario
  without touching the OS clock — 1 real second = 1 simulated second, no acceleration.
- **`#CMD:<args>` tag** (`tagging.py`): runs an externally configured script/command (set once in
  `config.json`'s `cmd_tag_executable`) and substitutes the tag with that command's stdout —
  e.g. `#CMD:<chol>` passes `chol` as an argv entry. Args are parsed shell-like (whitespace-
  separated, quoted spans count as one arg) and always passed as a plain argv list (never a shell
  string), so shell metacharacters typed into a shape's tag text can't be interpreted as shell
  syntax. Unlike every other tag, `#CMD` is never run inline during a content refresh — it only
  reads from an in-memory cache, refreshed on its own separate, much less frequent tick
  (`cmd_tick_seconds`, default 300s/5min) — deliberately decoupled from `content_tick_seconds`
  (default 30s) to avoid hammering a possibly slow/rate-limited external command every 30 seconds
  forever. `cmd_tag_global_args` (default `[]`) applies extra argv entries to EVERY `#CMD`
  invocation regardless of args-key — e.g. a shared flag every call should always get — appended
  after `cmd_tag_executable`'s own prefix but before each tag's own args. Give it as a JSON list
  of strings (recommended — each item used verbatim, e.g. `["--sep", " - ", "--swap"]`, no quote
  characters needed) or a shell-like string for convenience. See `cmd_tag_executable`'s docstring
  in `config.py` for the full design rationale.

## Known-tricky UNO gotchas learned the hard way

1. **UNO enums** must be resolved via `uno.Enum("com.sun.star.drawing.LineStyle", "NONE")` — a
   plain `from com.sun.star.drawing import LineStyle` (which works for *structs* like `Point`)
   does NOT work for enums.
2. **`Hidden=True` on `loadComponentFromURL` breaks `XSlideShowController.start()`** — it prevents
   Impress from creating a real window/frame at all, and the slideshow needs one to attach its
   rendering context to. Instead, load `Hidden=False` and then move the window off-screen
   (`window.setPosSize(-32000, -32000, 1, 1, 15)`) — a real but invisible window.
3. **`gotoFirstSlide()` does NOT skip hidden slides** the way `gotoNextSlide()`/`gotoSlideIndex()`
   do during an already-running show. This caused the slideshow to actually START on a hidden
   `#CHOLONLY` slide. **Always compute the first/next VISIBLE slide yourself**
   (`_next_visible_slide_index()`, called with `current_index=-1` for "find the first visible
   slide") and navigate via `_goto_slide()` (which prefers `gotoSlideIndex()`).
4. **Never eagerly mutate every slide's properties at load time** — only touch a slide's
   Change/Transition properties lazily, right when it's about to become current. Doing it
   eagerly for the whole deck broke first-time-shown entrance animations on unrelated slides.
5. **Re-setting a UNO shape property to its own unchanged value can still trigger a repaint** —
   always check "did this actually change?" before writing, for any property touched every tick
   (clock hand `PolyPolygon`, slide `Visible`, etc.).
6. **`presentation.py`/`uno_shapes.py` have NO top-level `import uno`** — all UNO imports are
   done lazily inside the functions that need them. This means **pure-logic helper functions in
   these files CAN be unit-tested without any LibreOffice installed** (e.g.
   `_reconcile_day_modes`, `_next_visible_slide_index` — these files aren't fully untestable just
   because most of the class needs live UNO).
7. **LibreOffice's bundled Python has no IANA tzdata and no `tzdata` pip package by default**,
   especially on portable/locked-down installs. Fixed by vendoring the whole tzdata database into
   `luah_signage/luah_signage/_vendor_tzdata/zoneinfo/` and registering it as a fallback via
   `zoneinfo.reset_tzpath()` in `tzdata_support.py` — works with zero internet/pip/admin rights.
8. **Console `print()` can crash on Hebrew/non-ASCII text** if the console's code page can't
   represent it — `app.py` reconfigures stdout/stderr to UTF-8 with `backslashreplace` at startup;
   `launch.py` also sets `PYTHONIOENCODING` as an earlier redundant safeguard.

## Package structure gotcha (important if adding new Python packages here)

`luah_signage`'s `pyproject.toml` MUST live in an outer folder with the actual package **nested
one level below** (`luah_signage/luah_signage/*.py`), exactly mirroring `pyzmanim-lib`'s layout
(`pyzmanim-lib/pyzmanim/*.py`). Putting `pyproject.toml` directly alongside the module files (flat)
silently produces a broken editable install (`pip install -e .` reports success, but the
generated finder's `MAPPING` dict is empty and `import luah_signage` fails from any cwd other
than one that happens to have it on `sys.path` by coincidence). Always verify a new package's
install by testing `import <pkg>` from an unrelated directory.

## Supported hashtags (full list — see `tagging.py`'s `TAG_REGISTRY`)

`#HEBDATE`, `#PARSHA`, `#DAFYOMI`, `#LIMUDYOMI`, `#DAYZMANIM`, `#FULLZMANIM`, `#ALOS72`,
`#SUNRISE`, `#SHMAGRA`, `#SHMAMGA`, `#TFILAGRA`, `#TFILAMGA`, `#CHAZOT`, `#BGMINHA`, `#LTMINHA`,
`#SUNSET`, `#TZAIS`, `#TZAISYESHIVA`, `#SHABBOS`, `#MOZASH`, `#DICLOCK`. Most accept an optional
`+N`/`-N` (minutes) or `+HH:MM`/`-HH:MM` offset suffix, e.g. `#SUNSET-01:30`. `#CMD:<args>` runs a configured external
command and substitutes its stdout (see `cmd_tag_executable` in `config.example.json`). Special
non-text tags: `#ANCLOCK` (analog clock), `#CHOLONLY`/`#NONCHOLONLY`/`#HIDDEN`/`#CONFIG:{json}`
(day-mode slide visibility + config overrides), `#MULTIMODE` (deprecated/inert, kept only for
backward compat).

## Test suite

69 tests total, all pure-Python (no LibreOffice needed), runnable individually:
```powershell
cd luah_signage
python tests/test_clock_geometry.py
python tests/test_tagging.py
python tests/test_watcher.py
python tests/test_slide_advance.py
python tests/test_presentation_daymode.py
python tests/test_presentation_slide_visibility.py
python tests/test_time_source.py
python tests/test_tzdata_support.py
python tests/test_config.py
```
Plus `pyzmanim-lib/tests/test_sanity.py` (9 tests) for the calendar/zmanim library itself.

Anything requiring a live LibreOffice connection (most of `presentation.py`'s UNO-calling methods,
`uno_shapes.py`, `uno_bridge.py`) is **not unit-testable in a typical dev sandbox** — such changes
need to be validated by actually running `python launch.py` on a real machine with LibreOffice
installed (optionally with `debug_slide_advance: true` in `config.json` for diagnostic prints).

## Current status

- ✅ Core slideshow navigation, hidden-slide skipping, transition timing, and clock-hand rendering
  — confirmed working correctly via live testing.
- ✅ `#CHOLONLY`/`#NONCHOLONLY` day-mode slide visibility — confirmed working live.
- ✅ Cross-machine tzdata handling — fixed and verified.
- ✅ Confirmed working on Windows, Linux, and Android (via Termux + an X server app).
- ⬜ `#MULTIMODE`'s full original behavior beyond what `#CHOLONLY`/`#NONCHOLONLY` already covers
  has not been further developed; `mod_shiur.bas`'s function-level logic was ported early on but
  hasn't been revisited since.
- ⬜ `debug_repaint_*` config toggles exist to help diagnose a machine-specific frozen-clock-hands
  bug reported on a small number of systems — see each field's docstring in `config.py`.

## License

This project is licensed under the **GNU Lesser General Public License v2.1** (LGPL-2.1) — see
[`LICENSE`](LICENSE). This follows the license of the upstream `vbzmanim` project this repo's
`pyzmanim-lib` is ported from (see "Credits / provenance" above for the full attribution chain).
