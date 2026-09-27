#!/usr/bin/env python3
"""
Cross-platform launcher for luah_signage that does NOT require pip-installing
anything into LibreOffice's bundled Python interpreter.

Why this is needed: LibreOffice's `uno`/`pyuno` module is a native extension
tightly ABI-locked to the exact Python build it was compiled against (e.g. it
may ship Python 3.13 while your system Python is 3.14) - it can only be
imported when running under a MATCHING interpreter. Meanwhile `pip install`
INTO LibreOffice's bundled Python often fails on Windows (Program Files
write permissions block pip's build-isolation temp dirs). This script
sidesteps both problems: it locates LibreOffice's bundled Python, adds the
pyzmanim-lib and luah_signage source directories (and, on Windows, an
existing `tzdata` install) onto PYTHONPATH, and re-launches
`luah_signage.app` under that interpreter as a subprocess - no pip install
or venv required for either package.

Fallback for "libs only, no bundled interpreter" LibreOffice installs: some
installs (notably Linux distro packages like Debian/Ubuntu's `python3-uno`)
ship ONLY the `uno`/`pyuno` native module compiled against the SYSTEM
Python, with no separate bundled `python` executable at all - there is
nothing to re-exec under in that case. If no bundled interpreter binary is
found, this script instead tries `import uno` directly under the CURRENT
(system) Python that's running this very script (via
uno_bridge.ensure_uno_importable(), which adds LibreOffice's program
directory to sys.path first) and, if that succeeds, runs the app directly
in-process rather than launching a subprocess. This only works if the
system Python's ABI (version + build) happens to match what pyuno was
compiled against - if it doesn't, importing raises ImportError and this
script reports a clear error instead of silently failing.

Usage: python launch.py
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent
PYZMANIM_LIB = REPO_ROOT / "pyzmanim-lib"
LUAH_SIGNAGE = REPO_ROOT / "luah_signage"

# find_libreoffice_program_dir() only uses os/shutil/glob (no `import uno` at
# module scope in uno_bridge.py), so it's safe to import and call under the
# system Python that's running this launcher script.
sys.path.insert(0, str(LUAH_SIGNAGE))
from luah_signage.uno_bridge import ensure_uno_importable, find_libreoffice_program_dir  # noqa: E402


def _find_tzdata_dir() -> str | None:
    """Locates an already-installed `tzdata` package's parent directory
    (under whichever Python is running this launcher) so it can be added to
    PYTHONPATH for LibreOffice's bundled Python too. `tzdata` is pure
    Python/data with no compiled extension, so it's safe to share across
    interpreter versions this way. Only needed on Windows - Linux systems
    normally already have the IANA tz database available system-wide."""
    try:
        import tzdata
    except ImportError:
        return None
    return str(Path(tzdata.__file__).resolve().parent.parent)


def _run_in_process_under_system_python() -> None:
    """Fallback used when LibreOffice has no bundled Python executable to
    re-exec under (see module docstring) - tries to import `uno` directly
    under THIS (system) interpreter and, if that succeeds, runs the app
    directly in the current process instead of via subprocess."""
    sys.path.insert(0, str(PYZMANIM_LIB))
    try:
        ensure_uno_importable()
    except Exception as exc:
        print(
            "Could not import LibreOffice's 'uno' module under this "
            f"system Python either ({exc!r}). LibreOffice's pyuno native "
            "module is ABI-locked to a specific Python build - this "
            f"system's Python ({sys.version.split()[0]}) does not match "
            "it, and no separate bundled LibreOffice Python executable "
            "was found to run under instead. Install a LibreOffice build "
            "that ships its own bundled Python (or one whose python3-uno "
            "package matches this system's Python version).",
            file=sys.stderr,
        )
        sys.exit(1)
    print(f"Running luah_signage in-process under system Python {sys.version.split()[0]} ...")
    from luah_signage.app import main as app_main

    app_main()


def main() -> None:
    program_dir = find_libreoffice_program_dir()
    if not program_dir:
        print(
            "Could not locate a LibreOffice installation. Set the "
            "LIBREOFFICE_PROGRAM_DIR environment variable to its 'program' "
            "directory and try again.",
            file=sys.stderr,
        )
        sys.exit(1)

    exe_name = "python.exe" if os.name == "nt" else "python"
    lo_python = Path(program_dir) / exe_name
    if not lo_python.exists():
        # No bundled interpreter binary shipped with this LibreOffice
        # install (e.g. some Linux distro packages only ship the pyuno
        # libs, compiled against the system Python) - fall back to
        # importing uno directly under this system Python instead of
        # erroring out. See _run_in_process_under_system_python's
        # docstring.
        print(
            f"No bundled LibreOffice Python found at {lo_python} - trying "
            "to import 'uno' directly under this system Python instead."
        )
        _run_in_process_under_system_python()
        return

    extra_paths = [str(PYZMANIM_LIB), str(LUAH_SIGNAGE), str(REPO_ROOT)]
    if os.name == "nt":
        tzdata_dir = _find_tzdata_dir()
        if tzdata_dir:
            extra_paths.append(tzdata_dir)
        else:
            print(
                "Warning: 'tzdata' package not found for this Python - "
                "zoneinfo timezone lookups may fail on Windows. "
                "Install it with: pip install tzdata",
                file=sys.stderr,
            )

    env = os.environ.copy()
    existing = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = os.pathsep.join(extra_paths + ([existing] if existing else []))
    # Belt-and-suspenders alongside app.py's own sys.stdout/stderr.reconfigure()
    # call: guarantees UTF-8 console output (never raising UnicodeEncodeError
    # on non-ASCII text, e.g. Hebrew shape content in debug logs) from the
    # very first line printed, regardless of the machine's console code page
    # or locale settings.
    env["PYTHONIOENCODING"] = "utf-8:backslashreplace"

    print(f"Launching luah_signage under {lo_python} ...")
    result = subprocess.run([str(lo_python), "-m", "luah_signage.app"], env=env)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
