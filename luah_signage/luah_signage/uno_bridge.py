"""
Connects Python to a running (or newly launched) LibreOffice instance via
UNO, over a socket. Requires a LibreOffice installation on the machine; the
`uno` module itself ships inside LibreOffice's `program` directory rather
than being pip-installable, so we locate and add that directory to
sys.path before importing it.

Cross-platform (Windows/Linux, x86/ARM): the program directory is located
via (in order) an explicit env var override, `soffice` on PATH (resolved to
its real install location - works for distro packages, Flatpak, and Snap on
any architecture since the path is architecture-independent), then a list of
common fixed install locations per OS.

Note: this module cannot be exercised in this dev environment (no
LibreOffice installed here) - it has been written carefully against the
documented UNO API but should be smoke-tested on the actual signage machine
once LibreOffice is installed there.
"""
from __future__ import annotations

import glob
import os
import shutil
import subprocess
import sys
import time

_WINDOWS_PROGRAM_DIR_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program",
    r"C:\Program Files (x86)\LibreOffice\program",
]

# Directory names are the same regardless of CPU architecture (x86_64 or
# ARM/aarch64) for these distro/package layouts - Debian/Ubuntu (and
# derivatives, e.g. Raspberry Pi OS) always install under
# /usr/lib/libreoffice regardless of arch.
_LINUX_PROGRAM_DIR_CANDIDATES = [
    "/usr/lib/libreoffice/program",
    "/usr/lib64/libreoffice/program",
    "/opt/libreoffice/program",
    *sorted(glob.glob("/opt/libreoffice*/program")),
    *sorted(glob.glob("/snap/libreoffice/current/lib/libreoffice/program")),
]


def _which_soffice_program_dir() -> str | None:
    exe = shutil.which("soffice")
    if not exe:
        return None
    real = os.path.realpath(exe)
    program_dir = os.path.dirname(real)
    if os.path.isdir(program_dir):
        return program_dir
    return None


def find_libreoffice_program_dir() -> str | None:
    env_override = os.environ.get("LIBREOFFICE_PROGRAM_DIR")
    if env_override and os.path.isdir(env_override):
        return env_override

    from_path = _which_soffice_program_dir()
    if from_path:
        return from_path

    candidates = _WINDOWS_PROGRAM_DIR_CANDIDATES if os.name == "nt" else _LINUX_PROGRAM_DIR_CANDIDATES
    for candidate in candidates:
        if os.path.isdir(candidate):
            return candidate
    return None


def ensure_uno_importable() -> None:
    """Makes `import uno` work by adding LibreOffice's program directory to
    sys.path if it isn't already importable (e.g. when running under a
    normal system Python rather than LibreOffice's bundled Python)."""
    try:
        import uno  # noqa: F401
        return
    except ImportError:
        pass

    program_dir = find_libreoffice_program_dir()
    if not program_dir:
        raise RuntimeError(
            "Could not locate a LibreOffice installation. Install "
            "LibreOffice, or set the LIBREOFFICE_PROGRAM_DIR environment "
            "variable to its 'program' directory."
        )
    sys.path.insert(0, program_dir)
    import uno  # noqa: F401


def launch_soffice(host: str = "localhost", port: int = 2002, headless: bool = False) -> subprocess.Popen:
    """Launches soffice with a UNO socket listener. `headless=False` (the
    default) still lets the slideshow itself display on screen (required
    for digital signage) - but the Start Center ("welcome") window is
    skipped, and the underlying Impress document/editing window is loaded
    hidden (see presentation.py's Hidden=True); only the fullscreen
    slideshow window that PresentationController explicitly starts should
    ever become visible.

    --nodefault: skip opening the Start Center/blank document when soffice
    starts with no document specified on the command line (we always load
    our document ourselves afterward via UNO, once connected) - without
    this flag, the Start Center window flashes up briefly/lingers.

    --norestore: also suppresses the "Document Recovery" dialog/wizard
    that would otherwise appear on the NEXT soffice startup after an
    unclean shutdown (crash, power outage, task-killed process, etc.) -
    important for unattended signage hardware where nobody is present to
    click through it.

    SAL_DISABLE_CRASHDUMP=1 (set as an environment variable for the
    soffice process, not a command-line flag): disables LibreOffice's own
    crash reporter entirely, so a crash never spawns its "send crash
    report?" dialog either - also important for unattended hardware after
    a non-graceful exit (e.g. a power outage killing soffice.bin
    mid-operation). This is a well-known env var LibreOffice itself
    checks for exactly this automation/unattended use case.

    NOTE: `--invisible` was tried here too (to suppress ALL windows at
    startup) but had to be reverted - it suppresses windows so completely
    that Impress's own slideshow view could not be created either, and
    `presentation.start()` crashed with `DisposedException: Binary URP
    bridge disposed during call` (the whole soffice process appears to
    terminate/become unusable once a real window is asked for while
    running under --invisible). Do not re-add --invisible."""
    program_dir = find_libreoffice_program_dir()
    if not program_dir:
        raise RuntimeError("LibreOffice installation not found")
    exe_name = "soffice.exe" if os.name == "nt" else "soffice"
    soffice_exe = os.path.join(program_dir, exe_name)

    args = [
        soffice_exe,
        "--norestore",
        "--nologo",
        "--nofirststartwizard",
        "--nodefault",
        f"--accept=socket,host={host},port={port};urp;",
    ]
    if headless:
        args.append("--headless")

    env = os.environ.copy()
    env["SAL_DISABLE_CRASHDUMP"] = "1"
    return subprocess.Popen(args, env=env)


def connect(host: str = "localhost", port: int = 2002, retries: int = 40, delay: float = 0.5):
    """Connects to a running soffice instance's UNO socket, retrying while
    it starts up. Returns the remote component context."""
    ensure_uno_importable()
    import uno
    from com.sun.star.connection import NoConnectException

    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    url = f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"

    last_exc: Exception | None = None
    for _ in range(retries):
        try:
            return resolver.resolve(url)
        except NoConnectException as exc:
            last_exc = exc
            time.sleep(delay)
    raise RuntimeError(f"Could not connect to LibreOffice on {host}:{port}") from last_exc


def connect_or_launch(host: str = "localhost", port: int = 2002, headless: bool = False):
    """Tries to connect to an already-running soffice first; if that fails
    immediately, launches a new instance and connects to it."""
    ensure_uno_importable()
    import uno
    from com.sun.star.connection import NoConnectException

    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    url = f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"
    try:
        return resolver.resolve(url)
    except NoConnectException:
        launch_soffice(host=host, port=port, headless=headless)
        return connect(host=host, port=port)
