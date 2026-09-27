"""
Ensures `zoneinfo.ZoneInfo(...)` lookups work even when NEITHER the OS
provides an IANA tz database NOR the `tzdata` PyPI package is installed
for the Python interpreter actually running the app.

Why this is needed: LibreOffice's bundled Python (used to run luah_signage
via launch.py - see that script's own docstring for why) is a fully
separate, isolated Python installation from the system Python, and does
NOT ship the IANA timezone database itself, nor the `tzdata` PyPI package.
`launch.py` tries to work around this by locating an already-`pip
install`ed `tzdata` package under the SYSTEM Python (the one running
launch.py itself) and forwarding its directory onto PYTHONPATH for
LibreOffice's Python to import too - but this still depends on `tzdata`
having been pip-installed for the system Python in the first place, which
may not be true on a fresh machine, and is especially likely to fail on
portable/locked-down LibreOffice installations (e.g. run from a USB/
network drive with no admin rights to install anything) - exactly the
scenario that caused a `zoneinfo.ZoneInfoNotFoundError` on a second PC.

To make timezone resolution work with ZERO external dependencies on any
machine, a full copy of the IANA tzdata database (as vendored by the
`tzdata` PyPI package, ~500KB across 648 files) is bundled directly in
this package under `_vendor_tzdata/zoneinfo/` - the same way pyzmanim and
luah_signage themselves are used as plain source trees without requiring
installation. `ensure_tzdata_available()` registers this vendored
directory as a fallback search path for the `zoneinfo` module (appended
AFTER any real system/pip-installed tz data, which is always preferred if
present and presumably kept more up to date) - call it once, before the
first `ZoneInfo(...)` construction; it's idempotent and safe to call
multiple times.
"""
from __future__ import annotations

import zoneinfo
from pathlib import Path

_VENDORED_ZONEINFO_DIR = Path(__file__).resolve().parent / "_vendor_tzdata" / "zoneinfo"

_registered = False


def ensure_tzdata_available() -> None:
    """Idempotently appends the vendored IANA tzdata directory to
    `zoneinfo`'s search path (`zoneinfo.TZPATH`), so `ZoneInfo(...)`
    lookups succeed even with no system tz database and no `tzdata`
    package installed for the running Python interpreter. Safe to call
    more than once (a no-op after the first successful call)."""
    global _registered
    if _registered:
        return
    if _VENDORED_ZONEINFO_DIR.is_dir():
        zoneinfo.reset_tzpath(to=list(zoneinfo.TZPATH) + [str(_VENDORED_ZONEINFO_DIR)])
    _registered = True
