"""
Sanity tests for luah_signage.tzdata_support (pure Python, no UNO
dependency). Run with: python tests/test_tzdata_support.py (also
pytest-compatible)
"""
from __future__ import annotations

from datetime import datetime

import zoneinfo

from luah_signage.tzdata_support import _VENDORED_ZONEINFO_DIR, ensure_tzdata_available


def test_vendored_directory_exists_and_has_files():
    assert _VENDORED_ZONEINFO_DIR.is_dir()
    assert (_VENDORED_ZONEINFO_DIR / "Asia" / "Jerusalem").exists()


def test_ensure_tzdata_available_allows_zoneinfo_lookup_with_empty_tzpath():
    # Simulate a machine with no system tz database at all (matches the
    # reported real-world failure: LibreOffice's bundled Python on a
    # portable install, with no tzdata package and no OS tz database).
    original_tzpath = zoneinfo.TZPATH
    try:
        zoneinfo.reset_tzpath(to=[])
        # Force re-registration even though the module may have already
        # registered it once (e.g. via an earlier test/import) - directly
        # call reset_tzpath again through ensure_tzdata_available's own
        # logic by resetting its internal guard.
        import luah_signage.tzdata_support as mod
        mod._registered = False
        ensure_tzdata_available()
        tz = zoneinfo.ZoneInfo("Asia/Jerusalem")
        offset = datetime(2024, 6, 21, tzinfo=tz).utcoffset()
        assert offset is not None
        assert offset.total_seconds() == 3 * 3600
    finally:
        zoneinfo.reset_tzpath(to=list(original_tzpath))


def test_ensure_tzdata_available_is_idempotent():
    ensure_tzdata_available()
    ensure_tzdata_available()  # must not raise or duplicate entries endlessly
    tz = zoneinfo.ZoneInfo("Asia/Jerusalem")
    assert tz is not None


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
