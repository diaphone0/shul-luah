"""
Sanity tests for luah_signage.watcher (pure Python, no UNO dependency).
Run with: python tests/test_watcher.py (also pytest-compatible)
"""
from __future__ import annotations

import os
import tempfile
import time
from pathlib import Path

from luah_signage.watcher import FileChangeWatcher


def test_no_change_initially():
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "luah.pptx"
        path.write_text("v1")
        watcher = FileChangeWatcher(path)
        assert watcher.check_for_change() is False


def test_detects_mtime_change():
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "luah.pptx"
        path.write_text("v1")
        watcher = FileChangeWatcher(path)
        assert watcher.check_for_change() is False

        # Bump mtime forward to simulate a cloud-sync rewrite.
        new_time = time.time() + 5
        os.utime(path, (new_time, new_time))

        assert watcher.check_for_change() is True
        # Subsequent checks are False until it changes again.
        assert watcher.check_for_change() is False


def test_handles_missing_file_gracefully():
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "does_not_exist.pptx"
        watcher = FileChangeWatcher(path)
        assert watcher.check_for_change() is False

        path.write_text("now it exists")
        assert watcher.check_for_change() is True


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
