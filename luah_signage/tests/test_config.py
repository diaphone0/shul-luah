"""
Sanity tests for luah_signage.config (pure Python, no UNO dependency).
Run with: python tests/test_config.py (also pytest-compatible)
"""
from __future__ import annotations

import json
import tempfile
from pathlib import Path

from luah_signage.config import DEFAULT_PPTX_PATH, load_config


def _write_config(tmp_dir: Path, **overrides) -> Path:
    data = {
        "location": {"latitude": 32.06, "longitude": 35.26, "elevation": 720},
    }
    data.update(overrides)
    path = tmp_dir / "config.json"
    path.write_text(json.dumps(data), encoding="utf-8")
    return path


def test_missing_pptx_path_defaults_to_repo_root_luah_pptx():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(Path(tmp))
        config = load_config(path)
        assert config.pptx_path == DEFAULT_PPTX_PATH


def test_empty_string_pptx_path_defaults_to_repo_root_luah_pptx():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(Path(tmp), pptx_path="")
        config = load_config(path)
        assert config.pptx_path == DEFAULT_PPTX_PATH


def test_explicit_pptx_path_is_respected():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(Path(tmp), pptx_path="C:/somewhere/custom.pptx")
        config = load_config(path)
        assert config.pptx_path == Path("C:/somewhere/custom.pptx")


def test_debug_repaint_defaults():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(Path(tmp))
        config = load_config(path)
        assert config.debug_repaint_lock_controllers is True
        assert config.debug_repaint_skip_unchanged_clock_writes is True
        assert config.debug_repaint_update_offscreen_clocks is True
        assert config.debug_repaint_nudge_shape is False


def test_debug_repaint_overrides():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(
            Path(tmp),
            debug_repaint_lock_controllers=False,
            debug_repaint_skip_unchanged_clock_writes=False,
            debug_repaint_update_offscreen_clocks=False,
            debug_repaint_nudge_shape=True,
        )
        config = load_config(path)
        assert config.debug_repaint_lock_controllers is False
        assert config.debug_repaint_skip_unchanged_clock_writes is False
        assert config.debug_repaint_update_offscreen_clocks is False
        assert config.debug_repaint_nudge_shape is True


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
