"""
Sanity tests for luah_signage.config (pure Python, no UNO dependency).
Run with: python tests/test_config.py (also pytest-compatible)
"""
from __future__ import annotations

import json
import tempfile
from pathlib import Path

from luah_signage import config as config_module
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
        # Must point at an actually-existing file - _resolve_pptx_path falls
        # back to DEFAULT_PPTX_PATH for a path that doesn't exist (see
        # test_nonexistent_pptx_path_defaults_to_repo_root_luah_pptx below).
        real_pptx = Path(tmp) / "custom.pptx"
        real_pptx.write_bytes(b"not a real pptx, just needs to exist")
        path = _write_config(Path(tmp), pptx_path=str(real_pptx))
        config = load_config(path)
        assert config.pptx_path == real_pptx


def test_nonexistent_pptx_path_defaults_to_repo_root_luah_pptx():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(Path(tmp), pptx_path="C:/this/path/does/not/exist.pptx")
        config = load_config(path)
        assert config.pptx_path == DEFAULT_PPTX_PATH


def test_pptx_path_pointing_at_a_directory_defaults_to_repo_root_luah_pptx():
    with tempfile.TemporaryDirectory() as tmp:
        # A directory, not a file - is_file() returns False for this too.
        path = _write_config(Path(tmp), pptx_path=tmp)
        config = load_config(path)
        assert config.pptx_path == DEFAULT_PPTX_PATH


def test_debug_repaint_defaults():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(Path(tmp))
        config = load_config(path)
        assert config.debug_repaint_lock_controllers is True
        assert config.debug_repaint_skip_unchanged_clock_writes is True
        assert config.debug_repaint_update_offscreen_clocks is True
        assert config.debug_repaint_nudge_shape is False


def test_debug_skip_slideshow_default_false():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(Path(tmp))
        config = load_config(path)
        assert config.debug_skip_slideshow is False


def test_debug_skip_slideshow_override():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(Path(tmp), debug_skip_slideshow=True)
        config = load_config(path)
        assert config.debug_skip_slideshow is True


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


def test_cmd_tag_defaults():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(Path(tmp))
        config = load_config(path)
        assert config.cmd_tag_executable is None
        assert config.cmd_tag_timeout_seconds == 10.0
        assert config.cmd_tag_global_args == []
        assert config.cmd_tick_seconds == 300.0


def test_cmd_tag_overrides():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(
            Path(tmp),
            cmd_tag_executable="python foo.py",
            cmd_tag_timeout_seconds=5.0,
            cmd_tag_global_args="--shared-flag",
            cmd_tick_seconds=60.0,
        )
        config = load_config(path)
        assert config.cmd_tag_executable == "python foo.py"
        assert config.cmd_tag_timeout_seconds == 5.0
        assert config.cmd_tag_global_args == "--shared-flag"
        assert config.cmd_tick_seconds == 60.0


def test_cmd_tag_global_args_as_json_list():
    with tempfile.TemporaryDirectory() as tmp:
        path = _write_config(
            Path(tmp),
            cmd_tag_global_args=["--sep", " - ", "--swap"],
        )
        config = load_config(path)
        assert config.cmd_tag_global_args == ["--sep", " - ", "--swap"]


def test_parse_config_override_valid_json():
    assert config_module.parse_config_override('{"cmd_tick_seconds": 60}') == {"cmd_tick_seconds": 60}


def test_parse_config_override_invalid_json_raises():
    try:
        config_module.parse_config_override("{not valid json")
        assert False, "expected a JSONDecodeError"
    except json.JSONDecodeError:
        pass


def test_parse_config_override_non_dict_raises():
    try:
        config_module.parse_config_override("[1, 2, 3]")
        assert False, "expected a ValueError"
    except ValueError:
        pass


def _base_config(tmp: Path):
    path = _write_config(Path(tmp))
    return load_config(path)


def test_apply_config_overrides_single_key():
    with tempfile.TemporaryDirectory() as tmp:
        base = _base_config(tmp)
        updated = config_module.apply_config_overrides(base, {"cmd_tick_seconds": 60})
        assert updated.cmd_tick_seconds == 60
        # Everything else unchanged.
        assert updated.content_tick_seconds == base.content_tick_seconds
        assert updated.pptx_path == base.pptx_path


def test_apply_config_overrides_does_not_mutate_original():
    with tempfile.TemporaryDirectory() as tmp:
        base = _base_config(tmp)
        config_module.apply_config_overrides(base, {"cmd_tick_seconds": 60})
        assert base.cmd_tick_seconds == 300.0  # unchanged


def test_apply_config_overrides_nested_clock_style_partial_merge():
    with tempfile.TemporaryDirectory() as tmp:
        base = _base_config(tmp)
        updated = config_module.apply_config_overrides(base, {"clock_style": {"show_second_hand": False}})
        assert updated.clock_style.show_second_hand is False
        # Every other clock_style field preserved from the base config.
        assert updated.clock_style.hour_color == base.clock_style.hour_color
        assert updated.clock_style.pivot_radius_mm == base.clock_style.pivot_radius_mm


def test_apply_config_overrides_nested_location_partial_merge():
    with tempfile.TemporaryDirectory() as tmp:
        base = _base_config(tmp)
        updated = config_module.apply_config_overrides(base, {"location": {"elevation": 999}})
        assert updated.location.elevation == 999
        assert updated.location.latitude == base.location.latitude
        assert updated.location.longitude == base.location.longitude


def test_apply_config_overrides_multiple_keys_at_once():
    with tempfile.TemporaryDirectory() as tmp:
        base = _base_config(tmp)
        updated = config_module.apply_config_overrides(
            base, {"cmd_tick_seconds": 60, "content_tick_seconds": 15, "auto_recover": False}
        )
        assert updated.cmd_tick_seconds == 60
        assert updated.content_tick_seconds == 15
        assert updated.auto_recover is False


def test_apply_config_overrides_unknown_key_raises():
    with tempfile.TemporaryDirectory() as tmp:
        base = _base_config(tmp)
        try:
            config_module.apply_config_overrides(base, {"not_a_real_field": 123})
            assert False, "expected a TypeError"
        except TypeError:
            pass


def test_apply_config_overrides_pptx_path_resolved_via_fallback():
    with tempfile.TemporaryDirectory() as tmp:
        base = _base_config(tmp)
        # Nonexistent path should fall back to DEFAULT_PPTX_PATH, same as
        # load_config's own pptx_path handling.
        updated = config_module.apply_config_overrides(base, {"pptx_path": "C:/does/not/exist.pptx"})
        assert updated.pptx_path == DEFAULT_PPTX_PATH


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
