"""
Polling-based file-change watcher, used to detect when the presentation
.pptx file has been updated by a cloud-sync client (OneDrive/Dropbox/Google
Drive/etc.) so the presentation can be reloaded live.

Polling (rather than OS filesystem-change notifications) is used
deliberately: cloud-sync clients often replace files via a
temp-file-then-rename dance that can be unreliable to observe with native
FS-event watchers, and polling mtime is simple and dependency-free (mirrors
the original VBA `watchdog` timer's FileDateTime-based approach).
"""
from __future__ import annotations

from pathlib import Path


class FileChangeWatcher:
    def __init__(self, path: Path):
        self.path = Path(path)
        self._last_mtime = self._read_mtime()

    def _read_mtime(self) -> float | None:
        try:
            return self.path.stat().st_mtime
        except FileNotFoundError:
            return None

    def check_for_change(self) -> bool:
        """Returns True (once) if the file's mtime has changed since the
        last check (or since construction). Updates internal state so
        subsequent calls return False until it changes again."""
        current = self._read_mtime()
        changed = current != self._last_mtime
        if changed:
            self._last_mtime = current
        return changed
