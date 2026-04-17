"""Shared utility helpers."""

from __future__ import annotations

from pathlib import Path


def ensure_parent_directory(path: Path) -> None:
    """Create the parent directory for the provided path if needed."""

    path.parent.mkdir(parents=True, exist_ok=True)
