"""Small service entry points used by both CLI and web frontends."""

from __future__ import annotations

from pathlib import Path

from jtet_pstmigrate.config import AppConfig, expand_user_paths
from jtet_pstmigrate.mapping import load_mapping


def load_config(path: Path | None) -> AppConfig:
    """Load and normalize application config."""
    return expand_user_paths(AppConfig.load(path))


__all__ = ["load_config", "load_mapping"]

