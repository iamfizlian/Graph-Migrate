"""Loguru configuration: pretty console + structured JSON file sink."""

from __future__ import annotations

import sys
from pathlib import Path

from loguru import logger


def configure_logging(log_dir: Path, level: str = "INFO", run_id: str = "run") -> Path:
    """Set up dual sinks. Returns the JSON log path for later reference."""
    log_dir.mkdir(parents=True, exist_ok=True)
    json_path = log_dir / f"{run_id}.jsonl"

    logger.remove()
    logger.add(
        sys.stderr,
        level=level,
        format=(
            "<dim>{time:HH:mm:ss}</dim> <level>{level:<7}</level> "
            "<cyan>{extra[ctx]}</cyan> {message}"
        ),
        filter=_inject_default_ctx,
        colorize=True,
    )
    logger.add(
        json_path,
        level="DEBUG",
        serialize=True,
        rotation="100 MB",
        retention=10,
        compression="zip",
        enqueue=True,
    )
    return json_path


def _inject_default_ctx(record: dict) -> bool:
    record["extra"].setdefault("ctx", "-")
    return True


def bind(**kwargs):
    """Convenience: `bind(ctx='mailbox=foo')` returns a contextual logger."""
    return logger.bind(**kwargs)
