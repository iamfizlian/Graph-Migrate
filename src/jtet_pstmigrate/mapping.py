"""Mapping CSV loading helpers shared by CLI and web frontends."""

from __future__ import annotations

import csv
from pathlib import Path

from jtet_pstmigrate.config import MappingRow


def load_mapping(csv_path: Path) -> list[MappingRow]:
    """Load a PST-to-mailbox mapping CSV."""
    rows: list[MappingRow] = []
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        required = {"PSTPath", "TargetMailbox"}
        if not required.issubset({c.strip() for c in (reader.fieldnames or [])}):
            raise ValueError(f"CSV missing required columns: {required}. Got: {reader.fieldnames}")
        for raw in reader:
            rows.append(
                MappingRow(
                    pst_path=Path(raw["PSTPath"].strip()),
                    target_mailbox=raw["TargetMailbox"].strip(),
                    target_root_folder=(raw.get("TargetRootFolder") or "").strip(),
                )
            )
    return rows

