"""Mapping selection/filtering shared by CLI, jobs, and the web UI."""

from __future__ import annotations

from dataclasses import dataclass, field

from jtet_pstmigrate.config import MappingRow


@dataclass(slots=True)
class SelectionFilters:
    mailboxes: list[str] | None = None
    pst_names: list[str] | None = None
    exclude_mailboxes: list[str] | None = None
    exclude_pst_names: list[str] | None = None
    limit: int | None = None


@dataclass(slots=True)
class SelectionResult:
    rows: list[MappingRow]
    total_rows: int
    warnings: list[str] = field(default_factory=list)

    @property
    def filtered(self) -> bool:
        return len(self.rows) != self.total_rows


def select_mapping(rows: list[MappingRow], filters: SelectionFilters) -> SelectionResult:
    """Apply CLI-compatible include/exclude filters to mapping rows."""
    selected = list(rows)
    warnings: list[str] = []

    if filters.mailboxes:
        wanted = {m.strip().lower() for m in filters.mailboxes if m and m.strip()}
        selected = [r for r in selected if r.target_mailbox.lower() in wanted]
        unmatched = wanted - {r.target_mailbox.lower() for r in selected}
        if unmatched:
            warnings.append(f"no mapping rows for: {', '.join(sorted(unmatched))}")

    if filters.pst_names:
        needles = [p.strip().lower() for p in filters.pst_names if p and p.strip()]
        selected = [r for r in selected if any(n in r.pst_path.name.lower() for n in needles)]

    if filters.exclude_mailboxes:
        unwanted = {
            m.strip().lower() for m in filters.exclude_mailboxes if m and m.strip()
        }
        before = {r.target_mailbox.lower() for r in selected}
        selected = [r for r in selected if r.target_mailbox.lower() not in unwanted]
        no_op_excludes = unwanted - before
        if no_op_excludes:
            warnings.append(
                "--exclude-mailbox had no effect for: "
                f"{', '.join(sorted(no_op_excludes))}"
            )

    if filters.exclude_pst_names:
        needles = [p.strip().lower() for p in filters.exclude_pst_names if p and p.strip()]
        selected = [
            r for r in selected if not any(n in r.pst_path.name.lower() for n in needles)
        ]

    if filters.limit is not None and filters.limit > 0:
        selected = selected[: filters.limit]

    return SelectionResult(rows=selected, total_rows=len(rows), warnings=warnings)

