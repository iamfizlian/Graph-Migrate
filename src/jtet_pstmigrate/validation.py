"""Pre-flight validation as structured data."""

from __future__ import annotations

from dataclasses import dataclass, field

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import AppConfig, MappingRow
from jtet_pstmigrate.graph_client import GraphClient
from jtet_pstmigrate.pst_reader import check_readpst


@dataclass(slots=True)
class ValidationCheck:
    name: str
    ok: bool
    message: str


@dataclass(slots=True)
class ValidationReport:
    checks: list[ValidationCheck] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return all(check.ok for check in self.checks)


def validate_environment(cfg: AppConfig, rows: list[MappingRow]) -> ValidationReport:
    """Run the same pre-flight checks as the CLI, without presentation concerns."""
    report = ValidationReport()

    try:
        version = check_readpst(cfg.paths.readpst_binary)
        report.checks.append(ValidationCheck("readpst available", True, version))
    except Exception as e:
        report.checks.append(ValidationCheck("readpst available", False, str(e)))

    missing = [r.pst_path for r in rows if not r.pst_path.exists()]
    if missing:
        report.checks.append(
            ValidationCheck("PST files exist", False, f"missing: {', '.join(str(p) for p in missing[:5])}")
        )
    else:
        total_gb = sum(r.pst_path.stat().st_size for r in rows) / 1024**3 if rows else 0.0
        report.checks.append(
            ValidationCheck("PST files exist", True, f"{len(rows)} files, {total_gb:0.2f} GB total")
        )

    try:
        pool = AppPool(cfg.apps)
        per_app_ok: list[str] = []
        per_app_fail: list[str] = []
        with GraphClient(pool, cfg.throttle) as graph:
            for name in pool.names:
                try:
                    graph.get("/$metadata", expect_status=(200,), app_id=name)
                    per_app_ok.append(name)
                except Exception as e:
                    per_app_fail.append(f"{name}: {e}")
        if per_app_fail:
            report.checks.append(
                ValidationCheck("Graph token (per app)", False, "; ".join(per_app_fail[:3]))
            )
        else:
            report.checks.append(
                ValidationCheck(
                    "Graph token (per app)",
                    True,
                    f"{len(per_app_ok)} app(s): {', '.join(per_app_ok)}",
                )
            )

        unique = sorted({r.target_mailbox for r in rows})
        unresolved: list[str] = []
        with GraphClient(pool, cfg.throttle) as graph:
            for upn in unique:
                try:
                    graph.get(
                        f"/users/{upn}/mailFolders/inbox",
                        params={"$select": "id,displayName"},
                        app_id=pool.names[0],
                    )
                except Exception as e:
                    unresolved.append(f"{upn} ({e})")
        if unresolved:
            report.checks.append(
                ValidationCheck(
                    "Mailboxes accessible",
                    False,
                    f"{len(unresolved)}/{len(unique)}: {unresolved[0]}",
                )
            )
        elif unique:
            report.checks.append(
                ValidationCheck("Mailboxes accessible", True, f"{len(unique)} mailboxes")
            )
    except Exception as e:
        report.checks.append(ValidationCheck("Graph token + connectivity", False, str(e)))

    return report

