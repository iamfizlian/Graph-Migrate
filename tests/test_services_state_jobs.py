from pathlib import Path

from jtet_pstmigrate.jobs import JobSpec, run_job
from jtet_pstmigrate.orchestrator import RunReport
from jtet_pstmigrate.reports import StateQueries
from jtet_pstmigrate.services import load_config
from jtet_pstmigrate.state import StateStore


def test_load_config_from_toml(tmp_path: Path) -> None:
    config = _config_file(tmp_path)

    cfg = load_config(config)

    assert cfg.apps[0].tenant_id == "contoso.onmicrosoft.com"
    assert cfg.paths.state_dir == (tmp_path / "state").resolve()


def test_state_queries_dashboard_totals(tmp_path: Path) -> None:
    state = StateStore(tmp_path / "state.sqlite")
    state.start_pst_run("/pst/a.pst", "alice@example.com")
    state.update_pst_run("/pst/a.pst", "alice@example.com", status="uploading", items_total=2)
    state.upsert_message(
        mailbox="alice@example.com",
        pst_path="/pst/a.pst",
        source_path="/work/a/1.eml",
        dedupe_key="imid:1",
        status="done",
    )
    state.upsert_message(
        mailbox="alice@example.com",
        pst_path="/pst/a.pst",
        source_path="/work/a/2.eml",
        dedupe_key="imid:2",
        status="failed",
    )

    totals = StateQueries(state).dashboard_totals()

    assert totals.runs == 1
    assert totals.done == 1
    assert totals.failed == 1


def test_run_job_passes_duplicate_import_mode_to_orchestrator(tmp_path: Path, monkeypatch) -> None:
    captured: dict[str, bool] = {}

    class FakePool:
        def __init__(self, apps) -> None:
            self.apps = apps

    class FakeOrchestrator:
        def __init__(self, cfg, state, pool, *, import_skipped_duplicates: bool = False) -> None:
            captured["import_skipped_duplicates"] = import_skipped_duplicates

        def run(self, rows):
            return [RunReport(pst_path=rows[0].pst_path, mailbox=rows[0].target_mailbox, status="done")]

    monkeypatch.setattr("jtet_pstmigrate.jobs.AppPool", FakePool)
    monkeypatch.setattr("jtet_pstmigrate.jobs.Orchestrator", FakeOrchestrator)

    record = run_job(
        JobSpec(
            kind="import-mail",
            config_path=_config_file(tmp_path),
            mapping_path=_mapping_file(tmp_path),
            import_skipped_duplicates=True,
        )
    )

    assert record.status == "done"
    assert captured["import_skipped_duplicates"] is True


def test_run_job_blocks_destructive_without_confirmation(tmp_path: Path) -> None:
    record = run_job(
        JobSpec(
            kind="reset-state",
            config_path=_config_file(tmp_path),
            mapping_path=None,
            confirmed=False,
        )
    )

    assert record.status == "blocked"
    assert "Confirmation is required" in (record.last_error or "")


def _config_file(tmp_path: Path) -> Path:
    config = tmp_path / "config.toml"
    config.write_text(
        f"""
[[apps]]
name = "primary"
tenant_id = "contoso.onmicrosoft.com"
client_id = "00000000-0000-0000-0000-000000000000"
client_secret = "secret"

[paths]
state_dir = "{tmp_path / "state"}"
work_dir = "{tmp_path / "work"}"
log_dir = "{tmp_path / "logs"}"
readpst_binary = "readpst"
""",
        encoding="utf-8",
    )
    return config


def _mapping_file(tmp_path: Path) -> Path:
    pst = tmp_path / "a.pst"
    pst.write_bytes(b"")
    mapping = tmp_path / "mapping.csv"
    mapping.write_text(
        f"PSTPath,TargetMailbox,TargetRootFolder\n{pst},alice@example.com,Imported\n",
        encoding="utf-8",
    )
    return mapping
