import json
from pathlib import Path

import httpx
import pytest

pytest.importorskip("fastapi")

from jtet_pstmigrate.jobs import JobRecord
from jtet_pstmigrate.webapp import create_app


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.mark.anyio
async def test_web_dashboard_loads(tmp_path: Path) -> None:
    transport = httpx.ASGITransport(
        app=create_app(config_path=_config_file(tmp_path), mapping_path=_mapping_file(tmp_path))
    )

    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.get("/")

    assert response.status_code == 200
    assert "Dashboard" in response.text


@pytest.mark.anyio
async def test_web_destructive_job_requires_confirmation(tmp_path: Path) -> None:
    transport = httpx.ASGITransport(
        app=create_app(config_path=_config_file(tmp_path), mapping_path=_mapping_file(tmp_path))
    )

    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.post("/jobs", data={"kind": "reset-state"}, follow_redirects=True)

    assert response.status_code == 200
    assert "blocked" in response.text
    assert "Confirmation is required" in response.text


@pytest.mark.anyio
async def test_run_page_can_submit_import_skipped_duplicates(tmp_path: Path) -> None:
    app = create_app(config_path=_config_file(tmp_path), mapping_path=_mapping_file(tmp_path))

    class CapturingJobs:
        def __init__(self) -> None:
            self.spec = None

        def latest(self):
            return []

        def submit(self, spec):
            self.spec = spec
            return JobRecord(job_id="job123", kind=spec.kind)

        def get(self, job_id: str):
            return JobRecord(job_id=job_id, kind="import-mail")

    jobs = CapturingJobs()
    app.state.jobs = jobs
    transport = httpx.ASGITransport(app=app)

    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        run_page = await client.get("/run")
        response = await client.post(
            "/jobs",
            data={
                "kind": "import-mail",
                "mailbox": "alice@example.com",
                "pst": "a.pst",
                "import_skipped_duplicates": "yes",
            },
            follow_redirects=False,
        )

    assert run_page.status_code == 200
    assert 'name="import_skipped_duplicates"' in run_page.text
    assert response.status_code == 303
    assert jobs.spec is not None
    assert jobs.spec.kind == "import-mail"
    assert jobs.spec.import_skipped_duplicates is True
    assert jobs.spec.filters.mailboxes == ["alice@example.com"]
    assert jobs.spec.filters.pst_names == ["a.pst"]


@pytest.mark.anyio
async def test_web_config_page_can_create_config(tmp_path: Path) -> None:
    config = tmp_path / "config.toml"
    transport = httpx.ASGITransport(app=create_app(config_path=config, mapping_path=_mapping_file(tmp_path)))

    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        initial = await client.get("/config")
        response = await client.post(
            "/config",
            data={
                "app_name": "primary",
                "tenant_id": "contoso.onmicrosoft.com",
                "client_id": "00000000-0000-0000-0000-000000000000",
                "client_secret": "new-secret",
                "client_certificate_path": "",
                "workers_per_mailbox": "3",
                "max_parallel_mailboxes": "5",
                "state_dir": str(tmp_path / "state"),
                "work_dir": str(tmp_path / "work"),
                "log_dir": str(tmp_path / "logs"),
                "readpst_binary": "readpst",
            },
            follow_redirects=True,
        )

    assert initial.status_code == 200
    assert "Config is not ready" in initial.text
    assert response.status_code == 200
    assert config.exists()
    assert 'client_secret = "new-secret"' in config.read_text(encoding="utf-8")
    assert "new-secret" not in response.text
    assert "********" in response.text


@pytest.mark.anyio
async def test_mapping_aware_pages_render_select_options(tmp_path: Path) -> None:
    transport = httpx.ASGITransport(
        app=create_app(config_path=_config_file(tmp_path), mapping_path=_mapping_file(tmp_path))
    )

    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        validate = await client.get("/validate")
        run = await client.get("/run")
        tools = await client.get("/tools")

    for response in (validate, run, tools):
        assert response.status_code == 200
        assert '<select name="mailbox">' in response.text
        assert '<select name="pst">' in response.text
        assert "alice@example.com" in response.text
        assert "a.pst" in response.text


@pytest.mark.anyio
async def test_log_detail_page_renders_log_file(tmp_path: Path) -> None:
    config = _config_file(tmp_path)
    log_dir = tmp_path / "logs"
    log_dir.mkdir(exist_ok=True)
    (log_dir / "run.jsonl").write_text(
        json.dumps(
            {
                "text": "2026-01-01 00:00:00 | INFO | hello\n",
                "record": {
                    "time": {"repr": "2026-01-01T00:00:00+00:00"},
                    "level": {"name": "INFO"},
                    "extra": {"ctx": "validate"},
                    "message": "hello",
                },
            }
        )
        + "\n"
        + json.dumps(
            {
                "record": {
                    "time": {"repr": "2026-01-01T00:00:01+00:00"},
                    "level": {"name": "ERROR"},
                    "extra": {"ctx": "graph"},
                    "message": "bad thing",
                    "exception": {"type": "RuntimeError", "value": "boom"},
                }
            }
        )
        + "\n",
        encoding="utf-8",
    )
    transport = httpx.ASGITransport(app=create_app(config_path=config, mapping_path=_mapping_file(tmp_path)))

    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.get("/logs/run.jsonl")

    assert response.status_code == 200
    assert "run.jsonl" in response.text
    assert "hello" in response.text
    assert "bad thing" in response.text
    assert "validate" in response.text
    assert "Errors" in response.text
    assert "<table>" in response.text


def _mapping_file(tmp_path: Path) -> Path:
    pst = tmp_path / "a.pst"
    pst.write_bytes(b"")
    mapping = tmp_path / "mapping.csv"
    mapping.write_text(
        f"PSTPath,TargetMailbox,TargetRootFolder\n{pst},alice@example.com,Imported\n",
        encoding="utf-8",
    )
    return mapping


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
