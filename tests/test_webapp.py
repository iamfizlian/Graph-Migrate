from pathlib import Path

import httpx
import pytest

pytest.importorskip("fastapi")

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
