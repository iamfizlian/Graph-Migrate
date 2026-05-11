from pathlib import Path

from jtet_pstmigrate.entra_setup import (
    CreatedMigrationApp,
    permissions_for_preset,
    write_config_for_created_apps,
)
from jtet_pstmigrate.services import load_config


def test_permissions_for_preset() -> None:
    assert permissions_for_preset("mail") == ["Mail.ReadWrite"]
    assert permissions_for_preset("full") == ["Mail.ReadWrite", "Calendars.ReadWrite", "Contacts.ReadWrite"]


def test_write_config_for_created_apps(tmp_path: Path) -> None:
    config = tmp_path / "config.toml"
    write_config_for_created_apps(
        config,
        [
            CreatedMigrationApp(
                name="pstmigrate-1",
                tenant_id="contoso.onmicrosoft.com",
                client_id="00000000-0000-0000-0000-000000000001",
                client_secret='secret"one',
                secret_expires_at="2026-12-01T00:00:00Z",
                app_object_id="app-object-1",
                service_principal_id="sp-object-1",
                permissions=["Mail.ReadWrite"],
            ),
            CreatedMigrationApp(
                name="pstmigrate-2",
                tenant_id="contoso.onmicrosoft.com",
                client_id="00000000-0000-0000-0000-000000000002",
                client_secret="secret-two",
                secret_expires_at="2026-12-01T00:00:00Z",
                app_object_id="app-object-2",
                service_principal_id="sp-object-2",
                permissions=["Mail.ReadWrite"],
            ),
        ],
    )

    text = config.read_text(encoding="utf-8")
    cfg = load_config(config)

    assert text.count("[[apps]]") == 2
    assert 'client_secret = "secret\\"one"' in text
    assert [app.name for app in cfg.apps] == ["pstmigrate-1", "pstmigrate-2"]
    assert cfg.apps[1].client_secret == "secret-two"

