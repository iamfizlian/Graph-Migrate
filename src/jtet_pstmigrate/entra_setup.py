"""Create Microsoft Entra app registrations for Graph migration access."""

from __future__ import annotations

import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

import httpx
import msal

GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
BOOTSTRAP_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
BOOTSTRAP_SCOPES = [
    "Application.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All",
    "Directory.Read.All",
]
MAIL_PERMISSIONS = ["Mail.ReadWrite"]
FULL_PERMISSIONS = ["Mail.ReadWrite", "Calendars.ReadWrite", "Contacts.ReadWrite"]
PermissionPreset = Literal["mail", "full"]
DeviceFlowCallback = Callable[[dict[str, Any]], None]
EventCallback = Callable[[str], None]


@dataclass(slots=True)
class CreatedMigrationApp:
    name: str
    tenant_id: str
    client_id: str
    client_secret: str
    secret_expires_at: str
    app_object_id: str
    service_principal_id: str
    permissions: list[str] = field(default_factory=list)


def permissions_for_preset(preset: PermissionPreset) -> list[str]:
    if preset == "mail":
        return list(MAIL_PERMISSIONS)
    return list(FULL_PERMISSIONS)


class EntraSetupClient:
    """Bootstrap and create Graph app-only credentials with admin consent."""

    def __init__(self, tenant_id: str, *, timeout_seconds: float = 60.0) -> None:
        self.tenant_id = tenant_id.strip()
        self._http = httpx.Client(
            base_url=GRAPH_BASE,
            timeout=httpx.Timeout(timeout_seconds),
            headers={"Accept": "application/json"},
        )
        self._token: str | None = None

    def close(self) -> None:
        self._http.close()

    def __enter__(self) -> EntraSetupClient:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def acquire_admin_device_token(
        self,
        *,
        device_flow_callback: DeviceFlowCallback | None = None,
        event_callback: EventCallback | None = None,
    ) -> str:
        """Sign in an admin through device code and cache the bootstrap token."""
        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        app = msal.PublicClientApplication(client_id=BOOTSTRAP_CLIENT_ID, authority=authority)
        flow = app.initiate_device_flow(scopes=BOOTSTRAP_SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Could not create device-code flow: {flow!r}")
        if device_flow_callback:
            device_flow_callback(flow)
        if event_callback:
            event_callback(
                "Open {uri} and enter code {code}".format(
                    uri=flow.get("verification_uri") or flow.get("verification_uri_complete"),
                    code=flow["user_code"],
                )
            )
        result = app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            error = result.get("error_description") or result.get("error") or str(result)
            raise RuntimeError(f"Admin device-code sign-in failed: {error}")
        self._token = result["access_token"]
        if event_callback:
            event_callback("Admin sign-in complete")
        return self._token

    def create_migration_apps(
        self,
        *,
        app_prefix: str,
        app_count: int,
        secret_lifetime_days: int,
        permission_values: list[str],
        event_callback: EventCallback | None = None,
    ) -> list[CreatedMigrationApp]:
        if app_count < 1:
            raise ValueError("app_count must be at least 1")
        if secret_lifetime_days < 1:
            raise ValueError("secret_lifetime_days must be at least 1")
        graph_sp = self.get_graph_service_principal()
        roles = self.resolve_graph_app_roles(graph_sp, permission_values)
        created: list[CreatedMigrationApp] = []
        for index in range(1, app_count + 1):
            suffix = f"-{index}" if app_count > 1 else ""
            display_name = f"{app_prefix.strip() or 'pstmigrate'}{suffix}"
            if event_callback:
                event_callback(f"Creating Entra app {display_name}")
            app = self.create_app_registration(display_name, roles)
            service_principal = self.create_service_principal(app["appId"])
            secret = self.add_client_secret(
                app["id"],
                display_name="pstmigrate setup secret",
                lifetime_days=secret_lifetime_days,
            )
            self.grant_app_roles(service_principal["id"], graph_sp["id"], roles)
            created.append(
                CreatedMigrationApp(
                    name=display_name,
                    tenant_id=self.tenant_id,
                    client_id=app["appId"],
                    client_secret=secret["secretText"],
                    secret_expires_at=secret["endDateTime"],
                    app_object_id=app["id"],
                    service_principal_id=service_principal["id"],
                    permissions=[role["value"] for role in roles],
                )
            )
            if event_callback:
                event_callback(f"Created {display_name} with {', '.join(permission_values)}")
        return created

    def get_graph_service_principal(self) -> dict[str, Any]:
        response = self._request(
            "GET",
            "/servicePrincipals",
            params={
                "$filter": f"appId eq '{GRAPH_APP_ID}'",
                "$select": "id,appId,displayName,appRoles",
            },
        )
        values = response.get("value") or []
        if not values:
            raise RuntimeError("Could not find the Microsoft Graph service principal in this tenant")
        return values[0]

    def resolve_graph_app_roles(
        self,
        graph_service_principal: dict[str, Any],
        permission_values: list[str],
    ) -> list[dict[str, Any]]:
        roles = graph_service_principal.get("appRoles") or []
        resolved: list[dict[str, Any]] = []
        for permission in permission_values:
            role = next(
                (
                    item
                    for item in roles
                    if item.get("value") == permission
                    and "Application" in (item.get("allowedMemberTypes") or [])
                ),
                None,
            )
            if role is None:
                raise RuntimeError(f"Could not find Microsoft Graph application role {permission}")
            resolved.append(role)
        return resolved

    def create_app_registration(self, display_name: str, roles: list[dict[str, Any]]) -> dict[str, Any]:
        payload = {
            "displayName": display_name,
            "signInAudience": "AzureADMyOrg",
            "requiredResourceAccess": [
                {
                    "resourceAppId": GRAPH_APP_ID,
                    "resourceAccess": [{"id": role["id"], "type": "Role"} for role in roles],
                }
            ],
        }
        return self._request("POST", "/applications", json=payload, expect_status=(201,))

    def create_service_principal(self, app_id: str) -> dict[str, Any]:
        return self._request(
            "POST",
            "/servicePrincipals",
            json={"appId": app_id},
            expect_status=(201,),
        )

    def add_client_secret(
        self,
        application_object_id: str,
        *,
        display_name: str,
        lifetime_days: int,
    ) -> dict[str, Any]:
        end_date = _utc_timestamp_after_days(lifetime_days)
        return self._request(
            "POST",
            f"/applications/{application_object_id}/addPassword",
            json={"passwordCredential": {"displayName": display_name, "endDateTime": end_date}},
        )

    def grant_app_roles(
        self,
        service_principal_id: str,
        graph_service_principal_id: str,
        roles: list[dict[str, Any]],
    ) -> None:
        for role in roles:
            payload = {
                "principalId": service_principal_id,
                "resourceId": graph_service_principal_id,
                "appRoleId": role["id"],
            }
            try:
                self._request(
                    "POST",
                    f"/servicePrincipals/{service_principal_id}/appRoleAssignments",
                    json=payload,
                    expect_status=(201,),
                )
            except RuntimeError as e:
                if "Permission being assigned already exists" not in str(e):
                    raise

    def _request(
        self,
        method: str,
        path: str,
        *,
        json: dict[str, Any] | None = None,
        params: dict[str, Any] | None = None,
        expect_status: tuple[int, ...] = (200,),
    ) -> dict[str, Any]:
        if not self._token:
            raise RuntimeError("Admin bootstrap token is not available")
        response = self._http.request(
            method,
            path,
            json=json,
            params=params,
            headers={"Authorization": f"Bearer {self._token}"},
        )
        if response.status_code not in expect_status:
            raise RuntimeError(f"Graph {response.status_code}: {_safe_body(response)}")
        if response.content:
            return response.json()
        return {}


def create_migration_apps_with_device_login(
    *,
    tenant_id: str,
    app_prefix: str,
    app_count: int,
    secret_lifetime_days: int,
    permission_preset: PermissionPreset,
    device_flow_callback: DeviceFlowCallback | None = None,
    event_callback: EventCallback | None = None,
) -> list[CreatedMigrationApp]:
    permissions = permissions_for_preset(permission_preset)
    with EntraSetupClient(tenant_id) as client:
        client.acquire_admin_device_token(
            device_flow_callback=device_flow_callback,
            event_callback=event_callback,
        )
        return client.create_migration_apps(
            app_prefix=app_prefix,
            app_count=app_count,
            secret_lifetime_days=secret_lifetime_days,
            permission_values=permissions,
            event_callback=event_callback,
        )


def write_config_for_created_apps(config_path: Path, apps: list[CreatedMigrationApp]) -> None:
    if not apps:
        raise ValueError("No created apps were provided")
    config_path.parent.mkdir(parents=True, exist_ok=True)
    app_blocks: list[str] = []
    for app in apps:
        app_blocks.append(
            "\n".join(
                [
                    "[[apps]]",
                    f'name = "{_toml_escape(app.name)}"',
                    f'tenant_id = "{_toml_escape(app.tenant_id)}"',
                    f'client_id = "{_toml_escape(app.client_id)}"',
                    f'client_secret = "{_toml_escape(app.client_secret)}"',
                ]
            )
        )
    body = f"""log_level = "INFO"

{chr(10).join(app_blocks)}

[throttle]
max_retries = 8
initial_backoff_seconds = 1.0
max_backoff_seconds = 120.0

[migration]
workers_per_mailbox = 4
max_parallel_mailboxes = 8
target_root_folder = "Imported PST"
fail_fast = false

[paths]
state_dir = ".pstmigrate-state"
work_dir = ".pstmigrate-work"
log_dir = "logs"
readpst_binary = "readpst"
"""
    config_path.write_text(body, encoding="utf-8")


def _utc_timestamp_after_days(days: int) -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(time.time() + days * 86400))


def _toml_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def _safe_body(response: httpx.Response) -> str:
    try:
        return str(response.json())
    except Exception:
        return response.text[:1000]
