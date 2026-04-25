"""Strongly-typed configuration loaded from TOML + env vars.

Precedence (highest wins):
  1. Environment variables (PSTMIGRATE_*)
  2. TOML file passed via --config
  3. Built-in defaults
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field, field_validator, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

if sys.version_info >= (3, 11):
    import tomllib
else:  # pragma: no cover
    import tomli as tomllib


class AuthConfig(BaseModel):
    """Microsoft Entra app-only authentication for one app registration.

    Either client_secret OR client_certificate_path must be set.
    Certificate auth is strongly preferred for production.
    """

    name: str | None = Field(
        default=None,
        description="Human-readable identifier for logs/state. Defaults to client_id[:8].",
    )
    tenant_id: str = Field(description="Entra tenant GUID or domain (contoso.onmicrosoft.com)")
    client_id: str = Field(description="App registration (client) GUID")
    client_secret: str | None = Field(default=None, description="Client secret value")
    client_certificate_path: Path | None = Field(default=None, description="PEM/PFX certificate path")
    client_certificate_password: str | None = Field(default=None)

    @field_validator("client_secret")
    @classmethod
    def _strip(cls, v: str | None) -> str | None:
        return v.strip() if v else None

    @property
    def display_name(self) -> str:
        return (self.name or self.client_id[:8]).lower()

    def assert_usable(self) -> None:
        if not self.client_secret and not self.client_certificate_path:
            raise ValueError(f"auth[{self.display_name}]: set client_secret or client_certificate_path")
        if self.client_certificate_path and not self.client_certificate_path.exists():
            raise ValueError(
                f"auth[{self.display_name}]: certificate not found: {self.client_certificate_path}"
            )


class ThrottleConfig(BaseModel):
    """Graph throttling / retry behaviour.

    Defaults are deliberately conservative. Graph publishes per-app limits at
    https://learn.microsoft.com/graph/throttling-limits — adjust to match.
    """

    max_retries: int = 8
    initial_backoff_seconds: float = 1.0
    max_backoff_seconds: float = 120.0
    backoff_multiplier: float = 2.0
    backoff_jitter: float = 0.25
    request_timeout_seconds: float = 120.0


class MigrationConfig(BaseModel):
    """Per-run migration tuning."""

    workers_per_mailbox: int = Field(default=2, ge=1, le=16, description="Concurrent uploads per mailbox")
    max_parallel_mailboxes: int = Field(default=4, ge=1, le=64)
    target_root_folder: str = Field(default="Imported PST", description="Default if mapping CSV omits it")
    skip_empty_folders: bool = True
    fail_fast: bool = False
    large_attachment_threshold_bytes: int = Field(default=3 * 1024 * 1024, description="MIME size to switch to upload session")


class PathsConfig(BaseModel):
    state_dir: Path = Path(".pstmigrate-state")
    work_dir: Path = Path(".pstmigrate-work")
    log_dir: Path = Path("logs")
    readpst_binary: Path = Path("readpst")


class AppConfig(BaseSettings):
    """Top-level config. Pydantic-settings handles env var binding."""

    model_config = SettingsConfigDict(
        env_prefix="PSTMIGRATE_",
        env_nested_delimiter="__",
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # Multi-app pool. Throughput scales linearly with len(apps) up to per-mailbox caps,
    # because Graph throttling is keyed off (app_id, target_resource) — separate apps
    # get separate budgets. A single-app deployment sets apps = [<one entry>].
    apps: list[AuthConfig] = Field(default_factory=list)
    # Backward-compat shorthand: a single [auth] table is auto-promoted into apps[0].
    auth: AuthConfig | None = Field(default=None)
    throttle: ThrottleConfig = ThrottleConfig()
    migration: MigrationConfig = MigrationConfig()
    paths: PathsConfig = PathsConfig()
    log_level: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = "INFO"

    @model_validator(mode="after")
    def _normalize_apps(self) -> AppConfig:
        if self.auth is not None and not self.apps:
            self.apps = [self.auth]
        if self.auth is not None and self.apps and self.auth not in self.apps:
            # Treat duplicate definitions as a misconfiguration so users notice.
            raise ValueError("config: set either [auth] OR [[apps]], not both")
        if not self.apps:
            raise ValueError("config: at least one [[apps]] entry (or [auth]) is required")
        # Ensure each app's display_name is unique
        seen: dict[str, int] = {}
        for app in self.apps:
            base = app.display_name
            seen[base] = seen.get(base, 0) + 1
            if seen[base] > 1:
                app.name = f"{base}-{seen[base]}"
        for app in self.apps:
            app.assert_usable()
        return self

    @classmethod
    def load(cls, config_path: Path | None = None) -> AppConfig:
        """Load TOML (if provided), then layer env vars on top via BaseSettings."""
        data: dict = {}
        if config_path:
            if not config_path.exists():
                raise FileNotFoundError(f"config file not found: {config_path}")
            with open(config_path, "rb") as f:
                data = tomllib.load(f)
        cfg = cls.model_validate(data) if data else cls()  # type: ignore[call-arg]
        return cfg


def expand_user_paths(cfg: AppConfig) -> AppConfig:
    """Resolve ~ in path fields."""
    cfg.paths.state_dir = Path(os.path.expanduser(cfg.paths.state_dir)).resolve()
    cfg.paths.work_dir = Path(os.path.expanduser(cfg.paths.work_dir)).resolve()
    cfg.paths.log_dir = Path(os.path.expanduser(cfg.paths.log_dir)).resolve()
    return cfg


class MappingRow(BaseModel):
    """One row of the input mapping CSV."""

    pst_path: Path
    target_mailbox: str
    target_root_folder: str = ""

    @field_validator("target_mailbox")
    @classmethod
    def _strip_upn(cls, v: str) -> str:
        v = v.strip().lower()
        if "@" not in v:
            raise ValueError(f"target_mailbox must be a UPN (got '{v}')")
        return v

    @field_validator("pst_path")
    @classmethod
    def _resolve_pst(cls, v: Path) -> Path:
        return v.expanduser()
