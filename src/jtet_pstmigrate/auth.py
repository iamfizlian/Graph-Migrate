"""Microsoft Entra app-only authentication via MSAL.

Uses ConfidentialClientApplication with either client secret or X.509 cert.
Tokens are cached in-process; MSAL handles refresh transparently.

For multi-app sharding (the only legitimate way to scale Graph throughput past
the per-app throttle cap), wrap N TokenProviders in an AppPool and call pick().
"""

from __future__ import annotations

import threading
import time
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

import msal
from loguru import logger

from jtet_pstmigrate.config import AuthConfig

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


@dataclass(slots=True)
class AccessToken:
    value: str
    expires_at_epoch: float

    def is_expiring(self, skew_seconds: int = 120) -> bool:
        return time.time() + skew_seconds >= self.expires_at_epoch


class TokenProvider:
    """Thread-safe wrapper around MSAL's confidential client.

    Call `.get()` whenever you need a token; this class refreshes lazily.
    """

    def __init__(self, cfg: AuthConfig):
        cfg.assert_usable()
        self._cfg = cfg
        self._lock = threading.Lock()
        self._token: AccessToken | None = None
        self._app = self._build_app(cfg)

    @property
    def name(self) -> str:
        return self._cfg.display_name

    @property
    def client_id(self) -> str:
        return self._cfg.client_id

    @staticmethod
    def _build_app(cfg: AuthConfig) -> msal.ConfidentialClientApplication:
        authority = f"https://login.microsoftonline.com/{cfg.tenant_id}"
        if cfg.client_certificate_path:
            cert_payload = _load_certificate(
                cfg.client_certificate_path, cfg.client_certificate_password
            )
            return msal.ConfidentialClientApplication(
                client_id=cfg.client_id,
                authority=authority,
                client_credential=cert_payload,
            )
        return msal.ConfidentialClientApplication(
            client_id=cfg.client_id,
            authority=authority,
            client_credential=cfg.client_secret,
        )

    def get(self) -> str:
        with self._lock:
            if self._token and not self._token.is_expiring():
                return self._token.value
            result = self._app.acquire_token_for_client(scopes=GRAPH_SCOPE)
            if "access_token" not in result:
                err = result.get("error_description") or result.get("error") or str(result)
                raise RuntimeError(f"MSAL token acquisition failed: {err}")
            expires_at = time.time() + int(result.get("expires_in", 3600))
            self._token = AccessToken(value=result["access_token"], expires_at_epoch=expires_at)
            logger.bind(ctx="auth").debug(
                "Acquired Graph token (expires in {}s)", int(expires_at - time.time())
            )
            return self._token.value


def _load_certificate(path: Path, password: str | None) -> dict:
    """Build the dict MSAL expects for cert auth.

    Supports PEM (concatenated cert + key) and PFX/P12.
    """
    suffix = path.suffix.lower()
    if suffix in {".pfx", ".p12"}:
        return _load_pfx(path, password)
    return _load_pem(path, password)


def _load_pem(path: Path, password: str | None) -> dict:
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.x509 import load_pem_x509_certificate

    raw = path.read_bytes()
    cert = load_pem_x509_certificate(raw)
    pwd_bytes = password.encode() if password else None
    key = serialization.load_pem_private_key(raw, password=pwd_bytes)
    thumbprint = cert.fingerprint(hashes.SHA1()).hex().upper()
    private_key_pem = key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    ).decode()
    return {
        "private_key": private_key_pem,
        "thumbprint": thumbprint,
        "public_certificate": cert.public_bytes(serialization.Encoding.PEM).decode(),
    }


class AppPool:
    """Round-robin pool of TokenProviders.

    Why this exists: Graph throttle buckets are computed per (app_id, target).
    With N independent app registrations all granted Mail.ReadWrite, your
    aggregate per-mailbox throughput multiplies by ~N (until the per-mailbox
    bucket caps you, which Microsoft will lift on request for migrations).
    """

    def __init__(self, configs: Iterable[AuthConfig]):
        configs = list(configs)
        if not configs:
            raise ValueError("AppPool requires at least one AuthConfig")
        self._providers: dict[str, TokenProvider] = {}
        for cfg in configs:
            tp = TokenProvider(cfg)
            if tp.name in self._providers:
                raise ValueError(f"AppPool: duplicate app name '{tp.name}'")
            self._providers[tp.name] = tp
        self._names: list[str] = list(self._providers.keys())
        self._cursor = 0
        self._lock = threading.Lock()

    def __len__(self) -> int:
        return len(self._providers)

    @property
    def names(self) -> list[str]:
        return list(self._names)

    def providers(self) -> list[TokenProvider]:
        return list(self._providers.values())

    def get(self, name: str) -> TokenProvider:
        return self._providers[name]

    def pick(self) -> str:
        """Round-robin. Returns the chosen app's display name."""
        with self._lock:
            name = self._names[self._cursor % len(self._names)]
            self._cursor += 1
            return name


def _load_pfx(path: Path, password: str | None) -> dict:
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.serialization import pkcs12

    pwd_bytes = password.encode() if password else None
    key, cert, _extras = pkcs12.load_key_and_certificates(path.read_bytes(), pwd_bytes)
    if cert is None or key is None:
        raise ValueError(f"PFX missing key/cert: {path}")
    thumbprint = cert.fingerprint(hashes.SHA1()).hex().upper()
    private_key_pem = key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    ).decode()
    return {
        "private_key": private_key_pem,
        "thumbprint": thumbprint,
        "public_certificate": cert.public_bytes(serialization.Encoding.PEM).decode(),
    }
