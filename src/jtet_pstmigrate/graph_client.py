"""Thin Graph REST client built on httpx.

Why not the official msgraph-sdk?
  - For this workload (POST /messages, PATCH, mailFolders create) we issue a
    handful of well-defined REST calls. The SDK adds significant abstraction,
    pulls in azure-core, and obscures retry/throttle behaviour we want to
    own explicitly.
  - Direct REST means we can honour Retry-After exactly, log raw status codes,
    and surface throttle reasons cleanly in the run summary.
"""

from __future__ import annotations

import random
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

import httpx
from loguru import logger

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import ThrottleConfig

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class GraphError(Exception):
    """Non-retryable Graph API error after retries are exhausted."""

    def __init__(self, status: int, body: Any, request_id: str | None = None):
        self.status = status
        self.body = body
        self.request_id = request_id
        super().__init__(f"Graph {status}: {body!r} (request-id={request_id})")


class ThrottledError(GraphError):
    """429 / 503 that we hit our retry budget on."""


@dataclass(slots=True)
class ThrottleStats:
    requests: int = 0
    retries_429: int = 0
    retries_5xx: int = 0
    total_backoff_seconds: float = 0.0


class GraphClient:
    """Synchronous Graph client. Thread-safe — share one across workers.

    Holds an AppPool. Each request either uses an explicitly-named app (when
    the orchestrator wants stable attribution per upload) or a round-robin
    selection. Stats are tracked per-app so the summary shows whether load
    actually balanced or one app got hot.
    """

    def __init__(self, pool: AppPool, cfg: ThrottleConfig):
        self._pool = pool
        self._cfg = cfg
        self._stats_lock = threading.Lock()
        self.stats: dict[str, ThrottleStats] = {name: ThrottleStats() for name in pool.names}
        # Single httpx.Client is internally thread-safe for sync requests.
        self._client = httpx.Client(
            base_url=GRAPH_BASE,
            timeout=httpx.Timeout(cfg.request_timeout_seconds),
            limits=httpx.Limits(max_connections=64, max_keepalive_connections=32),
        )

    @property
    def pool(self) -> AppPool:
        return self._pool

    def close(self) -> None:
        self._client.close()

    def __enter__(self) -> GraphClient:
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    def request(
        self,
        method: str,
        path: str,
        *,
        json: Any = None,
        content: bytes | None = None,
        headers: dict[str, str] | None = None,
        params: dict[str, Any] | None = None,
        expect_status: tuple[int, ...] = (200, 201, 202, 204),
        app_id: str | None = None,
    ) -> httpx.Response:
        """Issue a Graph request with throttling + retry.

        If `app_id` is None we round-robin from the pool; otherwise the named
        app is used (caller-controlled attribution). On terminal failure
        raises GraphError / ThrottledError.
        """
        chosen_app = app_id or self._pool.pick()
        attempt = 0
        while True:
            with self._stats_lock:
                self.stats[chosen_app].requests += 1
            request_headers = self._build_headers(
                headers,
                chosen_app,
                content_type="application/json" if json is not None else None,
            )
            try:
                response = self._client.request(
                    method,
                    path,
                    json=json,
                    content=content,
                    headers=request_headers,
                    params=params,
                )
            except (httpx.TransportError, httpx.TimeoutException) as e:
                if attempt >= self._cfg.max_retries:
                    raise GraphError(0, f"transport error: {e}") from e
                self._sleep(chosen_app, attempt, transient_reason=f"transport: {e}")
                attempt += 1
                continue

            if response.status_code in expect_status:
                return response

            if response.status_code in (429, 503):
                with self._stats_lock:
                    self.stats[chosen_app].retries_429 += 1
                if attempt >= self._cfg.max_retries:
                    raise ThrottledError(
                        response.status_code,
                        _safe_body(response),
                        response.headers.get("request-id"),
                    )
                retry_after = _parse_retry_after(response)
                self._sleep(
                    chosen_app,
                    attempt,
                    override_seconds=retry_after,
                    transient_reason=f"{response.status_code} throttled",
                )
                attempt += 1
                continue

            if 500 <= response.status_code < 600:
                with self._stats_lock:
                    self.stats[chosen_app].retries_5xx += 1
                if attempt >= self._cfg.max_retries:
                    raise GraphError(
                        response.status_code,
                        _safe_body(response),
                        response.headers.get("request-id"),
                    )
                self._sleep(chosen_app, attempt, transient_reason=f"{response.status_code} server error")
                attempt += 1
                continue

            raise GraphError(response.status_code, _safe_body(response), response.headers.get("request-id"))

    def _build_headers(
        self, extra: dict[str, str] | None, app_id: str, content_type: str | None
    ) -> dict[str, str]:
        token = self._pool.get(app_id).get()
        h = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        if content_type:
            h["Content-Type"] = content_type
        if extra:
            h.update(extra)
        return h

    def _sleep(
        self,
        app_id: str,
        attempt: int,
        *,
        override_seconds: float | None = None,
        transient_reason: str = "",
    ) -> None:
        if override_seconds is not None:
            delay = max(0.0, override_seconds)
        else:
            delay = min(
                self._cfg.max_backoff_seconds,
                self._cfg.initial_backoff_seconds * (self._cfg.backoff_multiplier**attempt),
            )
        jitter = delay * self._cfg.backoff_jitter * (random.random() * 2 - 1)
        delay = max(0.1, delay + jitter)
        with self._stats_lock:
            self.stats[app_id].total_backoff_seconds += delay
        logger.bind(ctx=f"graph[{app_id}]").warning(
            "Backoff {:.2f}s (attempt {}/{}) — {}",
            delay,
            attempt + 1,
            self._cfg.max_retries,
            transient_reason,
        )
        time.sleep(delay)

    # Convenience helpers ------------------------------------------------
    # All accept an optional app_id kwarg, forwarded to request().

    def get(self, path: str, **kwargs) -> httpx.Response:
        return self.request("GET", path, **kwargs)

    def post(self, path: str, **kwargs) -> httpx.Response:
        return self.request("POST", path, **kwargs)

    def patch(self, path: str, **kwargs) -> httpx.Response:
        return self.request("PATCH", path, **kwargs)

    def delete(self, path: str, **kwargs) -> httpx.Response:
        return self.request("DELETE", path, **kwargs)

    def put(self, path: str, **kwargs) -> httpx.Response:
        return self.request("PUT", path, **kwargs)


def _parse_retry_after(response: httpx.Response) -> float | None:
    raw = response.headers.get("Retry-After")
    if not raw:
        return None
    try:
        return float(raw)
    except ValueError:
        # HTTP-date form is rare but valid; fall back to default backoff.
        return None


def _safe_body(response: httpx.Response) -> Any:
    try:
        return response.json()
    except Exception:
        return response.text[:1000]


def safe_call(fn: Callable[[], Any], *, on_error_log: str = "graph") -> Any:
    """Wrap a call so per-message failures don't crash the worker.

    Returns (result, error). Caller decides what to do with error.
    """
    try:
        return fn(), None
    except GraphError as e:
        logger.bind(ctx=on_error_log).error("Graph call failed: {}", e)
        return None, e
