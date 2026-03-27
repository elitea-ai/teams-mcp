"""Microsoft Graph API HTTP helpers."""
from __future__ import annotations

import logging
from typing import Any

import httpx

from teams_mcp.auth import TeamsAuth

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

_NOT_AUTHENTICATED = "Not authenticated. Call the 'login' tool first."

_loggers_silenced = False


def _silence_noisy_loggers() -> None:
    """Suppress httpx/mcp request-level chatter. Called before each API call."""
    global _loggers_silenced
    if _loggers_silenced:
        return
    for name in ("httpx", "httpcore", "mcp.server", "mcp.server.lowlevel"):
        logging.getLogger(name).setLevel(logging.WARNING)
    _loggers_silenced = True


async def _graph_get(
    auth: TeamsAuth, endpoint: str, params: dict[str, str] | None = None
) -> dict[str, Any]:
    """GET from Microsoft Graph API."""
    _silence_noisy_loggers()
    token = auth.get_token()
    if not token:
        return {"error": _NOT_AUTHENTICATED}
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.get(
            f"{GRAPH_BASE}{endpoint}",
            headers={"Authorization": f"Bearer {token}"},
            params=params or {},
        )
    if resp.status_code == 401:
        return {"error": "Token expired or invalid. Call 'login' to re-authenticate."}
    if resp.status_code == 403:
        return {
            "error": (
                f"Permission denied for {endpoint}. "
                "This scope may require admin consent in your organization."
            )
        }
    if resp.status_code == 404:
        return {"error": f"Not found: {endpoint}", "code": "NotFound"}
    resp.raise_for_status()
    return resp.json()


async def _graph_get_paged(
    auth: TeamsAuth,
    endpoint: str,
    params: dict[str, str] | None = None,
    max_pages: int = 10,
) -> dict[str, Any]:
    """GET from Microsoft Graph API with pagination.

    Follows ``@odata.nextLink`` up to *max_pages* times, collecting all
    ``value`` items into one list.
    """
    _silence_noisy_loggers()
    token = auth.get_token()
    if not token:
        return {"error": _NOT_AUTHENTICATED}

    all_items: list[dict[str, Any]] = []
    url: str | None = f"{GRAPH_BASE}{endpoint}"
    current_params: dict[str, str] | None = params

    async with httpx.AsyncClient(timeout=30.0) as client:
        for _ in range(max_pages):
            if url is None:
                break
            resp = await client.get(
                url,
                headers={"Authorization": f"Bearer {token}"},
                params=current_params or {},
            )
            if resp.status_code == 401:
                return {"error": "Token expired or invalid. Call 'login' to re-authenticate."}
            if resp.status_code in (403, 404):
                break
            resp.raise_for_status()
            data = resp.json()
            all_items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            current_params = None

    return {"value": all_items}


async def _graph_post(
    auth: TeamsAuth, endpoint: str, body: dict[str, Any]
) -> dict[str, Any]:
    """POST to Microsoft Graph API."""
    _silence_noisy_loggers()
    token = auth.get_token()
    if not token:
        return {"error": _NOT_AUTHENTICATED}
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.post(
            f"{GRAPH_BASE}{endpoint}",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=body,
        )
    if resp.status_code == 401:
        return {"error": "Token expired or invalid. Call 'login' to re-authenticate."}
    if resp.status_code == 403:
        return {
            "error": (
                f"Permission denied for {endpoint}. "
                "This scope may require admin consent in your organization."
            )
        }
    resp.raise_for_status()
    return resp.json()
