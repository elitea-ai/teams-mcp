"""MSAL device-code authentication with persistent token cache."""
from __future__ import annotations

import asyncio
import os
import sys
from pathlib import Path
from typing import Any

import msal

# Pre-registered public client application.
# Users can override with TEAMS_CLIENT_ID env var for their own app registration.
DEFAULT_CLIENT_ID = "084a3e9f-a9f4-43f7-89f9-d229cf97853e"

SCOPES = [
    "User.Read",
    "Chat.Read",
    "ChatMessage.Send",
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
]


class TeamsAuth:
    """MSAL device-code auth with persistent token cache."""

    def __init__(
        self,
        client_id: str | None = None,
        tenant_id: str | None = None,
    ) -> None:
        self._client_id = (
            client_id
            or os.environ.get("TEAMS_CLIENT_ID", "")
            or DEFAULT_CLIENT_ID
        )
        self._tenant_id = (
            tenant_id
            or os.environ.get("TEAMS_TENANT_ID", "")
            or "common"
        )
        self._cache_path = Path(
            os.environ.get(
                "TEAMS_TOKEN_CACHE",
                Path.home() / ".teams-mcp" / "token_cache.json",
            )
        )
        self._cache = msal.SerializableTokenCache()
        self._load_cache()
        self._app = msal.PublicClientApplication(
            self._client_id,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            token_cache=self._cache,
        )

    # -- cache persistence --

    def _load_cache(self) -> None:
        if self._cache_path.is_file():
            self._cache.deserialize(self._cache_path.read_text("utf-8"))

    def _save_cache(self) -> None:
        if self._cache.has_state_changed:
            self._cache_path.parent.mkdir(parents=True, exist_ok=True)
            self._cache_path.write_text(self._cache.serialize(), "utf-8")
            try:
                self._cache_path.chmod(0o600)
            except OSError:
                pass

    # -- token acquisition --

    def get_token(self) -> str | None:
        """Try silent token acquisition from cache. Returns token or None."""
        accounts = self._app.get_accounts()
        if not accounts:
            return None
        result = self._app.acquire_token_silent(SCOPES, account=accounts[0])
        self._save_cache()
        if result and "access_token" in result:
            return result["access_token"]
        return None

    async def device_code_login_start(self) -> dict[str, Any]:
        """Start device code flow. Returns user_code and URL immediately."""
        flow = self._app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            return {"error": flow.get("error_description", "Failed to start device flow")}

        print(flow["message"], file=sys.stderr, flush=True)

        # Store flow for completion step
        self._pending_flow = flow
        return {
            "status": "pending",
            "user_code": flow["user_code"],
            "verification_uri": flow.get("verification_uri", "https://microsoft.com/devicelogin"),
            "message": flow["message"],
        }

    async def device_code_login_complete(self) -> dict[str, Any]:
        """Complete device code flow. Call after user entered the code in browser."""
        flow = getattr(self, "_pending_flow", None)
        if not flow:
            return {"error": "No pending login flow. Call 'login' first."}

        result = await asyncio.to_thread(
            self._app.acquire_token_by_device_flow, flow
        )
        self._save_cache()
        self._pending_flow = None

        if "access_token" in result:
            account = result.get("id_token_claims", {}).get("preferred_username", "")
            return {"status": "authenticated", "account": account}
        return {"error": result.get("error_description", "Authentication failed")}

    def logout(self) -> None:
        """Clear all cached tokens."""
        if self._cache_path.is_file():
            self._cache_path.unlink()
        self._cache = msal.SerializableTokenCache()
        self._app = msal.PublicClientApplication(
            self._client_id,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            token_cache=self._cache,
        )
