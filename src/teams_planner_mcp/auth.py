"""Microsoft Entra delegated authentication for Microsoft Graph."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

import msal

GRAPH_SCOPES = [
    "User.Read",
    "Team.ReadBasic.All",
    "TeamMember.Read.All",
    "Channel.ReadBasic.All",
    "TeamsTab.Read.All",
    "Group.Read.All",
    "Tasks.ReadWrite",
]


class GraphAuth:
    def __init__(self) -> None:
        client_id = os.getenv("MS_CLIENT_ID")
        if not client_id:
            raise RuntimeError(
                "MS_CLIENT_ID is missing. Add it to .env after registering an Entra app."
            )

        self.tenant_id = os.getenv("MS_TENANT_ID", "common")
        self.cache_path = Path(os.getenv("MS_TOKEN_CACHE_PATH", ".token_cache.json"))
        self.cache = msal.SerializableTokenCache()
        self._load_cache()
        self.app = msal.PublicClientApplication(
            client_id=client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            token_cache=self.cache,
        )

    def _load_cache(self) -> None:
        try:
            if self.cache_path.exists():
                self.cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
        except (OSError, ValueError, json.JSONDecodeError) as exc:
            print(f"Warning: could not load token cache: {exc}", file=sys.stderr)

    def _save_cache(self) -> None:
        if self.cache.has_state_changed:
            self.cache_path.write_text(self.cache.serialize(), encoding="utf-8")

    def get_access_token(self) -> str:
        accounts = self.app.get_accounts()
        result: dict[str, Any] | None = None
        if accounts:
            result = self.app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])

        if not result or "access_token" not in result:
            flow = self.app.initiate_device_flow(scopes=GRAPH_SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(f"Could not start Microsoft login: {flow}")
            print(flow["message"], file=sys.stderr, flush=True)
            result = self.app.acquire_token_by_device_flow(flow)

        self._save_cache()
        if "access_token" not in result:
            message = result.get("error_description", "Microsoft login failed")
            raise RuntimeError(message)
        return str(result["access_token"])
