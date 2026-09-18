"""Small, focused Microsoft Graph client used by the MCP tools."""

from __future__ import annotations

from datetime import datetime
from typing import Any

import httpx

from .auth import GraphAuth

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class GraphClient:
    def __init__(self, auth: GraphAuth) -> None:
        self.auth = auth

    async def request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, str] | None = None,
        json_body: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        token = self.auth.get_access_token()
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.request(
                method,
                f"{GRAPH_BASE}{path}",
                headers=headers,
                params=params,
                json=json_body,
            )
        if response.is_error:
            try:
                detail = response.json().get("error", {}).get("message", response.text)
            except ValueError:
                detail = response.text
            raise RuntimeError(f"Microsoft Graph {response.status_code}: {detail}")
        if not response.content:
            return {}
        return response.json()

    async def get_all(self, path: str) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        next_path: str | None = path
        while next_path:
            data = await self.request("GET", next_path)
            items.extend(data.get("value", []))
            next_link = data.get("@odata.nextLink")
            next_path = next_link.replace(GRAPH_BASE, "") if next_link else None
        return items

    async def joined_teams(self) -> list[dict[str, Any]]:
        return await self.get_all("/me/joinedTeams")

    async def channels(self, team_id: str) -> list[dict[str, Any]]:
        return await self.get_all(f"/teams/{team_id}/channels")

    async def members(self, team_id: str) -> list[dict[str, Any]]:
        return await self.get_all(f"/teams/{team_id}/members")

    async def channel_tabs(self, team_id: str, channel_id: str) -> list[dict[str, Any]]:
        return await self.get_all(f"/teams/{team_id}/channels/{channel_id}/tabs")

    async def plans(self, group_id: str) -> list[dict[str, Any]]:
        return await self.get_all(f"/groups/{group_id}/planner/plans")

    async def tasks(self, plan_id: str) -> list[dict[str, Any]]:
        return await self.get_all(f"/planner/plans/{plan_id}/tasks")

    async def create_task(
        self,
        plan_id: str,
        title: str,
        *,
        assignee_id: str | None = None,
        due_date: str | None = None,
        bucket_id: str | None = None,
        percent_complete: int = 0,
    ) -> dict[str, Any]:
        body: dict[str, Any] = {
            "planId": plan_id,
            "title": title,
            "percentComplete": percent_complete,
        }
        if due_date:
            normalized_due_date = due_date.strip()
            if len(normalized_due_date) == 10:
                normalized_due_date += "T23:59:59+00:00"
            parsed = datetime.fromisoformat(normalized_due_date.replace("Z", "+00:00"))
            if parsed.tzinfo is None:
                normalized_due_date = parsed.isoformat() + "Z"
                parsed = datetime.fromisoformat(
                    normalized_due_date.replace("Z", "+00:00")
                )
            body["dueDateTime"] = parsed.isoformat().replace("+00:00", "Z")
        if bucket_id:
            body["bucketId"] = bucket_id
        if assignee_id:
            body["assignments"] = {
                assignee_id: {
                    "@odata.type": "microsoft.graph.plannerAssignment",
                    "orderHint": " !",
                }
            }
        return await self.request("POST", "/planner/tasks", json_body=body)

    async def add_task_description(self, task_id: str, description: str) -> None:
        details = await self.request("GET", f"/planner/tasks/{task_id}/details")
        etag = details.get("@odata.etag")
        token = self.auth.get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "If-Match": etag or "*",
            "Content-Type": "application/json",
        }
        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.request(
                "PATCH",
                f"{GRAPH_BASE}/planner/tasks/{task_id}/details",
                headers=headers,
                json={"description": description},
            )
        if response.is_error:
            raise RuntimeError(f"Could not add task description: {response.text}")
