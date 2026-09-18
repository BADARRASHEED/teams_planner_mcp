"""Teams Planner MCP server."""

from __future__ import annotations

import json
from typing import Any

from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

from .auth import GraphAuth
from .graph_client import GraphClient

load_dotenv()
mcp = FastMCP("teams_planner_mcp")
_client: GraphClient | None = None


def client() -> GraphClient:
    global _client
    if _client is None:
        _client = GraphClient(GraphAuth())
    return _client


def result(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2)


@mcp.tool()
async def list_teams() -> str:
    """List Microsoft Teams that the signed-in user has joined."""
    return result(await client().joined_teams())


@mcp.tool()
async def list_channels(team_id: str) -> str:
    """List channels for a Team. Use list_teams first to get team_id."""
    return result(await client().channels(team_id))


@mcp.tool()
async def list_members(team_id: str) -> str:
    """List Team members and their IDs for reliable task assignment."""
    return result(await client().members(team_id))


@mcp.tool()
async def list_channel_tabs(team_id: str, channel_id: str) -> str:
    """List apps/tabs installed in a channel, including a possible Planner tab."""
    return result(await client().channel_tabs(team_id, channel_id))


@mcp.tool()
async def list_plans(team_id: str) -> str:
    """List Planner plans associated with a Team. In Graph, a standard Team uses its group ID."""
    return result(await client().plans(team_id))


@mcp.tool()
async def list_plan_tasks(plan_id: str) -> str:
    """List existing tasks in a Planner plan."""
    return result(await client().tasks(plan_id))


@mcp.tool()
async def create_planner_task(
    plan_id: str,
    title: str,
    assignee_id: str | None = None,
    due_date: str | None = None,
    description: str | None = None,
    bucket_id: str | None = None,
) -> str:
    """Create a Planner task. Resolve plan_id and assignee_id with discovery tools first."""
    if not title.strip():
        raise ValueError("title cannot be empty")
    task = await client().create_task(
        plan_id,
        title.strip(),
        assignee_id=assignee_id,
        due_date=due_date,
        bucket_id=bucket_id,
    )
    if description and task.get("id"):
        await client().add_task_description(task["id"], description)
    return result(
        {
            "created": True,
            "task_id": task.get("id"),
            "title": task.get("title", title),
            "plan_id": plan_id,
            "assignee_id": assignee_id,
            "due_date": task.get("dueDateTime", due_date),
        }
    )


def main() -> None:
    mcp.run(transport="stdio")
