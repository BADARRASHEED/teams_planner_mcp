# Teams Planner MCP

Local MCP server for discovering Microsoft Teams/Planner resources and creating assigned Planner tasks from Codex.

## 1. Microsoft Entra app

Create an app registration in the Microsoft Entra admin center:

1. Go to **App registrations → New registration**.
2. Copy **Application (client) ID**.
3. Under **Authentication**, enable **Allow public client flows**.
4. Under **API permissions → Microsoft Graph → Delegated permissions**, add:
   `User.Read`, `Team.ReadBasic.All`, `TeamMember.Read.All`, `Channel.ReadBasic.All`, `TeamsTab.Read.All`, `Group.Read.All`, and `Tasks.ReadWrite`.
5. Ask your Microsoft 365 administrator for admin consent if your tenant requires it.

## 2. Environment

Copy `.env.example` to `.env` and replace `MS_CLIENT_ID` with the Application (client) ID. Keep `.env` private.

## 3. Install with uv (run these yourself)

```powershell
uv venv
uv pip install -r requirements.txt
```

## 4. Test the server

```powershell
uv run teams-planner-mcp
```

The first tool call prints a Microsoft device-login URL and code in the terminal. Complete that login once; the token cache is stored locally in `.token_cache.json`.

## 5. Register in Codex

Codex supports adding a local stdio MCP server from its CLI. Run this command from any terminal after the environment is ready:

```powershell
codex mcp add teams_planner_mcp -- uv --directory C:\Users\badar-butt\Desktop\teams_planner_mcp run teams-planner-mcp
```

You can verify the registration with:

```powershell
codex mcp get teams_planner_mcp
```

After registration, ask Codex to call `list_teams`, then `list_channels`, `list_channel_tabs`, `list_members`, and `list_plans` before creating a task. Planner tasks are stored in the plan underlying the channel's Planner tab, so always verify the selected Team, plan, and member ID when names are ambiguous.
