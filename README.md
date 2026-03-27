# mcp-server-teams

A Model Context Protocol server providing tools to read and send Microsoft Teams messages.

Authenticate with your Microsoft account using device-code flow and give any MCP-compatible client — Claude Desktop, Cursor, VS Code, or your own agent — read/write access to Teams chats, teams, and channels.

## Features

| Tool | Description |
|------|-------------|
| `login` | Start device-code authentication (returns code + URL) |
| `login-complete` | Complete authentication after entering the code |
| `logout` | Clear cached tokens |
| `list-chats` | List recent chats with pagination |
| `list-chats-next` | Fetch next page of chats |
| `list-chat-messages` | Read messages from a chat |
| `list-chat-members` | List members of a chat |
| `find-chat` | Search chats by person name or topic |
| `send-chat-message` | Send a message to a chat |
| `list-joined-teams` | List teams you belong to |
| `list-team-channels` | List channels in a team |

## Quick Start

```bash
# Run directly with uvx (no install needed)
uvx mcp-server-teams
```

Then call `login` from your MCP client. You'll receive a code — enter it at the URL in your browser. After that, call `login-complete` to finish authentication.

### CLI Options

```bash
mcp-server-teams --client-id YOUR_APP_ID --tenant-id YOUR_TENANT_ID
```

## MCP Client Configuration

### Claude Desktop / Cursor

Add to your MCP config:

```json
{
  "mcpServers": {
    "teams": {
      "command": "uvx",
      "args": ["mcp-server-teams"]
    }
  }
}
```

### With custom Azure AD app

```json
{
  "mcpServers": {
    "teams": {
      "command": "uvx",
      "args": ["mcp-server-teams"],
      "env": {
        "TEAMS_CLIENT_ID": "your-app-client-id",
        "TEAMS_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

### With CLI arguments

```json
{
  "mcpServers": {
    "teams": {
      "command": "uvx",
      "args": ["mcp-server-teams", "--client-id", "your-app-id", "--tenant-id", "your-tenant-id"]
    }
  }
}
```

## Configuration

| Environment Variable | Default | Description |
|---------------------|---------|-------------|
| `TEAMS_CLIENT_ID` | Pre-registered public client | Azure AD application (client) ID |
| `TEAMS_TENANT_ID` | `common` | Azure AD directory (tenant) ID |
| `TEAMS_TOKEN_CACHE` | `~/.teams-mcp/token_cache.json` | Path to token cache file |
| `TEAMS_CACHE_DIR` | `~/.teams-mcp/` | Directory for contacts cache |

## Authentication

The server uses [device code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code) — a two-step process designed for headless environments:

1. Call `login` → receive a code and URL
2. Open the URL in any browser, enter the code, sign in
3. Call `login-complete` → server exchanges the code for tokens

Tokens are cached locally and refreshed automatically. Call `logout` to clear them.

### Required Scopes

- `User.Read` — read your profile
- `Chat.Read` — read chat messages
- `ChatMessage.Send` — send chat messages
- `Team.ReadBasic.All` — list teams
- `Channel.ReadBasic.All` — list channels

All scopes use **delegated** (user-consent) permissions — no admin approval needed.

## Development

```bash
git clone https://github.com/elitea-ai/teams-mcp.git
cd teams-mcp
uv sync --dev
uv run pytest
```

## License

MIT
