"""FastMCP server definition and tool handlers."""
from __future__ import annotations

import os
from typing import Any

import httpx
from mcp.server.fastmcp import FastMCP

from mcp_server_teams.auth import TeamsAuth
from mcp_server_teams.contacts import (
    _load_contacts,
    _resolve_sender,
    _search_contacts,
    _update_contacts_from_chats,
    _update_contacts_from_members,
)
from mcp_server_teams.graph import (
    _NOT_AUTHENTICATED,
    _graph_get,
    _graph_get_paged,
    _graph_post,
    _silence_noisy_loggers,
)

# ---------------------------------------------------------------------------
# FastMCP server + tools
# ---------------------------------------------------------------------------

_page_links: dict[str, str] = {}

mcp = FastMCP(
    "teams",
    instructions=(
        "MS Teams MCP server. Call 'login' first to authenticate via device code flow. "
        "Then use chat and team tools to read and send messages.\n\n"
        "Use 'find-chat' to search for a chat by person name before reading messages. "
        "Use 'list-chat-members' to see who is in a chat. "
        "The server maintains a contacts cache so repeated lookups are fast."
    ),
)
auth = TeamsAuth()


@mcp.tool()
async def login() -> dict[str, Any]:
    """Start authentication to Microsoft Teams via device code flow.

    Returns a user_code and URL. The user must open the URL in a browser
    and enter the code. Then call 'login-complete' to finish authentication.
    """
    return await auth.device_code_login_start()


@mcp.tool(name="login-complete")
async def login_complete() -> dict[str, Any]:
    """Complete the device code login after user entered the code in browser.

    Call this after the user has entered the code from 'login' at the URL.
    Blocks until authentication completes (up to 15 minutes).
    """
    return await auth.device_code_login_complete()


@mcp.tool()
def logout() -> dict[str, str]:
    """Log out of Microsoft Teams and clear cached tokens."""
    auth.logout()
    return {"status": "logged out", "message": "Token cache cleared. Call 'login' to re-authenticate."}


@mcp.tool(name="list-chats")
async def list_chats(
    limit: int = 50, expand_members: bool = False, slim: bool = False,
) -> dict[str, Any]:
    """List your recent Teams chats.

    Args:
        limit: Max chats per page (1-50). One page per call.
        expand_members: If True, include member list per chat (slower, larger response).
        slim: If True, return only id, chatType, topic (smallest response).
    """
    page_size = min(max(1, limit), 50)
    params: dict[str, str] = {"$top": str(page_size)}
    if expand_members:
        params["$expand"] = "members"
    params["$orderby"] = "lastMessagePreview/createdDateTime desc"
    if not slim:
        params["$select"] = (
            "id,chatType,topic,createdDateTime,lastMessagePreview"
        )
    else:
        params["$select"] = "id,chatType,topic"

    data = await _graph_get(auth, "/chats", params)
    if "error" in data:
        return data

    raw_chats = data.get("value", [])
    has_more = "@odata.nextLink" in data

    _update_contacts_from_chats(raw_chats)

    chats = []
    for c in raw_chats:
        chat: dict[str, Any] = {
            "id": c.get("id"),
            "chatType": c.get("chatType"),
            "topic": c.get("topic"),
        }
        if not slim:
            chat["createdDateTime"] = c.get("createdDateTime")
            preview = c.get("lastMessagePreview")
            if preview:
                chat["lastMessage"] = {
                    "from": (preview.get("from") or {}).get("user", {}).get("displayName"),
                    "content": (preview.get("body") or {}).get("content", "")[:200],
                    "createdDateTime": preview.get("createdDateTime"),
                }
            if expand_members:
                chat["members"] = [
                    {
                        "displayName": m.get("displayName") or "",
                        "email": m.get("email") or "",
                    }
                    for m in c.get("members", [])
                ]
        chats.append(chat)

    next_link = data.get("@odata.nextLink")
    if next_link:
        _page_links["chats"] = next_link

    return {"chats": chats, "count": len(chats), "hasMore": has_more}


@mcp.tool(name="list-chats-next")
async def list_chats_next(slim: bool = False) -> dict[str, Any]:
    """Fetch the next page of chats (call after list-chats returned hasMore=true).

    Args:
        slim: If True, return only id, chatType, topic.
    """
    _silence_noisy_loggers()
    next_link = _page_links.get("chats")
    if not next_link:
        return {"chats": [], "count": 0, "hasMore": False, "error": "No next page. Call list-chats first."}

    token = auth.get_token()
    if not token:
        return {"error": _NOT_AUTHENTICATED}

    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.get(next_link, headers={"Authorization": f"Bearer {token}"})
    if resp.status_code == 401:
        return {"error": "Token expired or invalid. Call 'login' to re-authenticate."}
    if resp.status_code in (403, 404):
        _page_links.pop("chats", None)
        return {"chats": [], "count": 0, "hasMore": False}
    resp.raise_for_status()
    data = resp.json()

    has_more = "@odata.nextLink" in data
    if has_more:
        _page_links["chats"] = data["@odata.nextLink"]
    else:
        _page_links.pop("chats", None)

    chats = []
    for c in data.get("value", []):
        chat: dict[str, Any] = {
            "id": c.get("id"),
            "chatType": c.get("chatType"),
            "topic": c.get("topic"),
        }
        if not slim:
            chat["createdDateTime"] = c.get("createdDateTime")
            preview = c.get("lastMessagePreview")
            if preview:
                chat["lastMessage"] = {
                    "from": (preview.get("from") or {}).get("user", {}).get("displayName"),
                    "content": (preview.get("body") or {}).get("content", "")[:200],
                    "createdDateTime": preview.get("createdDateTime"),
                }
        chats.append(chat)
    return {"chats": chats, "count": len(chats), "hasMore": has_more}


@mcp.tool(name="list-chat-messages")
async def list_chat_messages(chatId: str, limit: int = 30) -> dict[str, Any]:
    """Read messages from a specific chat.

    Args:
        chatId: The chat ID (from list-chats or find-chat results).
        limit: Number of messages to return (1-50, default 30).
    """
    limit = max(1, min(limit, 50))
    data = await _graph_get(auth, f"/chats/{chatId}/messages", {"$top": str(limit)})
    if "error" in data:
        return data
    user_cache = _load_contacts().get("users", {})
    messages = []
    for m in data.get("value", []):
        from_user = (m.get("from") or {}).get("user") or {}
        raw_mentions = m.get("mentions") or []
        mentions = [
            {
                "displayName": (mn.get("mentioned") or {}).get("user", {}).get("displayName", ""),
                "userId": (mn.get("mentioned") or {}).get("user", {}).get("id", ""),
            }
            for mn in raw_mentions
            if (mn.get("mentioned") or {}).get("user")
        ]
        msg: dict[str, Any] = {
            "id": m.get("id"),
            "from": _resolve_sender(from_user, user_cache),
            "body": (m.get("body") or {}).get("content", ""),
            "contentType": (m.get("body") or {}).get("contentType"),
            "createdDateTime": m.get("createdDateTime"),
            "messageType": m.get("messageType"),
            "attachments": m.get("attachments"),
            "mentions": mentions if mentions else None,
        }
        messages.append(msg)
    return {"messages": messages, "count": len(messages)}


@mcp.tool(name="list-chat-members")
async def list_chat_members(chatId: str) -> dict[str, Any]:
    """List members of a specific chat.

    Args:
        chatId: The chat ID (from list-chats or find-chat results).
    """
    data = await _graph_get(auth, f"/chats/{chatId}/members")
    if "error" in data:
        return data
    members = []
    for m in data.get("value", []):
        members.append({
            "displayName": m.get("displayName") or "",
            "email": m.get("email") or "",
            "roles": m.get("roles") or [],
        })
    _update_contacts_from_members(chatId, members)
    return {"members": members, "count": len(members)}


@mcp.tool(name="find-chat")
async def find_chat(query: str) -> dict[str, Any]:
    """Find a chat by person name, topic, or keyword.

    Args:
        query: Person name, chat topic, or keyword to search for.
    """
    query_lower = query.lower()
    matches = _search_contacts(query_lower)

    if matches:
        return {"matches": matches[:10], "count": len(matches), "source": "cache"}

    fresh = await list_chats(limit=50)
    if "error" in fresh:
        return fresh

    matches = _search_contacts(query_lower)
    if matches:
        return {"matches": matches[:10], "count": len(matches), "source": "fresh"}
    return {"matches": [], "count": 0, "message": f"No chats found matching '{query}'."}


@mcp.tool(name="send-chat-message")
async def send_chat_message(
    chatId: str, content: str, contentType: str = "text"
) -> dict[str, Any]:
    """Send a message to a Teams chat.

    Args:
        chatId: The chat ID (from list-chats or find-chat results).
        content: Message content (text or HTML).
        contentType: 'text' or 'html' (default 'text').
    """
    body = {"body": {"contentType": contentType, "content": content}}
    data = await _graph_post(auth, f"/chats/{chatId}/messages", body)
    if "error" in data:
        return data
    return {
        "status": "sent",
        "messageId": data.get("id"),
        "createdDateTime": data.get("createdDateTime"),
    }


@mcp.tool(name="list-joined-teams")
async def list_joined_teams() -> dict[str, Any]:
    """List Microsoft Teams you have joined."""
    data = await _graph_get(auth, "/me/joinedTeams")
    if "error" in data:
        return data
    teams = [
        {
            "id": t.get("id"),
            "displayName": t.get("displayName"),
            "description": t.get("description"),
        }
        for t in data.get("value", [])
    ]
    return {"teams": teams, "count": len(teams)}


@mcp.tool(name="list-team-channels")
async def list_team_channels(teamId: str) -> dict[str, Any]:
    """List channels in a specific team.

    Args:
        teamId: The team ID (from list-joined-teams).
    """
    data = await _graph_get(auth, f"/teams/{teamId}/channels")
    if "error" in data:
        return data
    channels = [
        {
            "id": ch.get("id"),
            "displayName": ch.get("displayName"),
            "description": ch.get("description"),
            "membershipType": ch.get("membershipType"),
        }
        for ch in data.get("value", [])
    ]
    return {"channels": channels, "count": len(channels)}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

async def serve(
    client_id: str | None = None,
    tenant_id: str | None = None,
) -> None:
    """Run the MCP server with stdio transport."""
    global auth
    # Allow CLI args to override env vars
    if client_id:
        os.environ["TEAMS_CLIENT_ID"] = client_id
    if tenant_id:
        os.environ["TEAMS_TENANT_ID"] = tenant_id
    auth = TeamsAuth(client_id=client_id, tenant_id=tenant_id)
    await mcp.run_async(transport="stdio")


def main() -> None:
    """CLI entry point (legacy, prefer package __init__.main)."""
    import asyncio
    asyncio.run(serve())
