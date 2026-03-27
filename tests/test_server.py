"""Smoke tests for mcp-server-teams."""
import pytest


def test_import():
    """Package imports without error."""
    from mcp_server_teams import __version__
    assert __version__


def test_server_has_tools():
    """FastMCP server registers all expected tools."""
    from mcp_server_teams.server import mcp
    tools = mcp._tool_manager._tools
    expected = {
        "login", "login-complete", "logout",
        "list-chats", "list-chats-next",
        "list-chat-messages", "list-chat-members",
        "find-chat", "send-chat-message",
        "list-joined-teams", "list-team-channels",
    }
    assert expected.issubset(set(tools.keys())), f"Missing tools: {expected - set(tools.keys())}"


def test_auth_class_instantiates():
    """TeamsAuth can be created without errors."""
    from mcp_server_teams.auth import TeamsAuth
    auth = TeamsAuth()
    assert auth._client_id
    assert auth._tenant_id
