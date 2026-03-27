"""Contacts and chat cache for fast lookups."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

_CACHE_DIR = Path(os.environ.get("TEAMS_CACHE_DIR", Path.home() / ".teams-mcp"))
_CONTACTS_PATH = _CACHE_DIR / "contacts.json"


def _load_contacts() -> dict[str, Any]:
    """Load the contacts/chat cache from disk."""
    if _CONTACTS_PATH.is_file():
        try:
            return json.loads(_CONTACTS_PATH.read_text("utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {"chats": {}, "users": {}}


def _save_contacts(data: dict[str, Any]) -> None:
    """Persist the contacts/chat cache."""
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)
    _CONTACTS_PATH.write_text(json.dumps(data, indent=2, ensure_ascii=False), "utf-8")


def _update_contacts_from_chats(chats: list[dict[str, Any]]) -> None:
    """Update the contacts cache with chat info (members, topics)."""
    contacts = _load_contacts()
    for chat in chats:
        cid = chat.get("id", "")
        if not cid:
            continue
        entry = contacts["chats"].get(cid, {})
        entry["id"] = cid
        entry["chatType"] = chat.get("chatType")
        entry["topic"] = chat.get("topic")
        if "members" in chat:
            entry["members"] = chat["members"]
            for m in chat["members"]:
                if m.get("displayName"):
                    existing = contacts["users"].get(m["displayName"].lower(), {})
                    user_entry = {
                        "displayName": m["displayName"],
                        "chatIds": list(set(
                            existing.get("chatIds", []) + [cid]
                        )),
                    }
                    email = m.get("email", "") or existing.get("email", "")
                    if email:
                        user_entry["email"] = email
                    contacts["users"][m["displayName"].lower()] = user_entry
        if chat.get("lastMessage", {}).get("from"):
            name = chat["lastMessage"]["from"]
            entry.setdefault("members", [])
            if not any(m.get("displayName") == name for m in entry.get("members", [])):
                entry["members"].append({"displayName": name})
            contacts["users"].setdefault(name.lower(), {
                "displayName": name, "chatIds": [],
            })
            if cid not in contacts["users"][name.lower()]["chatIds"]:
                contacts["users"][name.lower()]["chatIds"].append(cid)
        contacts["chats"][cid] = entry
    _save_contacts(contacts)


def _update_contacts_from_members(chat_id: str, members: list[dict[str, Any]]) -> None:
    """Update the contacts cache with chat member info."""
    contacts = _load_contacts()
    entry = contacts["chats"].get(chat_id, {"id": chat_id})
    entry["members"] = members
    contacts["chats"][chat_id] = entry
    for m in members:
        name = m.get("displayName") or ""
        if name:
            user = contacts["users"].get(name.lower(), {
                "displayName": name, "chatIds": [],
            })
            user["displayName"] = name
            email = m.get("email", "")
            if email:
                user["email"] = email
            if chat_id not in user.get("chatIds", []):
                user.setdefault("chatIds", []).append(chat_id)
            contacts["users"][name.lower()] = user
    _save_contacts(contacts)


def _resolve_sender(
    from_user: dict[str, Any],
    user_cache: dict[str, Any] | None = None,
) -> dict[str, Any] | None:
    """Build a structured sender dict from a Graph API ``from.user`` object."""
    display_name = from_user.get("displayName") or ""
    user_id = from_user.get("id") or ""
    if not display_name and not user_id:
        return None

    email = ""
    if user_cache:
        if display_name:
            user_entry = user_cache.get(display_name.lower(), {})
            email = user_entry.get("email", "")
        if not email and user_id:
            for _k, entry in user_cache.items():
                if entry.get("userId") == user_id:
                    email = entry.get("email", "")
                    break

    return {
        "displayName": display_name,
        "userId": user_id,
        "email": email,
    }


def _search_contacts(query_lower: str) -> list[dict[str, Any]]:
    """Search contacts cache by name or topic. Returns lightweight matches."""
    contacts = _load_contacts()
    matches: list[dict[str, Any]] = []
    seen_ids: set[str] = set()

    def _add_match(
        cid: str, chat_entry: dict, matched_user: str = "",
    ) -> None:
        if cid in seen_ids:
            return
        seen_ids.add(cid)
        members = chat_entry.get("members", [])
        match: dict[str, Any] = {
            "id": cid,
            "topic": chat_entry.get("topic"),
            "chatType": chat_entry.get("chatType"),
            "memberCount": len(members),
        }
        if matched_user:
            match["matchedUser"] = matched_user
        matches.append(match)

    for name_key, user_info in contacts.get("users", {}).items():
        if query_lower in name_key:
            for cid in user_info.get("chatIds", []):
                chat_entry = contacts["chats"].get(cid, {})
                if chat_entry:
                    _add_match(cid, chat_entry, user_info.get("displayName", ""))

    for cid, chat_entry in contacts.get("chats", {}).items():
        topic = (chat_entry.get("topic") or "").lower()
        if query_lower in topic:
            _add_match(cid, chat_entry)

    return matches
