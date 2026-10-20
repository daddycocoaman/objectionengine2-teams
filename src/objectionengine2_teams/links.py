"""Recognize documented Microsoft Teams chat and channel deep links."""

from dataclasses import dataclass
import json
import re
from typing import Literal
from urllib.parse import parse_qs, unquote, urlsplit
from uuid import UUID


_TEAMS_HOSTS = {"teams.microsoft.com", "teams.cloud.microsoft"}
_ID_RE = re.compile(r"^19:[A-Za-z0-9._@:-]{1,500}$")


@dataclass(frozen=True)
class TeamsTarget:
    kind: Literal["chat", "channel"]
    chat_id: str | None = None
    team_id: str | None = None
    channel_id: str | None = None
    tenant_id: str | None = None


def _uuid_parameter(query: dict[str, list[str]], name: str) -> str | None:
    values = query.get(name, [])
    if not values:
        return None
    if len(values) != 1:
        raise ValueError(f"Teams link has more than one {name} parameter")
    try:
        return str(UUID(values[0]))
    except ValueError as exc:
        raise ValueError(f"Teams link has an invalid {name} parameter") from exc


def _teams_id(raw: str) -> str:
    value = unquote(raw)
    if not _ID_RE.fullmatch(value):
        raise ValueError("Teams link has an invalid chat or channel ID")
    return value


def parse_teams_link(link: str) -> TeamsTarget:
    """Return the conversation addressed by a Teams deep link.

    A message link selects its containing chat or channel, not just one message.
    """
    parsed = urlsplit(link)
    if (
        parsed.scheme != "https"
        or parsed.hostname not in _TEAMS_HOSTS
        or parsed.username is not None
        or parsed.password is not None
        or parsed.port is not None
    ):
        raise ValueError("Expected an HTTPS link on teams.microsoft.com or teams.cloud.microsoft")

    parts = parsed.path.strip("/").split("/")
    query = parse_qs(parsed.query)
    tenant_id = _uuid_parameter(query, "tenantId")
    group_id = _uuid_parameter(query, "groupId")

    if len(parts) >= 4 and parts[:2] == ["l", "chat"] and parts[3] == "conversations":
        chat_id = _teams_id(parts[2])
        return TeamsTarget("chat", chat_id=chat_id, tenant_id=tenant_id)

    if len(parts) >= 4 and parts[:2] == ["l", "channel"]:
        channel_id = _teams_id(parts[2])
        if group_id is None:
            raise ValueError("Channel links must include groupId (the team ID)")
        return TeamsTarget("channel", team_id=group_id, channel_id=channel_id, tenant_id=tenant_id)

    if len(parts) == 4 and parts[:2] == ["l", "team"] and parts[3] == "conversations":
        channel_id = _teams_id(parts[2])
        if group_id is None:
            raise ValueError("Team links must include groupId (the team ID)")
        return TeamsTarget("channel", team_id=group_id, channel_id=channel_id, tenant_id=tenant_id)

    if len(parts) == 4 and parts[:2] == ["l", "message"]:
        conversation_id = _teams_id(parts[2])
        if not parts[3]:
            raise ValueError("Message links must contain a message ID")
        if group_id is not None:
            return TeamsTarget("channel", team_id=group_id, channel_id=conversation_id, tenant_id=tenant_id)
        context_values = query.get("context", [])
        if context_values:
            try:
                context = json.loads(context_values[0])
            except json.JSONDecodeError as exc:
                raise ValueError("Teams link has an invalid context parameter") from exc
            if not isinstance(context, dict) or context.get("contextType") != "chat":
                raise ValueError("A channel message link must include groupId")
        if conversation_id.endswith("@thread.tacv2"):
            raise ValueError("A channel message link must include groupId")
        return TeamsTarget("chat", chat_id=conversation_id, tenant_id=tenant_id)

    raise ValueError("Unsupported Teams link; copy a chat, channel, or message link from Teams")
