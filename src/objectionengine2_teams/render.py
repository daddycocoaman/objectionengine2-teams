"""Public Teams-to-video API."""

from pathlib import Path

from objectionengine2.beans.comment import Comment
from objectionengine2.renderer import render_comment_list

from .auth import BROWSER_REDIRECT_URI, OFFICE_CLIENT_ID, acquire_access_token
from .convert import messages_to_comments
from .graph import GraphClient
from .links import TeamsTarget, parse_teams_link


def _end_target(start: TeamsTarget, end_link: str | None) -> TeamsTarget | None:
    if end_link is None:
        return None
    end = parse_teams_link(end_link)
    if start.start_message_id is None or end.start_message_id is None:
        raise ValueError("Both start and end links must point to messages")
    if start.kind != end.kind:
        raise ValueError("Start and end links must point to the same conversation")
    if start.kind == "chat" and start.chat_id != end.chat_id:
        raise ValueError("Start and end links must point to the same chat")
    if start.kind == "channel":
        if (start.team_id, start.channel_id) != (end.team_id, end.channel_id):
            raise ValueError("Start and end links must point to the same channel")
        if start.parent_message_id != end.parent_message_id:
            raise ValueError("Start and end links must point to the same channel thread")
    if start.tenant_id and end.tenant_id and start.tenant_id != end.tenant_id:
        raise ValueError("Start and end links must point to the same tenant")
    return end


def fetch_comments(
    link: str,
    *,
    end_link: str | None = None,
    max_messages: int | None = None,
    client_id: str = OFFICE_CLIENT_ID,
    auth_method: str = "browser",
    redirect_uri: str = BROWSER_REDIRECT_URI,
) -> list[Comment]:
    """Read a Teams conversation from the linked message through an optional end."""
    target = parse_teams_link(link)
    end = _end_target(target, end_link)
    tenant_id = target.tenant_id or (end.tenant_id if end else None)
    token = acquire_access_token(
        target.kind,
        tenant_id=tenant_id,
        client_id=client_id,
        auth_method=auth_method,
        redirect_uri=redirect_uri,
    )
    with GraphClient(token) as graph:
        messages = (
            graph.iter_messages(target, end_message_id=end.start_message_id)
            if end else graph.iter_messages(target)
        )
        comments = messages_to_comments(
            messages,
            start_message_id=target.start_message_id,
            end_message_id=end.start_message_id if end else None,
            max_messages=max_messages,
        )
    return comments


def render_teams_link(
    link: str,
    output_filename: str | Path = "teams.mp4",
    *,
    end_link: str | None = None,
    max_messages: int | None = None,
    client_id: str = OFFICE_CLIENT_ID,
    auth_method: str = "browser",
    redirect_uri: str = BROWSER_REDIRECT_URI,
    music_code: str = "pwr",
    resolution_scale: float = 1,
) -> Path:
    """Render a Teams conversation, starting at the linked message, to MP4."""
    comments = fetch_comments(
        link,
        end_link=end_link,
        max_messages=max_messages,
        client_id=client_id,
        auth_method=auth_method,
        redirect_uri=redirect_uri,
    )
    if not comments:
        raise ValueError("This Teams conversation has no renderable messages")
    output = Path(output_filename)
    render_comment_list(
        comments,
        output_filename=str(output),
        music_code=music_code,
        resolution_scale=resolution_scale,
    )
    return output
