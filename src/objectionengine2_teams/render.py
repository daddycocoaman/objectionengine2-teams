"""Public Teams-to-video API."""

from pathlib import Path

from objectionengine2.beans.comment import Comment
from objectionengine2.renderer import render_comment_list

from .auth import OFFICE_CLIENT_ID, acquire_access_token
from .convert import messages_to_comments
from .graph import GraphClient
from .links import parse_teams_link


def fetch_comments(
    link: str,
    *,
    max_messages: int | None = None,
    client_id: str = OFFICE_CLIENT_ID,
) -> list[Comment]:
    """Read a Teams chat or channel and return chronological dialogue comments."""
    target = parse_teams_link(link)
    token = acquire_access_token(target.kind, tenant_id=target.tenant_id, client_id=client_id)
    with GraphClient(token) as graph:
        comments = messages_to_comments(graph.iter_messages(target), max_messages=max_messages)
    return comments


def render_teams_link(
    link: str,
    output_filename: str | Path = "teams.mp4",
    *,
    max_messages: int | None = None,
    client_id: str = OFFICE_CLIENT_ID,
    music_code: str = "pwr",
    resolution_scale: float = 1,
) -> Path:
    """Render all readable text messages from a Teams conversation to MP4."""
    comments = fetch_comments(link, max_messages=max_messages, client_id=client_id)
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
