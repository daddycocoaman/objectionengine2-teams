"""Map Microsoft Graph chatMessage bodies to objectionengine2 comments."""

from html.parser import HTMLParser
import re
from typing import Any, Iterable

from objectionengine2.beans.comment import Comment


class _TeamsHTMLParser(HTMLParser):
    _BREAKS = {"br", "div", "p", "li", "blockquote", "tr"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.hidden_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "style"}:
            self.hidden_depth += 1
        elif self.hidden_depth == 0 and tag in self._BREAKS:
            self.parts.append("\n")
        elif self.hidden_depth == 0 and tag == "img":
            alt = dict(attrs).get("alt")
            if alt:
                self.parts.append(alt)

    def handle_endtag(self, tag: str) -> None:
        if tag in {"script", "style"} and self.hidden_depth:
            self.hidden_depth -= 1
        elif self.hidden_depth == 0 and tag in self._BREAKS:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if self.hidden_depth == 0:
            self.parts.append(data)


def _body_text(message: dict[str, Any]) -> str:
    body = message.get("body") or {}
    if not isinstance(body, dict):
        body = {}
    content = body.get("content") or ""
    if not isinstance(content, str):
        content = ""
    if str(body.get("contentType", "")).lower() == "html":
        parser = _TeamsHTMLParser()
        parser.feed(content)
        content = "".join(parser.parts)
    lines = [re.sub(r"\s+", " ", line).strip() for line in content.splitlines()]
    text = "\n".join(line for line in lines if line)
    if text:
        return text
    attachments = message.get("attachments") or []
    if isinstance(attachments, list) and attachments:
        names = [attachment.get("name") for attachment in attachments if isinstance(attachment, dict)]
        names = [name for name in names if isinstance(name, str) and name.strip()]
        return "Shared an attachment" + (": " + ", ".join(names) if names else "")
    return ""


def _speaker(message: dict[str, Any]) -> tuple[str, str]:
    actor = message.get("from") or {}
    if not isinstance(actor, dict):
        actor = {}
    for key in ("user", "application", "device"):
        value = actor.get(key)
        if isinstance(value, dict):
            name = value.get("displayName") or "Unknown speaker"
            identity = value.get("id") or name
            return str(identity), str(name)
    return "unknown", "Unknown speaker"


def messages_to_comments(
    messages: Iterable[dict[str, Any]],
    *,
    start_message_id: str | None = None,
    end_message_id: str | None = None,
    max_messages: int | None = None,
) -> list[Comment]:
    """Convert an inclusive message range to chronological dialogue."""
    if max_messages is not None and max_messages < 1:
        raise ValueError("max_messages must be positive")
    ordered: list[dict[str, Any]] = []
    seen: set[str] = set()
    for message in messages:
        message_id = str(message.get("id") or "")
        if message_id and message_id in seen:
            continue
        seen.add(message_id)
        ordered.append(message)
    ordered.sort(key=lambda message: (str(message.get("createdDateTime") or ""), str(message.get("id") or "")))

    ids = [str(message.get("id") or "") for message in ordered]
    if start_message_id is not None and start_message_id not in ids:
        raise ValueError(f"Start message {start_message_id} was not found in this conversation")
    if end_message_id is not None and end_message_id not in ids:
        raise ValueError(f"End message {end_message_id} was not found in this conversation")
    if not ordered:
        return []
    start = ids.index(start_message_id) if start_message_id is not None else 0
    end = ids.index(end_message_id) if end_message_id is not None else len(ordered) - 1
    if end < start:
        raise ValueError("The end message occurs before the start message")

    comments: list[Comment] = []
    for message in ordered[start : end + 1]:
        if message.get("messageType", "message") != "message" or message.get("deletedDateTime"):
            continue
        text = _body_text(message)
        if not text:
            continue
        user_id, user_name = _speaker(message)
        comments.append(Comment(user_id=user_id, user_name=user_name, text_content=text))
    if max_messages is not None:
        comments = comments[:max_messages] if start_message_id else comments[-max_messages:]
    return comments
