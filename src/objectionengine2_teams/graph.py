"""Small Microsoft Graph client for Teams message collections."""

from collections.abc import Callable, Iterator
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from time import sleep
from typing import Any
from urllib.parse import quote, urlsplit

import requests

from .links import TeamsTarget


_GRAPH_ROOT = "https://graph.microsoft.com/v1.0"


class GraphError(RuntimeError):
    """Microsoft Graph did not return a usable Teams message collection."""


class GraphClient:
    def __init__(
        self,
        access_token: str,
        *,
        session: requests.Session | None = None,
        wait: Callable[[float], None] = sleep,
        max_retries: int = 4,
    ) -> None:
        if not access_token:
            raise ValueError("A Microsoft Graph access token is required")
        self._session = session if session is not None else requests.Session()
        self._owns_session = session is None
        self._headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
        self._wait = wait
        self._max_retries = max_retries

    def close(self) -> None:
        if self._owns_session:
            self._session.close()

    def __enter__(self) -> "GraphClient":
        return self

    def __exit__(self, *_exc: object) -> None:
        self.close()

    @staticmethod
    def _validate_url(url: str) -> None:
        parsed = urlsplit(url)
        if (
            parsed.scheme != "https"
            or parsed.hostname != "graph.microsoft.com"
            or parsed.username is not None
            or parsed.password is not None
            or parsed.port is not None
            or not parsed.path.startswith("/v1.0/")
        ):
            raise GraphError("Microsoft Graph returned an unsafe pagination URL")

    @staticmethod
    def _retry_delay(response: requests.Response, attempt: int) -> float:
        retry_after = response.headers.get("Retry-After")
        if retry_after:
            try:
                return min(120.0, max(0.0, float(retry_after)))
            except ValueError:
                try:
                    deadline = parsedate_to_datetime(retry_after)
                    return min(120.0, max(0.0, (deadline - datetime.now(timezone.utc)).total_seconds()))
                except (TypeError, ValueError):
                    pass
        return min(30.0, float(2**attempt))

    def _get_json(self, url: str, scope: str) -> dict[str, Any]:
        self._validate_url(url)
        for attempt in range(self._max_retries + 1):
            try:
                response = self._session.get(
                    url, headers=self._headers, timeout=30, allow_redirects=False
                )
            except requests.RequestException as exc:
                raise GraphError(f"Could not reach Microsoft Graph: {exc}") from exc
            if response.status_code in (429, 503, 504) and attempt < self._max_retries:
                self._wait(self._retry_delay(response, attempt))
                continue
            if response.status_code == 403:
                raise GraphError(
                    f"Microsoft Graph denied access (403). The signed-in user needs access "
                    f"to this conversation, and client consent for {scope} is required."
                )
            if response.status_code == 401:
                raise GraphError("Microsoft Graph rejected the access token (401); sign in again")
            if response.status_code >= 400 or response.status_code < 200 or response.status_code >= 300:
                detail = ""
                try:
                    detail = response.json().get("error", {}).get("message", "")
                except (ValueError, AttributeError):
                    pass
                raise GraphError(f"Microsoft Graph returned HTTP {response.status_code}: {detail}".strip())
            try:
                data = response.json()
            except ValueError as exc:
                raise GraphError("Microsoft Graph returned invalid JSON") from exc
            if not isinstance(data, dict):
                raise GraphError("Microsoft Graph returned an unexpected response")
            return data
        raise AssertionError("unreachable")

    def _pages(self, url: str, scope: str) -> Iterator[dict[str, Any]]:
        seen: set[str] = set()
        while url:
            if url in seen:
                raise GraphError("Microsoft Graph returned a repeated pagination URL")
            seen.add(url)
            page = self._get_json(url, scope)
            items = page.get("value")
            if not isinstance(items, list):
                raise GraphError("Microsoft Graph omitted the message collection")
            for item in items:
                if not isinstance(item, dict):
                    raise GraphError("Microsoft Graph returned a malformed message")
                yield item
            next_url = page.get("@odata.nextLink")
            if next_url is not None and not isinstance(next_url, str):
                raise GraphError("Microsoft Graph returned a malformed pagination URL")
            url = next_url or ""

    def iter_messages(
        self, target: TeamsTarget, *, end_message_id: str | None = None
    ) -> Iterator[dict[str, Any]]:
        """Yield chat messages, one channel thread, or a complete channel."""
        if target.kind == "chat" and target.chat_id:
            chat = quote(target.chat_id, safe="")
            url = f"{_GRAPH_ROOT}/chats/{chat}/messages?$top=50&$orderby=createdDateTime desc"
            found_start = False
            found_end = end_message_id is None
            for message in self._pages(url, "Chat.Read"):
                yield message
                if str(message.get("id")) == target.start_message_id:
                    found_start = True
                if str(message.get("id")) == end_message_id:
                    found_end = True
                if target.start_message_id and found_start and found_end:
                    break
            return
        if target.kind == "channel" and target.team_id and target.channel_id:
            team = quote(target.team_id, safe="")
            channel = quote(target.channel_id, safe="")
            root = f"{_GRAPH_ROOT}/teams/{team}/channels/{channel}/messages"
            if target.start_message_id:
                parent_id = target.parent_message_id or target.start_message_id
                thread_url = f"{root}/{quote(parent_id, safe='')}"
                parent = self._get_json(thread_url, "ChannelMessage.Read.All")
                if str(parent.get("id")) != parent_id:
                    raise GraphError("Microsoft Graph returned a different channel thread")
                yield parent
                yield from self._pages(
                    f"{thread_url}/replies?$top=50", "ChannelMessage.Read.All"
                )
                return
            for message in self._pages(f"{root}?$top=50", "ChannelMessage.Read.All"):
                yield message
                message_id = message.get("id")
                if not isinstance(message_id, str) or not message_id:
                    raise GraphError("A channel message is missing its ID")
                reply_url = f"{root}/{quote(message_id, safe='')}/replies?$top=50"
                yield from self._pages(reply_url, "ChannelMessage.Read.All")
            return
        raise ValueError("Incomplete Teams target")
