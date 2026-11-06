"""Small Microsoft Graph client for Teams message collections."""

from collections.abc import Callable, Iterator
import base64
from datetime import datetime, timezone
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from email.utils import parsedate_to_datetime
from time import sleep
from typing import Any
from urllib.parse import quote, urlsplit

import requests
from PIL import Image, UnidentifiedImageError
from platformdirs import user_cache_path

from .links import TeamsTarget


_GRAPH_ROOT = "https://graph.microsoft.com/v1.0"


def is_sharepoint_url(url: str) -> bool:
    try:
        parsed = urlsplit(url)
        return (parsed.scheme == "https" and parsed.hostname is not None
                and parsed.hostname.endswith(".sharepoint.com")
                and parsed.username is None and parsed.password is None
                and parsed.port is None)
    except ValueError:
        return False


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

    def download_image(self, url: str) -> str | None:
        """Cache a Graph-hosted image for use as renderer evidence."""
        try:
            self._validate_url(url)
            parsed = urlsplit(url)
            if "/hostedContents/" not in parsed.path or not parsed.path.endswith("/$value"):
                return None
            response = self._session.get(
                url, headers={**self._headers, "Accept": "image/*"},
                timeout=30, allow_redirects=False, stream=True,
            )
            try:
                response.raise_for_status()
                if not 200 <= response.status_code < 300:
                    return None
                data = bytearray()
                for chunk in response.iter_content(65536):
                    data.extend(chunk)
                    if len(data) > 20 * 1024 * 1024:
                        return None
                with Image.open(BytesIO(data)) as image:
                    image.load()
                    image = image.convert("RGBA")
                directory = user_cache_path("objectionengine2-teams", appauthor=False) / "images"
                directory.mkdir(parents=True, exist_ok=True, mode=0o700)
                path = directory / f"{sha256(url.encode()).hexdigest()}.png"
                image.save(path, format="PNG")
                return str(path)
            finally:
                response.close()
        except (GraphError, requests.RequestException, OSError, ValueError, UnidentifiedImageError):
            return None

    def download_sharepoint_image(self, url: str) -> str | None:
        """Resolve a SharePoint sharing URL with Graph, then fetch its image bytes."""
        if not is_sharepoint_url(url):
            return None
        token = "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
        endpoint = f"{_GRAPH_ROOT}/shares/{token}/driveItem/content"
        try:
            response = self._session.get(endpoint, headers=self._headers, timeout=30, allow_redirects=False)
            try:
                if response.status_code != 302:
                    raise GraphError(f"Could not access SharePoint image (Graph HTTP {response.status_code}); Files.Read consent and file access may be required")
                location = response.headers.get("Location", "")
            finally:
                response.close()
            parsed = urlsplit(location)
            host = parsed.hostname or ""
            if (parsed.scheme != "https" or parsed.username is not None or parsed.password is not None
                or parsed.port is not None or not any(host.endswith("." + suffix) for suffix in ("sharepoint.com", "1drv.com", "onedrive.com"))):
                raise GraphError("Graph returned an unsupported SharePoint download host")
            response = self._session.get(location, headers={"Accept": "image/*"}, timeout=30, allow_redirects=False, stream=True)
            try:
                response.raise_for_status()
                if not 200 <= response.status_code < 300:
                    raise GraphError(f"SharePoint image download failed (HTTP {response.status_code})")
                data = bytearray()
                for chunk in response.iter_content(65536):
                    data.extend(chunk)
                    if len(data) > 20 * 1024 * 1024:
                        raise GraphError("SharePoint image exceeds the 20 MiB limit")
                with Image.open(BytesIO(data)) as image:
                    image.load()
                    image = image.convert("RGBA")
                directory = user_cache_path("objectionengine2-teams", appauthor=False) / "images"
                directory.mkdir(parents=True, exist_ok=True, mode=0o700)
                path = directory / f"{sha256(url.encode()).hexdigest()}.png"
                image.save(path, format="PNG")
                return str(path)
            finally:
                response.close()
        except (requests.RequestException, OSError, ValueError, UnidentifiedImageError) as exc:
            raise GraphError("Could not download or decode SharePoint image") from exc

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
