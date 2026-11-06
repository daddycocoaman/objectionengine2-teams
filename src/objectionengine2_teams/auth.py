"""Delegated Microsoft Graph authentication with a persistent Azure Identity cache."""

import hashlib
import os
import sys
import tempfile
from pathlib import Path

from azure.core.exceptions import AzureError
from azure.identity import (
    AuthenticationRecord,
    DeviceCodeCredential,
    InteractiveBrowserCredential,
    SharedTokenCacheCredential,
    TokenCachePersistenceOptions,
)
from azure.identity.broker import InteractiveBrowserBrokerCredential
from msal import PublicClientApplication

OFFICE_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
BROWSER_REDIRECT_URI = "http://localhost:8400"
_SCOPES = {
    "chat": "https://graph.microsoft.com/Chat.ReadWrite",
    "channel": "https://graph.microsoft.com/ChannelMessage.Read.All",
    "files": "https://graph.microsoft.com/Files.Read",
}


class AuthenticationError(RuntimeError):
    """Microsoft Entra ID did not issue a usable Graph token."""


def _state_dir() -> Path:
    """Return the per-user directory for non-secret authentication records."""
    if os.name == "nt":
        root = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    elif sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support"
    else:
        root = Path(
            os.environ.get("XDG_STATE_HOME") or Path.home() / ".local" / "state"
        )
    return root / "objectionengine2-teams"


def _cache_key(client_id: str, tenant_id: str) -> str:
    return hashlib.sha256(
        f"{client_id.lower()}\0{tenant_id.lower()}".encode()
    ).hexdigest()[:24]


def _load_record(path: Path, client_id: str) -> AuthenticationRecord | None:
    try:
        serialized = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return None
    except OSError as exc:
        raise AuthenticationError(
            f"Could not read the cached account record at {path}: {exc}"
        ) from exc
    try:
        record = AuthenticationRecord.deserialize(serialized)
    except (TypeError, ValueError, KeyError) as exc:
        raise AuthenticationError(
            f"The cached account record at {path} is invalid. Delete it and sign in again."
        ) from exc
    return record if record.client_id.lower() == client_id.lower() else None


def _save_record(path: Path, record: AuthenticationRecord) -> None:
    """Atomically save account identity; Azure Identity stores tokens separately."""
    temporary_path: Path | None = None
    try:
        path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=path.parent,
            prefix=".account-",
            delete=False,
        ) as temporary:
            temporary_path = Path(temporary.name)
            temporary.write(record.serialize())
        os.replace(temporary_path, path)
    except OSError as exc:
        raise AuthenticationError(
            f"Could not save the cached account record at {path}: {exc}"
        ) from exc
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def acquire_access_token(
    kind: str,
    *,
    tenant_id: str | None = None,
    client_id: str = OFFICE_CLIENT_ID,
    auth_method: str = "broker",
    redirect_uri: str = BROWSER_REDIRECT_URI,
    record_dir: str | Path | None = None,
) -> str:
    """Return a Graph token, reusing the user's cached account when possible.

    Broker sign-in is the default. Set ``auth_method='browser'`` or
    ``auth_method='device-code'`` to use another interactive method. All
    methods use the Microsoft Office public client ID unless overridden.
    """
    if kind not in _SCOPES:
        raise ValueError("kind must be 'chat', 'channel', or 'files'")
    if auth_method not in ("broker", "browser", "device-code", "cache"):
        raise ValueError(
            "auth_method must be 'broker', 'browser', 'device-code', or 'cache'"
        )

    tenant = tenant_id or "organizations"
    key = _cache_key(client_id, tenant)
    record_path = (
        Path(record_dir) if record_dir is not None else _state_dir()
    ) / f"{key}.json"
    record = _load_record(record_path, client_id)
    options = TokenCachePersistenceOptions(
        name=f"objectionengine2-teams-{key}", allow_unencrypted_storage=True
    )
    kwargs = {
        "client_id": client_id,
        "tenant_id": tenant,
        "cache_persistence_options": options,
    }
    if record is not None:
        kwargs["authentication_record"] = record
    scope = _SCOPES[kind]
    credential = None
    try:
        if auth_method == "broker":
            credential = InteractiveBrowserBrokerCredential(
                parent_window_handle=PublicClientApplication.CONSOLE_WINDOW_HANDLE,
                use_default_broker_account=True,
                **kwargs,
            )
        elif auth_method == "cache":
            credential = SharedTokenCacheCredential(**kwargs)
        elif auth_method == "browser":
            credential = InteractiveBrowserCredential(
                redirect_uri=redirect_uri, **kwargs
            )
        elif auth_method == "device-code":
            credential = DeviceCodeCredential(**kwargs)
        if record is None and auth_method in ("browser", "device-code"):
            record = credential.authenticate(scopes=[scope])
            _save_record(record_path, record)
        return credential.get_token(scope).token
    except AzureError as exc:
        raise AuthenticationError(
            f"Microsoft sign-in could not grant {scope} to client {client_id}: {exc}. "
            "Your tenant may require administrator consent for this Graph permission."
        ) from exc
    finally:
        if credential is not None:
            credential.close()
