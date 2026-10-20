"""Delegated Microsoft Graph authentication through MSAL device code flow."""

from collections.abc import Callable

OFFICE_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
_SCOPES = {"chat": ["Chat.Read"], "channel": ["ChannelMessage.Read.All"]}


class AuthenticationError(RuntimeError):
    """Microsoft Entra ID did not issue a usable Graph token."""


def acquire_access_token(
    kind: str,
    *,
    tenant_id: str | None = None,
    client_id: str = OFFICE_CLIENT_ID,
    show_device_code: Callable[[str], None] = print,
    app_factory: Callable[..., object] | None = None,
) -> str:
    """Sign in with the requested delegated read permission.

    The default client ID is the Microsoft Office ID requested by this project.
    A tenant may deny that client or require administrator consent, especially
    for ChannelMessage.Read.All. Callers can supply their own public-client ID.
    """
    if kind not in _SCOPES:
        raise ValueError("kind must be 'chat' or 'channel'")
    if app_factory is None:
        from msal import PublicClientApplication

        app_factory = PublicClientApplication

    authority = f"https://login.microsoftonline.com/{tenant_id or 'organizations'}"
    app = app_factory(client_id=client_id, authority=authority)
    scopes = _SCOPES[kind]

    for account in app.get_accounts():
        result = app.acquire_token_silent(scopes, account=account)
        if result and result.get("access_token"):
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        detail = flow.get("error_description") or flow.get("error") or "unknown error"
        raise AuthenticationError(f"Unable to start Microsoft sign-in: {detail}")
    show_device_code(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        detail = result.get("error_description") or result.get("error") or "unknown error"
        raise AuthenticationError(
            f"Microsoft sign-in did not grant {scopes[0]} to client {client_id}: {detail}. "
            "Ask your tenant administrator to grant the permission, or use --client-id "
            "with a public app registration that has it."
        )
    return result["access_token"]
