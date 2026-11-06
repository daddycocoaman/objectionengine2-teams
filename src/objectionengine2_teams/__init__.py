"""Use Microsoft Graph messages as objectionengine2 dialogue."""

import warnings

# pydub emits SyntaxWarnings while its modules are compiled on newer Python.
with warnings.catch_warnings():
    warnings.simplefilter("ignore", SyntaxWarning)
    import pydub  # noqa: F401

from .auth import OFFICE_CLIENT_ID, AuthenticationError, acquire_access_token
from .convert import messages_to_comments
from .graph import GraphClient, GraphError
from .links import TeamsTarget, parse_teams_link
from .render import fetch_comments, render_teams_link

__all__ = [
    "OFFICE_CLIENT_ID",
    "AuthenticationError",
    "GraphClient",
    "GraphError",
    "TeamsTarget",
    "acquire_access_token",
    "fetch_comments",
    "messages_to_comments",
    "parse_teams_link",
    "render_teams_link",
]
