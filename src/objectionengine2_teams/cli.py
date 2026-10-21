"""Command-line interface for rendering a Teams conversation."""

import argparse

from .auth import BROWSER_REDIRECT_URI, OFFICE_CLIENT_ID, AuthenticationError
from .graph import GraphError
from .render import render_teams_link


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="objectionengine2-teams",
        description="Render a Microsoft Teams chat or channel as an Objection Engine video.",
    )
    parser.add_argument("link", help="A Teams chat, channel, or starting message link")
    parser.add_argument("-o", "--output", default="teams.mp4", help="Output MP4 path")
    parser.add_argument(
        "--end-link", help="Inclusive ending message link in the same conversation"
    )
    parser.add_argument(
        "--max-messages",
        type=int,
        help="At most N messages from the start link, or the newest N for a whole chat/channel",
    )
    parser.add_argument(
        "--client-id", default=OFFICE_CLIENT_ID, help="Public client application ID"
    )
    parser.add_argument(
        "--auth-method",
        choices=("broker", "browser", "device-code", "cache"),
        default="broker",
        help="Sign-in method (default: broker)",
    )
    parser.add_argument(
        "--redirect-uri",
        default=BROWSER_REDIRECT_URI,
        help="Registered redirect URI for browser sign-in",
    )
    parser.add_argument(
        "--music-code", default="pwr", help="Objection Engine music code"
    )
    parser.add_argument("--resolution-scale", type=float, default=1.0)
    args = parser.parse_args(argv)
    if args.max_messages is not None and args.max_messages < 1:
        parser.error("--max-messages must be positive")
    try:
        output = render_teams_link(
            args.link,
            args.output,
            end_link=args.end_link,
            max_messages=args.max_messages,
            client_id=args.client_id,
            auth_method=args.auth_method,
            redirect_uri=args.redirect_uri,
            music_code=args.music_code,
            resolution_scale=args.resolution_scale,
        )
    except (AuthenticationError, GraphError, ValueError) as exc:
        parser.exit(1, f"{parser.prog}: {exc}\n")
    print(f"Saved {output}")
    return 0
