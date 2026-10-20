"""Command-line interface for rendering a Teams conversation."""

import argparse

from .auth import AuthenticationError, OFFICE_CLIENT_ID
from .graph import GraphError
from .render import render_teams_link


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="objectionengine2-teams",
        description="Render a Microsoft Teams chat or channel as an Objection Engine video.",
    )
    parser.add_argument("link", help="A Teams chat, channel, or message link")
    parser.add_argument("-o", "--output", default="teams.mp4", help="Output MP4 path")
    parser.add_argument(
        "--max-messages", type=int, help="Render only the newest N messages, ordered chronologically"
    )
    parser.add_argument("--client-id", default=OFFICE_CLIENT_ID, help="Public client application ID")
    parser.add_argument("--music-code", default="pwr", help="Objection Engine music code")
    parser.add_argument("--resolution-scale", type=float, default=1.0)
    args = parser.parse_args(argv)
    if args.max_messages is not None and args.max_messages < 1:
        parser.error("--max-messages must be positive")
    try:
        output = render_teams_link(
            args.link,
            args.output,
            max_messages=args.max_messages,
            client_id=args.client_id,
            music_code=args.music_code,
            resolution_scale=args.resolution_scale,
        )
    except (AuthenticationError, GraphError, ValueError) as exc:
        parser.exit(1, f"{parser.prog}: {exc}\n")
    print(f"Saved {output}")
    return 0
