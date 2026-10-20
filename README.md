# objectionengine2-teams

Render a Microsoft Teams chat or channel as an Objection Engine 2 video. This is
a separate Python 3.10–3.14 package using Microsoft Graph delegated access.

## Install

Clone `objectionengine2` and `objectionengine2-teams` next to each other. This
source checkout uses the adjacent `../objectionengine2` package with uv. Built
packages depend on the pinned public core Git commit until the core package is
published:

```bash
cd ~/development/objectionengine2-teams
uv sync --locked
```

Install FFmpeg and `ffprobe` for video rendering. Objection Engine downloads
its asset pack into the current directory's `assets/` folder on first render.
To reuse the pack from the adjacent core checkout, run from that checkout:

```bash
cd ~/development/objectionengine2
uv run --project ../objectionengine2-teams objectionengine2-teams \
  'https://teams.microsoft.com/l/chat/19:example@thread.v2/conversations' \
  -o outputs/teams.mp4
```

Replace the example URL with a real link copied from a chat or channel. A
message link selects the whole conversation. Channel renders include replies.
Use `--max-messages 20` to render only the newest 20 text messages. The CLI
sorts them by creation time before rendering. Rich attachments, reactions,
deleted messages, and system events are not represented; an attachment-only
message becomes a short placeholder.

## Sign-in and permissions

The default public client ID is the requested Microsoft Office ID,
`d3590ed6-52b3-4102-aeff-aad2292ab01c`. The CLI starts MSAL device-code
sign-in and requests only the delegated permission needed for the link:

| Link | Delegated Microsoft Graph permission |
| --- | --- |
| Chat | `Chat.Read` |
| Channel | `ChannelMessage.Read.All` |

The signed-in user must have access to the conversation. Microsoft Entra may
reject this client or require administrator consent, especially for the channel
permission. The client ID does not itself guarantee a grant. If the Office
client cannot obtain the requested scope, register a public client application
in your tenant with the Graph delegated permission and pass its ID through
`--client-id`. The library does not store tokens on disk.

## Python API

```python
from objectionengine2_teams import fetch_comments, render_teams_link

comments = fetch_comments(teams_url, max_messages=20)
render_teams_link(teams_url, "outputs/teams.mp4", max_messages=20)
```

`fetch_comments` returns `objectionengine2.beans.comment.Comment` objects.
Both calls use the device-code flow; calling them separately signs in twice.
To inspect messages and render in a single sign-in, use the comments with
`objectionengine2.renderer.render_comment_list` directly.

## Tests

```bash
cd ~/development/objectionengine2-teams
uv run --locked pytest -q
uv build --no-sources
```

The tests use fake Graph responses and MSAL clients. They never access a
tenant, download the asset pack, or render private Teams content. To exercise
the complete video encoder with local fixture messages, run:

```bash
uv run --locked pytest -q -m asset_integration
```

That test calls `objectionengine2.utils.ensure_assets_are_available()` first,
which downloads the asset pack when necessary. A live Teams render still
requires a link, sign-in, and Graph consent; it cannot be validated by the
offline suite.
