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
channel link renders the channel; a channel **message link** renders only its
thread. Its `parentMessageId` identifies the thread root, while the message ID
in the URL path is the first message rendered. A chat message link likewise
starts at the linked message. You can bound the render with another message
link in the same chat or channel thread:

```bash
uv run --project ../objectionengine2-teams --locked objectionengine2-teams \
  '<start message link>' --end-link '<end message link>' -o outputs/teams.mp4
```

Both endpoints are included. `--max-messages 20` selects the first 20 rendered
messages after a start link, or the newest 20 for an unbounded chat/channel
link. The CLI sorts them by creation time. Rich attachments, reactions,
deleted messages, and system events are not represented; an attachment-only
message becomes a short placeholder.

## Sign-in and permissions

The default public client ID is the requested Microsoft Office ID,
`d3590ed6-52b3-4102-aeff-aad2292ab01c`. Azure Identity opens a browser for
sign-in by default and requests only the delegated permission needed for the link:

| Link | Delegated Microsoft Graph permission |
| --- | --- |
| Chat | `Chat.Read` |
| Channel | `ChannelMessage.Read.All` |

The signed-in user must have access to the conversation. Microsoft Entra may
reject this client or require administrator consent, especially for the channel
permission. The client ID does not itself guarantee a grant. If the Office
client cannot obtain the requested scope, register a public client application
in your tenant with the Graph delegated permission and pass its ID through
`--client-id`. Browser sign-in uses `http://localhost:8400` as its callback;
the URI must be registered for the selected client. Use `--redirect-uri` if your
app registration uses another local callback.

For WSL, SSH, or another environment where browser sign-in is inconvenient,
use device-code sign-in:

```bash
uv run --locked objectionengine2-teams '<Teams message link>' \
  --auth-method device-code -o outputs/teams.mp4
```

Azure Identity persists tokens in a per-user cache and reuses them on later
runs. It encrypts the cache when the platform supports it; on platforms without
user-data encryption (including some WSL installations), it may store tokens
unencrypted. Keep your user profile private. A separate non-secret account
record is saved under the user state directory
(`~/.local/state/objectionengine2-teams/` on Linux unless `XDG_STATE_HOME` is
set). The two sign-in methods share the same cache and account record.

## Sentiment and objections

Teams messages have no explicit `Comment.score`, so the core renderer treats
them as neutral by default. Enable its Hugging Face sentiment model with the
Teams package's optional `sentiment` extra and `oe_sentiment_model=hf`. From the
adjacent core checkout, where `assets/` is available:

```bash
env -u oe_bypass_sentiment oe_sentiment_model=hf \
  uv run --project ../objectionengine2-teams --locked --extra sentiment \
  objectionengine2-teams '<Teams message link>' -o outputs/teams.mp4
```

The first run downloads the sentiment model. A strongly negative classification
(confidence above 0.5) has roughly a 70% chance of triggering an objection;
a strongly positive one has roughly a 30% chance. An objection is therefore
possible, not guaranteed, for any particular message. If you need to choose
the sentiment yourself, call `fetch_comments`, set `comment.score` to a negative
or positive value above 0.5 in magnitude, then pass the comments to
`objectionengine2.renderer.render_comment_list`. The objection still has a
random chance of occurring.

## Python API

```python
from objectionengine2_teams import fetch_comments, render_teams_link

comments = fetch_comments(start_message_url, end_link=end_message_url)
render_teams_link(start_message_url, "outputs/teams.mp4", end_link=end_message_url)
```

`fetch_comments` returns `objectionengine2.beans.comment.Comment` objects.
Both calls default to browser sign-in. Pass `auth_method="device-code"` to use
device-code sign-in. Later calls reuse the cached account and token when valid.
To inspect messages and render without fetching them twice, use the comments
with `objectionengine2.renderer.render_comment_list` directly.

## Tests

```bash
cd ~/development/objectionengine2-teams
uv run --locked pytest -q
uv build --no-sources
```

The tests use fake Graph responses and Azure Identity credentials. They never access a
tenant, download the asset pack, or render private Teams content. To exercise
the complete video encoder with local fixture messages, run:

```bash
uv run --locked pytest -q -m asset_integration
```

That test calls `objectionengine2.utils.ensure_assets_are_available()` first,
which downloads the asset pack when necessary. A live Teams render still
requires a link, sign-in, and Graph consent; it cannot be validated by the
offline suite.
