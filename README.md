# objectionengine2-teams

Render a Microsoft Teams chat or channel as an Objection Engine 2 video. This is
a separate Python 3.10–3.14 package using Microsoft Graph delegated access.

## Install

Install Python 3.10–3.14, uv, FFmpeg and `ffprobe`, then install the command
from PyPI:

```bash
uv tool install --torch-backend cpu objectionengine2-teams
```

The install includes `objectionengine2` and its sentiment dependencies. The
first render downloads the Hugging Face sentiment model.

Objection Engine saves its
asset pack and sentiment model under the platformdirs user cache for
`objectionengine2`. The first render populates the cache. See the cache path
with `uv run --no-project --torch-backend cpu --with objectionengine2-teams python -c 'from objectionengine2.cache import cache_path; print(cache_path())'`.
Run the installed command:

```bash
mkdir -p outputs
objectionengine2-teams \
  'https://teams.microsoft.com/l/chat/19:example@thread.v2/conversations' \
  -o outputs/teams.mp4
```

Replace the example URL with a real link copied from a chat or channel. A
successful render saves its settings next to the video as `teams.mp4.json`.
The file includes the rendered conversation, Teams links and authentication
settings, but no tokens. Evidence paths refer to files in the local image cache.
Scene choices are randomized, so the settings do not guarantee an identical video.
A channel link renders the channel; a channel **message link** renders only its
thread. Its `parentMessageId` identifies the thread root, while the message ID
in the URL path is the first message rendered. A chat message link likewise
starts at the linked message. You can bound the render with another message
link in the same chat or channel thread:

```bash
objectionengine2-teams \
  '<start message link>' --end-link '<end message link>' -o outputs/teams.mp4
```

Both endpoints are included. `--max-messages 20` selects the first 20 rendered
messages after a start link, or the newest 20 for an unbounded chat/channel
link. The CLI sorts them by creation time, skips empty messages, and excludes
thread titles from dialogue. Rich
reactions, deleted messages, and system events are not represented. The first
inline image in each message is displayed as evidence when Microsoft Graph
provides it as hosted content. SharePoint image links and image attachments are
also supported when Graph can resolve their sharing URL and the signed-in user
can access the file. Resolving one requests the delegated `Files.Read` permission,
which may require additional consent. Other attachments become a short placeholder
when there is no message text. Images are cached in the user cache for
`objectionengine2-teams`; remove its `images` directory to delete them. Images
that cannot be downloaded are omitted.

## Sign-in and permissions

The default public client ID is the requested Microsoft Office ID,
`d3590ed6-52b3-4102-aeff-aad2292ab01c`. Azure Identity uses
`InteractiveBrowserBrokerCredential` by default. The CLI requests Microsoft
Graph **delegated** permissions for the signed-in user:

| Used for | Delegated Microsoft Graph permission | When requested |
| --- | --- | --- |
| Reading chat messages | `Chat.ReadWrite` | For a chat link |
| Reading channel messages | `ChannelMessage.Read.All` | For a channel link |
| Downloading SharePoint images | `Files.Read` | When a supported SharePoint image is encountered |

The signed-in user must have access to the conversation and, for SharePoint
images, the referenced files. Teams-hosted inline images use the conversation
permission; `Files.Read` is requested separately only for SharePoint images.
The CLI may prompt for additional consent when it first encounters one. If file
access or consent fails, it warns and omits the image while continuing the render.
With `--auth-method cache`, the required token must already be available in the
cache because that mode cannot prompt.

Microsoft Entra may reject the default client or require administrator consent,
especially for channel messages. The client ID does not itself guarantee a
grant. If the Office client cannot obtain the requested scopes, register a
public client application in your tenant with the delegated permissions needed
for your links and pass its ID through `--client-id`. Broker sign-in uses the
platform's broker redirect URI; on
Windows and WSL, that is
`ms-appx-web://Microsoft.AAD.BrokerPlugin/<client-id>`. The application must
allow that redirect. The `--redirect-uri` option applies only to explicit
`--auth-method browser` sign-in, whose default callback is
`http://localhost:8400`.

On Windows, broker sign-in uses Web Account Manager. On WSL 2.4.13 or newer,
it can use the Windows broker as well. On other Linux setups it may open a
browser if a supported broker is unavailable. To select another method, use:

```bash
objectionengine2-teams '<Teams message link>' \
  --auth-method browser -o outputs/teams.mp4
objectionengine2-teams '<Teams message link>' \
  --auth-method device-code -o outputs/teams.mp4
```

Azure Identity persists tokens in a per-user cache and reuses them on later
runs. Broker sign-in first tries the default broker account silently and then
offers interactive sign-in if needed. Use `--auth-method cache` to read an
already cached token without an interactive fallback. Azure Identity encrypts
the persistent cache when the platform supports it; on platforms without
user-data encryption (including some WSL installations), it may store tokens
unencrypted. Keep your user profile private. Browser and device-code sign-in
also save a separate non-secret account record under the user state directory
(`~/.local/state/objectionengine2-teams/` on Linux unless `XDG_STATE_HOME` is
set). The methods use the same named persistent token cache.

## Sentiment and objections

Teams messages have no explicit `Comment.score`. Teams renders use the core
Hugging Face sentiment model by default. The normal command is enough:

```bash
objectionengine2-teams \
  '<Teams message link>' -o outputs/teams.mp4
```

Set `oe_sentiment_model=neutral` or `oe_bypass_sentiment=1` to use neutral
sentiment for a render. An explicit `oe_sentiment_model` value is respected.
A strongly negative classification
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
`render_teams_link` enables Hugging Face sentiment by default for its render.
Both calls default to broker sign-in. Pass `auth_method="browser"`,
`auth_method="device-code"`, or `auth_method="cache"` to select another
method. Later calls reuse cached tokens when valid.
To inspect messages and render without fetching them twice, use the comments
with `objectionengine2.renderer.render_comment_list` directly.

## Tests

For development, clone the repository and install its locked dependencies with
uv:

```bash
cd ~/development/objectionengine2-teams
uv sync --locked
uv run --locked pytest -q
uv build --no-sources
```

The default tests use fake Graph responses and Azure Identity credentials. They
never access a tenant, download assets or models, or render private Teams content. To exercise
the complete video encoder with local fixture messages, run:

```bash
uv run --locked pytest -q -m asset_integration
```

That test calls `objectionengine2.utils.ensure_assets_are_available()` first,
then renders with default Hugging Face sentiment. It populates the platform
cache with assets and the model when necessary. A live Teams render still
requires a link, sign-in, and Graph consent; it cannot be validated by the
offline suite.
