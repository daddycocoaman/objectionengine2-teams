"""Explicit opt-in end-to-end render using the core asset pack."""

from pathlib import Path
import json
import subprocess

import pytest

from objectionengine2.utils import ensure_assets_are_available
from objectionengine2_teams import render as render_module


pytestmark = pytest.mark.asset_integration


def test_renders_mp4_after_ensuring_assets(monkeypatch, tmp_path):
    core_checkout = Path(__file__).resolve().parents[2] / "objectionengine2"
    if not core_checkout.is_dir():
        pytest.skip("Adjacent objectionengine2 checkout is required")
    monkeypatch.chdir(core_checkout)
    ensure_assets_are_available()

    monkeypatch.setattr(render_module, "acquire_access_token", lambda *args, **kwargs: "fixture-token")

    class FakeGraph:
        def __init__(self, token):
            assert token == "fixture-token"

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return None

        def iter_messages(self, target):
            assert target.kind == "chat"
            return iter(
                [
                    {
                        "id": "1",
                        "createdDateTime": "2026-01-01T00:00:00Z",
                        "messageType": "message",
                        "from": {"user": {"id": "alex", "displayName": "Alex"}},
                        "body": {"contentType": "html", "content": "<p>Objection!</p>"},
                    }
                ]
            )

    monkeypatch.setattr(render_module, "GraphClient", FakeGraph)
    output = tmp_path / "teams.mp4"
    render_module.render_teams_link(
        "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations", output
    )
    assert output.is_file()
    assert output.stat().st_size > 1000
    result = subprocess.run(
        ["ffprobe", "-v", "error", "-show_entries", "stream=codec_type:format=duration", "-of", "json", str(output)],
        capture_output=True,
        text=True,
        check=True,
    )
    metadata = json.loads(result.stdout)
    assert {stream["codec_type"] for stream in metadata["streams"]} == {"video", "audio"}
    assert float(metadata["format"]["duration"]) > 0
