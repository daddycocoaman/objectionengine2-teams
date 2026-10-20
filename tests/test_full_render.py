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

        def iter_messages(self, target, *, end_message_id=None):
            assert target.kind == "channel"
            assert target.parent_message_id == "100"
            assert target.start_message_id == "101"
            assert end_message_id == "102"
            return iter(
                [
                    {
                        "id": str(message_id),
                        "createdDateTime": f"2026-01-0{message_id - 99}T00:00:00Z",
                        "messageType": "message",
                        "from": {"user": {"id": "alex", "displayName": "Alex"}},
                        "body": {"contentType": "html", "content": f"<p>{text}</p>"},
                    }
                    for message_id, text in (
                        (100, "Earlier"), (101, "Objection!"),
                        (102, "Hold it!"), (103, "After"),
                    )
                ]
            )

    monkeypatch.setattr(render_module, "GraphClient", FakeGraph)
    output = tmp_path / "teams.mp4"
    group_id = "72602e12-78ac-474c-99d6-f619710353a9"
    base = "https://teams.microsoft.com/l/message/19:abc@thread.tacv2/"
    render_module.render_teams_link(
        f"{base}101?groupId={group_id}&parentMessageId=100", output,
        end_link=f"{base}102?groupId={group_id}&parentMessageId=100",
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
