import pytest

from objectionengine2.beans.comment import Comment
from objectionengine2_teams.convert import messages_to_comments
from objectionengine2_teams import render as render_module


def _message(message_id, created, content, *, user="Alex", **extra):
    return {
        "id": message_id,
        "createdDateTime": created,
        "messageType": "message",
        "from": {"user": {"id": user.lower(), "displayName": user}},
        "body": {"contentType": "html", "content": content},
        **extra,
    }


def test_message_conversion_orders_and_cleans_html():
    messages = [
        _message("2", "2026-01-02T00:00:00Z", "<p>Second &amp; third<br>line</p>"),
        _message("1", "2026-01-01T00:00:00Z", "<div>First</div>", user="Sam"),
        _message("2", "2026-01-02T00:00:00Z", "duplicate"),
        _message("3", "2026-01-03T00:00:00Z", "<p>system</p>", messageType="systemEventMessage"),
        _message("4", "2026-01-04T00:00:00Z", "deleted", deletedDateTime="now"),
    ]
    comments = messages_to_comments(messages)
    assert [comment.text_content for comment in comments] == ["First", "Second & third\nline"]
    assert comments[0].user_id == "sam"
    assert comments[1].user_name == "Alex"
    assert [comment.text_content for comment in messages_to_comments(messages, max_messages=1)] == [
        "Second & third\nline"
    ]


def test_attachment_only_message_uses_placeholder():
    message = _message("1", "2026-01-01T00:00:00Z", "", attachments=[{"name": "chart.png"}])
    message.pop("body")
    assert messages_to_comments([message])[0].text_content == "Shared an attachment: chart.png"


def test_fetch_comments_connects_link_auth_and_graph(monkeypatch):
    auth_calls = []
    monkeypatch.setattr(
        render_module,
        "acquire_access_token",
        lambda *args, **kwargs: auth_calls.append((args, kwargs)) or "token",
    )

    class FakeGraph:
        def __init__(self, token):
            assert token == "token"

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return None

        def iter_messages(self, target):
            assert target.kind == "chat"
            assert target.chat_id == "19:abc@thread.v2"
            return iter([_message("1", "2026-01-01T00:00:00Z", "<p>Objection!</p>")])

    monkeypatch.setattr(render_module, "GraphClient", FakeGraph)
    comments = render_module.fetch_comments(
        "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations"
    )
    assert [comment.text_content for comment in comments] == ["Objection!"]
    assert auth_calls[0][0] == ("chat",)


def test_render_bridge_passes_comments_to_core(monkeypatch, tmp_path):
    comments = [Comment(user_id="alex", user_name="Alex", text_content="Objection!")]
    calls = []
    monkeypatch.setattr(render_module, "fetch_comments", lambda *args, **kwargs: comments)
    monkeypatch.setattr(render_module, "render_comment_list", lambda *args, **kwargs: calls.append((args, kwargs)))
    output = tmp_path / "teams.mp4"
    assert render_module.render_teams_link("https://example.invalid", output, max_messages=1) == output
    assert calls == [((comments,), {"output_filename": str(output), "music_code": "pwr", "resolution_scale": 1})]


def test_empty_chat_is_clear_error(monkeypatch):
    monkeypatch.setattr(render_module, "fetch_comments", lambda *args, **kwargs: [])
    with pytest.raises(ValueError, match="no renderable messages"):
        render_module.render_teams_link("https://example.invalid")
