import json
import os

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


def test_inline_image_is_passed_to_evidence_and_image_only_message_is_kept():
    message = _message("1", "2026-01-01T00:00:00Z", '<img src="https://graph.microsoft.com/v1.0/chats/a/messages/b/hostedContents/c/$value">')
    seen = []
    comments = messages_to_comments([message], image_resolver=lambda url: seen.append(url) or "/tmp/evidence.png")
    assert comments[0].text_content == "Shared an image"
    assert comments[0].evidence_path == "/tmp/evidence.png"
    assert seen == ["https://graph.microsoft.com/v1.0/chats/a/messages/b/hostedContents/c/$value"]


def test_generic_image_alt_text_is_not_added_to_dialogue():
    message = _message("1", "2026-01-01T00:00:00Z", '<p>Look <img src="https://graph.microsoft.com/v1.0/chats/a/messages/b/hostedContents/c/$value" alt="image"> here</p>')
    comment = messages_to_comments([message], image_resolver=lambda _: "/tmp/evidence.png")[0]
    assert comment.text_content == "Look here"
    assert comment.evidence_path == "/tmp/evidence.png"


def test_sharepoint_image_links_and_attachments_are_selected():
    url = "https://tenant.sharepoint.com/:i:/r/sites/team/photo?TeamsCID=abc"
    linked = _message("1", "2026-01-01T00:00:00Z", f'<p>See <a href="{url}">this photo</a></p>')
    attached = _message("2", "2026-01-02T00:00:00Z", "", attachments=[{"name": "photo.jpg", "contentUrl": "https://tenant.sharepoint.com/file"}])
    seen = []
    comments = messages_to_comments([linked, attached], image_resolver=lambda image_url: seen.append(image_url) or "/tmp/evidence.png")
    assert seen == [url, "https://tenant.sharepoint.com/file"]
    assert all(comment.evidence_path == "/tmp/evidence.png" for comment in comments)


def test_fetch_comments_lazily_authenticates_for_sharepoint_images(monkeypatch):
    link = "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations"
    image_url = "https://tenant.sharepoint.com/:i:/r/sites/team/photo?TeamsCID=abc"
    auth_calls = []
    monkeypatch.setattr(render_module, "acquire_access_token", lambda kind, **kwargs: auth_calls.append(kind) or kind + "-token")

    class FakeGraph:
        def __init__(self, token):
            self.token = token

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            pass

        def iter_messages(self, target):
            return iter([_message("1", "2026-01-01T00:00:00Z", f'<p>See <a href="{image_url}">photo</a></p>')])

        def download_sharepoint_image(self, url):
            assert self.token == "files-token"
            assert url == image_url
            return "/tmp/evidence.png"

    monkeypatch.setattr(render_module, "GraphClient", FakeGraph)
    comments = render_module.fetch_comments(link)
    assert auth_calls == ["chat", "files"]
    assert comments[0].evidence_path == "/tmp/evidence.png"


def test_thread_title_is_excluded_and_range_is_respected():
    root = _message(
        "100", "2026-01-01T00:00:00Z", "<p>Body &amp; details</p>",
        subject="  Important   update  ",
    )
    reply = _message("101", "2026-01-02T00:00:00Z", "<p>Reply</p>")

    comments = messages_to_comments([reply, root])
    assert [comment.text_content for comment in comments] == [
        "Body & details", "Reply",
    ]
    assert comments[0].user_name == "Alex"
    assert [comment.text_content for comment in messages_to_comments(
        [root, reply], start_message_id="101"
    )] == ["Reply"]


def test_title_only_root_is_skipped():
    title_only = _message("100", "2026-01-01T00:00:00Z", "", subject="Topic")
    no_title = _message("101", "2026-01-02T00:00:00Z", "Body", subject="  ")
    assert [comment.text_content for comment in messages_to_comments([no_title, title_only])] == [
        "Body",
    ]


def test_empty_and_invisible_messages_are_skipped_before_limit():
    messages = [
        _message("1", "2026-01-01T00:00:00Z", "<p>&nbsp;\u200b</p>"),
        _message("2", "2026-01-02T00:00:00Z", "<p>  Hello  </p>"),
        _message("3", "2026-01-03T00:00:00Z", "<div><br></div>"),
    ]
    assert [comment.text_content for comment in messages_to_comments(messages, max_messages=1)] == ["Hello"]


def test_message_range_is_inclusive_and_starts_at_path_id():
    messages = [
        _message(str(number), f"2026-01-0{number}T00:00:00Z", str(number))
        for number in range(1, 5)
    ]
    comments = messages_to_comments(
        reversed(messages), start_message_id="2", end_message_id="3"
    )
    assert [comment.text_content for comment in comments] == ["2", "3"]
    comments = messages_to_comments(messages, start_message_id="2", max_messages=1)
    assert [comment.text_content for comment in comments] == ["2"]


def test_message_range_rejects_missing_or_reversed_boundaries():
    messages = [_message(str(i), f"2026-01-0{i}T00:00:00Z", str(i)) for i in range(1, 4)]
    with pytest.raises(ValueError, match="Start message.*not found"):
        messages_to_comments(messages, start_message_id="9")
    with pytest.raises(ValueError, match="End message.*not found"):
        messages_to_comments(messages, start_message_id="1", end_message_id="9")
    with pytest.raises(ValueError, match="end message occurs before"):
        messages_to_comments(messages, start_message_id="3", end_message_id="2")


def test_end_link_must_select_same_channel_thread_before_sign_in():
    start = (
        "https://teams.microsoft.com/l/message/19:abc@thread.tacv2/101"
        "?groupId=72602e12-78ac-474c-99d6-f619710353a9&parentMessageId=100"
    )
    end = (
        "https://teams.microsoft.com/l/message/19:abc@thread.tacv2/102"
        "?groupId=72602e12-78ac-474c-99d6-f619710353a9&parentMessageId=999"
    )
    with pytest.raises(ValueError, match="same channel thread"):
        render_module.fetch_comments(start, end_link=end)


def test_fetch_comments_renders_only_requested_channel_thread_range(monkeypatch):
    group_id = "72602e12-78ac-474c-99d6-f619710353a9"
    base = "https://teams.microsoft.com/l/message/19:abc@thread.tacv2/"
    start = f"{base}101?groupId={group_id}&parentMessageId=100"
    end = f"{base}102?groupId={group_id}&parentMessageId=100"
    monkeypatch.setattr(render_module, "acquire_access_token", lambda *args, **kwargs: "token")

    class FakeGraph:
        def __init__(self, token):
            assert token == "token"

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return None

        def iter_messages(self, target, *, end_message_id=None):
            assert target.parent_message_id == "100"
            assert target.start_message_id == "101"
            assert end_message_id == "102"
            return iter(
                [
                    _message(str(i), f"2026-01-0{i - 99}T00:00:00Z", str(i),
                             subject="Thread title" if i == 100 else None)
                    for i in (102, 100, 103, 101)
                ]
            )

    monkeypatch.setattr(render_module, "GraphClient", FakeGraph)
    assert [comment.text_content for comment in render_module.fetch_comments(start, end_link=end)] == [
        "101", "102"
    ]


def test_fetch_comments_excludes_channel_thread_title(monkeypatch):
    group_id = "72602e12-78ac-474c-99d6-f619710353a9"
    link = (
        "https://teams.microsoft.com/l/message/19:abc@thread.tacv2/100"
        f"?groupId={group_id}&parentMessageId=100"
    )
    monkeypatch.setattr(render_module, "acquire_access_token", lambda *args, **kwargs: "token")

    class FakeGraph:
        def __init__(self, token):
            assert token == "token"

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return None

        def iter_messages(self, target):
            assert target.start_message_id == target.parent_message_id == "100"
            return iter([
                _message("100", "2026-01-01T00:00:00Z", "<p>First post</p>",
                         subject="Thread title"),
                _message("101", "2026-01-02T00:00:00Z", "<p>Reply</p>"),
            ])

    monkeypatch.setattr(render_module, "GraphClient", FakeGraph)
    assert [comment.text_content for comment in render_module.fetch_comments(link)] == [
        "First post", "Reply",
    ]


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
    assert auth_calls[0][1]["auth_method"] == "broker"


def test_fetch_comments_passes_device_code_method(monkeypatch):
    calls = []
    monkeypatch.setattr(
        render_module, "acquire_access_token",
        lambda *args, **kwargs: calls.append(kwargs) or "token",
    )

    class FakeGraph:
        def __init__(self, token):
            assert token == "token"

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return None

        def iter_messages(self, target):
            return iter(())

    monkeypatch.setattr(render_module, "GraphClient", FakeGraph)
    render_module.fetch_comments(
        "https://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
        auth_method="device-code",
    )
    assert calls[0]["auth_method"] == "device-code"


def test_render_bridge_passes_comments_to_core(monkeypatch, tmp_path):
    monkeypatch.delenv("oe_sentiment_model", raising=False)
    monkeypatch.delenv("oe_bypass_sentiment", raising=False)
    comments = [Comment(user_id="alex", user_name="Alex", text_content="Objection!")]
    calls = []
    monkeypatch.setattr(render_module, "fetch_comments", lambda *args, **kwargs: comments)

    def fake_render(*args, **kwargs):
        assert os.environ["oe_sentiment_model"] == "hf"
        calls.append((args, kwargs))

    monkeypatch.setattr(render_module, "render_comment_list", fake_render)
    output = tmp_path / "teams.mp4"
    assert render_module.render_teams_link("https://example.invalid", output, max_messages=1) == output
    assert calls == [((comments,), {"output_filename": str(output), "music_code": "pwr", "resolution_scale": 1})]
    config = json.loads((tmp_path / "teams.mp4.json").read_text())
    assert config["link"] == "https://example.invalid"
    assert config["max_messages"] == 1
    assert config["sentiment_model"] == "hf"
    assert config["output_filename"] == str(output)
    assert config["conversation"] == [{
        "user_id": "alex", "user_name": "Alex", "text_content": "Objection!",
        "score": 0, "evidence_path": None,
    }]
    assert "oe_sentiment_model" not in os.environ


def test_render_respects_explicit_neutral_sentiment(monkeypatch, tmp_path):
    monkeypatch.setenv("oe_sentiment_model", "neutral")
    monkeypatch.setattr(
        render_module, "fetch_comments",
        lambda *args, **kwargs: [Comment(text_content="Hello")],
    )
    modes = []
    monkeypatch.setattr(
        render_module,
        "render_comment_list",
        lambda *args, **kwargs: modes.append(os.environ["oe_sentiment_model"]),
    )
    render_module.render_teams_link("https://example.invalid", tmp_path / "teams.mp4")
    assert modes == ["neutral"]
    assert json.loads((tmp_path / "teams.mp4.json").read_text())["sentiment_model"] == "neutral"
    assert os.environ["oe_sentiment_model"] == "neutral"


def test_render_restores_sentiment_setting_after_failure(monkeypatch, tmp_path):
    monkeypatch.delenv("oe_sentiment_model", raising=False)
    monkeypatch.setattr(
        render_module, "fetch_comments",
        lambda *args, **kwargs: [Comment(text_content="Hello")],
    )

    def fail(*args, **kwargs):
        assert os.environ["oe_sentiment_model"] == "hf"
        raise RuntimeError("render failed")

    monkeypatch.setattr(render_module, "render_comment_list", fail)
    with pytest.raises(RuntimeError, match="render failed"):
        render_module.render_teams_link("https://example.invalid", tmp_path / "teams.mp4")
    assert "oe_sentiment_model" not in os.environ
    assert not (tmp_path / "teams.mp4.json").exists()


def test_empty_chat_is_clear_error(monkeypatch):
    monkeypatch.setattr(render_module, "fetch_comments", lambda *args, **kwargs: [])
    with pytest.raises(ValueError, match="no renderable messages"):
        render_module.render_teams_link("https://example.invalid")
