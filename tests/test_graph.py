from io import BytesIO

import pytest
from PIL import Image

from objectionengine2_teams.graph import GraphClient, GraphError
from objectionengine2_teams.links import TeamsTarget


class FakeResponse:
    def __init__(self, payload, status=200, headers=None):
        self.payload = payload
        self.status_code = status
        self.headers = headers or {}

    def json(self):
        return self.payload


class FakeSession:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def get(self, url, **kwargs):
        self.calls.append((url, kwargs))
        return self.responses.pop(0)


def test_download_hosted_image_caches_png_and_rejects_external_urls(monkeypatch, tmp_path):
    from objectionengine2_teams import graph as graph_module

    content = BytesIO()
    Image.new("RGB", (2, 2), "red").save(content, format="JPEG")

    class ImageResponse(FakeResponse):
        def raise_for_status(self):
            pass

        def iter_content(self, _size):
            yield content.getvalue()

        def close(self):
            pass

    monkeypatch.setattr(graph_module, "user_cache_path", lambda *args, **kwargs: tmp_path)
    session = FakeSession([ImageResponse(None)])
    client = GraphClient("secret", session=session)
    assert client.download_image("https://example.com/v1.0/hostedContents/id/$value") is None
    path = client.download_image("https://graph.microsoft.com/v1.0/chats/a/messages/b/hostedContents/c/$value")
    assert path is not None
    with Image.open(path) as image:
        assert image.format == "PNG"
        assert image.size == (2, 2)
    assert len(session.calls) == 1
    assert session.calls[0][1]["allow_redirects"] is False


def test_chat_pagination_and_token_header():
    next_url = "https://graph.microsoft.com/v1.0/chats/19%3Aabc%40thread.v2/messages?$skiptoken=next"
    session = FakeSession(
        [
            FakeResponse({"value": [{"id": "2"}], "@odata.nextLink": next_url}),
            FakeResponse({"value": [{"id": "1"}]}),
        ]
    )
    graph = GraphClient("secret-token", session=session)
    target = TeamsTarget("chat", chat_id="19:abc@thread.v2")
    assert [item["id"] for item in graph.iter_messages(target)] == ["2", "1"]
    assert session.calls[0][0].endswith(
        "/chats/19%3Aabc%40thread.v2/messages?$top=50&$orderby=createdDateTime desc"
    )
    assert session.calls[1][0] == next_url
    assert session.calls[0][1]["headers"]["Authorization"] == "Bearer secret-token"
    assert session.calls[0][1]["allow_redirects"] is False


def test_channel_fetches_paginated_replies():
    next_url = "https://graph.microsoft.com/v1.0/teams/team/channels/channel/messages/7/replies?$skiptoken=next"
    session = FakeSession(
        [
            FakeResponse({"value": [{"id": "7"}]}),
            FakeResponse({"value": [{"id": "8"}], "@odata.nextLink": next_url}),
            FakeResponse({"value": [{"id": "9"}]}),
        ]
    )
    graph = GraphClient("token", session=session)
    target = TeamsTarget("channel", team_id="team", channel_id="channel")
    assert [item["id"] for item in graph.iter_messages(target)] == ["7", "8", "9"]
    assert session.calls[1][0].endswith("/messages/7/replies?$top=50")


def test_channel_message_link_fetches_only_its_thread():
    session = FakeSession(
        [
            FakeResponse({"id": "100", "body": {"content": "root"}}),
            FakeResponse({"value": [{"id": "101"}, {"id": "102"}]}),
        ]
    )
    graph = GraphClient("token", session=session)
    target = TeamsTarget(
        "channel", team_id="team", channel_id="channel",
        start_message_id="101", parent_message_id="100",
    )
    assert [item["id"] for item in graph.iter_messages(target)] == ["100", "101", "102"]
    assert [call[0] for call in session.calls] == [
        "https://graph.microsoft.com/v1.0/teams/team/channels/channel/messages/100",
        "https://graph.microsoft.com/v1.0/teams/team/channels/channel/messages/100/replies?$top=50",
    ]


def test_chat_message_link_stops_when_start_is_found():
    next_url = "https://graph.microsoft.com/v1.0/chats/chat/messages?$skiptoken=next"
    session = FakeSession(
        [FakeResponse({"value": [{"id": "4"}, {"id": "3"}], "@odata.nextLink": next_url})]
    )
    graph = GraphClient("token", session=session)
    target = TeamsTarget("chat", chat_id="chat", start_message_id="3")
    assert [item["id"] for item in graph.iter_messages(target)] == ["4", "3"]
    assert len(session.calls) == 1


def test_rejects_external_next_link_before_sending_token():
    session = FakeSession(
        [FakeResponse({"value": [], "@odata.nextLink": "https://evil.example/steal"})]
    )
    graph = GraphClient("secret", session=session)
    with pytest.raises(GraphError, match="unsafe pagination"):
        list(graph.iter_messages(TeamsTarget("chat", chat_id="19:abc@thread.v2")))
    assert len(session.calls) == 1


def test_retries_graph_throttling_using_retry_after():
    session = FakeSession(
        [FakeResponse({}, 429, {"Retry-After": "2"}), FakeResponse({"value": []})]
    )
    waits = []
    graph = GraphClient("token", session=session, wait=waits.append)
    assert list(graph.iter_messages(TeamsTarget("chat", chat_id="19:abc@thread.v2"))) == []
    assert waits == [2.0]


def test_forbidden_response_names_required_scope():
    graph = GraphClient("token", session=FakeSession([FakeResponse({}, 403)]))
    with pytest.raises(GraphError, match="Chat.Read"):
        list(graph.iter_messages(TeamsTarget("chat", chat_id="19:abc@thread.v2")))
