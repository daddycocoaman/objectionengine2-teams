import pytest

from objectionengine2_teams import parse_teams_link


TEAM = "72602e12-78ac-474c-99d6-f619710353a9"
TENANT = "aaaabbbb-0000-cccc-1111-dddd2222eeee"


def test_chat_and_chat_message_links():
    chat = "19:c6d70e392a384916c3262b15406d763e@thread.v2"
    direct = parse_teams_link(f"https://teams.microsoft.com/l/chat/{chat}/conversations")
    message = parse_teams_link(
        f"https://teams.microsoft.com/l/message/{chat}/1563480968434"
        "?context=%7B%22contextType%22%3A%22chat%22%7D"
    )
    assert direct.kind == message.kind == "chat"
    assert direct.chat_id == message.chat_id == chat
    assert direct.start_message_id is None
    assert message.start_message_id == "1563480968434"


def test_channel_and_channel_message_links():
    channel = "19:9be3de4e70874c71a608dee9ba803ed3@thread.tacv2"
    encoded = "19%3A9be3de4e70874c71a608dee9ba803ed3%40thread.tacv2"
    direct = parse_teams_link(
        f"https://teams.microsoft.com/l/channel/{encoded}/General"
        f"?groupId={TEAM}&tenantId={TENANT}"
    )
    message = parse_teams_link(
        f"https://teams.microsoft.com/l/message/{channel}/1648741500652?groupId={TEAM}"
    )
    assert direct.kind == message.kind == "channel"
    assert direct.channel_id == message.channel_id == channel
    assert direct.team_id == message.team_id == TEAM
    assert direct.tenant_id == TENANT
    assert message.start_message_id == message.parent_message_id == "1648741500652"


def test_channel_reply_link_preserves_thread_root_and_selected_message():
    target = parse_teams_link(
        "https://teams.microsoft.com/l/message/19:abc@thread.tacv2/1790883499999"
        f"?tenantId={TENANT}&groupId={TEAM}&parentMessageId=1790883478616"
    )
    assert target.kind == "channel"
    assert target.start_message_id == "1790883499999"
    assert target.parent_message_id == "1790883478616"


def test_team_link_identifies_channel():
    target = parse_teams_link(
        "https://teams.microsoft.com/l/team/19%3Aabc%40thread.tacv2/conversations"
        f"?groupId={TEAM}"
    )
    assert target.kind == "channel"
    assert target.channel_id == "19:abc@thread.tacv2"


@pytest.mark.parametrize(
    "url",
    [
        "http://teams.microsoft.com/l/chat/19:abc@thread.v2/conversations",
        "https://teams.microsoft.com.evil.example/l/chat/19:abc@thread.v2/conversations",
        "https://teams.microsoft.com/l/chat/0/0?users=a@example.com",
        "https://teams.microsoft.com/l/channel/19:abc@thread.tacv2/General",
        "https://teams.microsoft.com/l/message/19:abc@thread.tacv2/1",
        f"https://teams.microsoft.com/l/message/19:abc@thread.tacv2/1?groupId={TEAM}&parentMessageId=bad/id",
        f"https://teams.microsoft.com/l/channel/19:abc@thread.tacv2/General?groupId=bad",
    ],
)
def test_rejects_unsupported_or_unsafe_links(url):
    with pytest.raises(ValueError):
        parse_teams_link(url)
