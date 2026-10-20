import pytest

from objectionengine2_teams.auth import (
    OFFICE_CLIENT_ID,
    AuthenticationError,
    acquire_access_token,
)


class FakeApp:
    def __init__(self, client_id, authority):
        self.client_id = client_id
        self.authority = authority
        self.scopes = None
        self.result = {"access_token": "example-token"}

    def get_accounts(self):
        return []

    def initiate_device_flow(self, scopes):
        self.scopes = scopes
        return {"user_code": "TEST-CODE", "message": "Visit the device URL"}

    def acquire_token_by_device_flow(self, flow):
        assert flow["user_code"] == "TEST-CODE"
        return self.result


@pytest.mark.parametrize(
    "kind,scope", [("chat", "Chat.Read"), ("channel", "ChannelMessage.Read.All")]
)
def test_device_flow_requests_only_needed_scope(kind, scope):
    apps = []
    prompts = []

    def factory(**kwargs):
        app = FakeApp(**kwargs)
        apps.append(app)
        return app

    token = acquire_access_token(
        kind,
        tenant_id="aaaabbbb-0000-cccc-1111-dddd2222eeee",
        app_factory=factory,
        show_device_code=prompts.append,
    )
    assert token == "example-token"
    assert apps[0].client_id == OFFICE_CLIENT_ID
    assert apps[0].authority.endswith("/aaaabbbb-0000-cccc-1111-dddd2222eeee")
    assert apps[0].scopes == [scope]
    assert prompts == ["Visit the device URL"]


def test_authentication_failure_has_consent_guidance():
    app = FakeApp(OFFICE_CLIENT_ID, "authority")
    app.result = {"error": "consent_required"}
    with pytest.raises(AuthenticationError, match="ChannelMessage.Read.All.*administrator"):
        acquire_access_token("channel", app_factory=lambda **_: app, show_device_code=lambda _: None)
