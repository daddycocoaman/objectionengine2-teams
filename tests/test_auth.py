import stat
from types import SimpleNamespace

import pytest
from azure.core.exceptions import ClientAuthenticationError
from azure.identity import AuthenticationRecord

from objectionengine2_teams import auth


class FakeCredential:
    instances = []
    failure = None

    def __init__(self, **kwargs):
        self.kwargs = kwargs
        self.auth_scopes = None
        self.token_scopes = None
        self.closed = False
        self.__class__.instances.append(self)

    def authenticate(self, *, scopes):
        if self.failure:
            raise self.failure
        self.auth_scopes = scopes
        return AuthenticationRecord(
            tenant_id=self.kwargs["tenant_id"],
            client_id=self.kwargs["client_id"],
            authority="login.microsoftonline.com",
            home_account_id="test-account",
            username="test@example.com",
        )

    def get_token(self, *scopes):
        self.token_scopes = scopes
        return SimpleNamespace(token="example-token")

    def close(self):
        self.closed = True


class FakeBrowserCredential(FakeCredential):
    instances = []


class FakeDeviceCredential(FakeCredential):
    instances = []


@pytest.fixture(autouse=True)
def fake_credentials(monkeypatch):
    FakeBrowserCredential.instances.clear()
    FakeDeviceCredential.instances.clear()
    FakeBrowserCredential.failure = None
    FakeDeviceCredential.failure = None
    monkeypatch.setattr(auth, "InteractiveBrowserCredential", FakeBrowserCredential)
    monkeypatch.setattr(auth, "DeviceCodeCredential", FakeDeviceCredential)


@pytest.mark.parametrize(
    "kind,scope", [
        ("chat", "https://graph.microsoft.com/Chat.Read"),
        ("channel", "https://graph.microsoft.com/ChannelMessage.Read.All"),
    ]
)
def test_browser_is_default_with_office_client_and_least_scope(kind, scope, tmp_path):
    token = auth.acquire_access_token(
        kind,
        tenant_id="aaaabbbb-0000-cccc-1111-dddd2222eeee",
        record_dir=tmp_path,
    )
    credential = FakeBrowserCredential.instances[0]
    assert token == "example-token"
    assert not FakeDeviceCredential.instances
    assert credential.kwargs["client_id"] == auth.OFFICE_CLIENT_ID
    assert credential.kwargs["tenant_id"] == "aaaabbbb-0000-cccc-1111-dddd2222eeee"
    assert credential.kwargs["redirect_uri"] == auth.BROWSER_REDIRECT_URI
    assert credential.auth_scopes == [scope]
    assert credential.token_scopes == (scope,)
    assert credential.kwargs["cache_persistence_options"].allow_unencrypted_storage is True
    assert credential.kwargs["cache_persistence_options"].name.startswith("objectionengine2-teams-")
    assert credential.closed
    records = list(tmp_path.glob("*.json"))
    assert len(records) == 1
    assert stat.S_IMODE(records[0].stat().st_mode) == 0o600
    assert "example-token" not in records[0].read_text()


def test_device_code_uses_same_cache_and_saved_account(tmp_path):
    auth.acquire_access_token("chat", record_dir=tmp_path)
    browser = FakeBrowserCredential.instances[0]
    token = auth.acquire_access_token("chat", auth_method="device-code", record_dir=tmp_path)
    device = FakeDeviceCredential.instances[0]
    assert token == "example-token"
    assert device.kwargs["client_id"] == auth.OFFICE_CLIENT_ID
    assert device.kwargs["authentication_record"].username == "test@example.com"
    assert device.kwargs["cache_persistence_options"].name == browser.kwargs["cache_persistence_options"].name
    assert device.auth_scopes is None
    assert device.token_scopes == ("https://graph.microsoft.com/Chat.Read",)
    assert device.closed


def test_device_code_can_start_sign_in_without_a_record(tmp_path):
    auth.acquire_access_token("channel", auth_method="device-code", record_dir=tmp_path)
    assert not FakeBrowserCredential.instances
    assert FakeDeviceCredential.instances[0].auth_scopes == [
        "https://graph.microsoft.com/ChannelMessage.Read.All"
    ]


def test_cache_is_isolated_by_tenant_and_client(tmp_path):
    auth.acquire_access_token("chat", tenant_id="tenant-a", record_dir=tmp_path)
    auth.acquire_access_token("chat", tenant_id="tenant-b", record_dir=tmp_path)
    auth.acquire_access_token("chat", tenant_id="tenant-a", client_id="other-client", record_dir=tmp_path)
    assert len(list(tmp_path.glob("*.json"))) == 3
    assert all(c.auth_scopes is not None for c in FakeBrowserCredential.instances)
    assert len({c.kwargs["cache_persistence_options"].name for c in FakeBrowserCredential.instances}) == 3


def test_invalid_record_has_clear_remedy(tmp_path):
    path = tmp_path / f"{auth._cache_key(auth.OFFICE_CLIENT_ID, 'organizations')}.json"
    path.write_text("not json")
    with pytest.raises(auth.AuthenticationError, match="Delete it and sign in again"):
        auth.acquire_access_token("chat", record_dir=tmp_path)
    assert not FakeBrowserCredential.instances


def test_authentication_failure_has_consent_guidance(tmp_path):
    FakeBrowserCredential.failure = ClientAuthenticationError(message="consent_required")
    with pytest.raises(auth.AuthenticationError, match="ChannelMessage.Read.All.*administrator"):
        auth.acquire_access_token("channel", record_dir=tmp_path)
    assert FakeBrowserCredential.instances[0].closed


def test_invalid_auth_method_is_rejected_before_sign_in(tmp_path):
    with pytest.raises(ValueError, match="auth_method"):
        auth.acquire_access_token("chat", auth_method="unknown", record_dir=tmp_path)
    assert not FakeBrowserCredential.instances
