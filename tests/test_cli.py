from objectionengine2_teams import cli


def test_end_link_option_reaches_render_api(monkeypatch, tmp_path, capsys):
    calls = []
    output = tmp_path / "scene.mp4"

    def fake_render(*args, **kwargs):
        calls.append((args, kwargs))
        return output

    monkeypatch.setattr(cli, "render_teams_link", fake_render)
    assert cli.main(["start-link", "--end-link", "end-link", "-o", str(output)]) == 0
    assert calls[0][0] == ("start-link", str(output))
    assert calls[0][1]["end_link"] == "end-link"
    assert calls[0][1]["auth_method"] == "browser"
    assert "Saved" in capsys.readouterr().out


def test_device_code_option_reaches_render_api(monkeypatch):
    calls = []
    monkeypatch.setattr(
        cli, "render_teams_link", lambda *args, **kwargs: calls.append(kwargs) or "teams.mp4"
    )
    assert cli.main(["start-link", "--auth-method", "device-code"]) == 0
    assert calls[0]["auth_method"] == "device-code"
    assert calls[0]["client_id"] == cli.OFFICE_CLIENT_ID
