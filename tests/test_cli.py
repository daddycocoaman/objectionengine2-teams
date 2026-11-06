import os
from pathlib import Path
import shutil
import subprocess
import sys

import pydub
from objectionengine2_teams import cli


def test_pydub_compile_warnings_are_suppressed_only_during_import(tmp_path):
    shutil.copytree(
        Path(pydub.__file__).parent, tmp_path / "pydub",
        ignore=shutil.ignore_patterns("__pycache__"),
    )
    env = dict(os.environ, PYTHONPATH=str(tmp_path))
    result = subprocess.run(
        [sys.executable, "-B", "-c", 'import warnings; warnings.simplefilter("always", SyntaxWarning); import objectionengine2_teams; warnings.warn("other warning", SyntaxWarning)'],
        env=env, capture_output=True, text=True, check=True,
    )
    assert "pydub/utils.py" not in result.stderr.replace("\\", "/")
    assert "other warning" in result.stderr


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
    assert calls[0][1]["auth_method"] == "broker"
    assert "Saved" in capsys.readouterr().out


def test_device_code_option_reaches_render_api(monkeypatch):
    calls = []
    monkeypatch.setattr(
        cli, "render_teams_link", lambda *args, **kwargs: calls.append(kwargs) or "teams.mp4"
    )
    assert cli.main(["start-link", "--auth-method", "device-code"]) == 0
    assert calls[0]["auth_method"] == "device-code"
    assert calls[0]["client_id"] == cli.OFFICE_CLIENT_ID
