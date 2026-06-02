"""Python 环境查找与打开方式测试."""

from hnust_exam.services import python_env


def test_normalize_python_selection_accepts_file(tmp_path):
    python = tmp_path / "python3"
    python.write_text("", encoding="utf-8")

    assert python_env.normalize_python_selection(str(python)) == str(python)


def test_normalize_python_selection_accepts_macos_idle_app(monkeypatch, tmp_path):
    monkeypatch.setattr(python_env.sys, "platform", "darwin")
    idle_app = tmp_path / "IDLE.app"
    idle_app.mkdir()

    assert python_env.normalize_python_selection(str(idle_app)) == str(idle_app)


def test_normalize_python_selection_accepts_macos_python_folder(monkeypatch, tmp_path):
    monkeypatch.setattr(python_env.sys, "platform", "darwin")
    python_dir = tmp_path / "Python 3.13"
    idle_app = python_dir / "IDLE.app"
    idle_app.mkdir(parents=True)

    assert python_env.normalize_python_selection(str(python_dir)) == str(idle_app)


def test_open_with_idle_uses_selected_macos_idle_app(monkeypatch, tmp_path):
    monkeypatch.setattr(python_env.sys, "platform", "darwin")
    idle_app = tmp_path / "IDLE.app"
    idle_app.mkdir()
    program_file = tmp_path / "Prog00001.py"
    program_file.write_text("print('hello')\n", encoding="utf-8")
    calls = []

    def fake_popen(args, **kwargs):
        calls.append((args, kwargs))
        return object()

    monkeypatch.setattr(python_env.subprocess, "Popen", fake_popen)

    assert python_env.open_with_idle(str(program_file), str(idle_app)) is True
    assert calls == [
        (["open", "-a", str(idle_app), str(program_file)], {})
    ]


def test_open_with_python_idle_requires_idlelib(monkeypatch, tmp_path):
    program_file = tmp_path / "Prog00001.py"
    program_file.write_text("", encoding="utf-8")
    calls = []

    def fake_popen(args, **kwargs):
        calls.append((args, kwargs))
        return object()

    monkeypatch.setattr(python_env, "_python_supports_idle", lambda path: False)
    monkeypatch.setattr(python_env.subprocess, "Popen", fake_popen)

    assert python_env._open_with_python_idle("/tmp/python3", str(program_file)) is False
    assert calls == []


def test_open_with_default_app_uses_macos_open(monkeypatch, tmp_path):
    monkeypatch.setattr(python_env.sys, "platform", "darwin")
    program_file = tmp_path / "Prog00001.py"
    program_file.write_text("", encoding="utf-8")
    calls = []

    def fake_run(args, check):
        calls.append((args, check))

    monkeypatch.setattr(python_env.subprocess, "run", fake_run)

    assert python_env.open_with_default_app(str(program_file)) is True
    assert calls == [(["open", str(program_file)], True)]
