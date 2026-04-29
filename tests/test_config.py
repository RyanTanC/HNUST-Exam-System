"""配置管理测试."""

import os
import json
import tempfile
import shutil

import pytest

from hnust_exam.services.config_manager import ConfigManager


@pytest.fixture
def temp_config(monkeypatch):
    """创建临时配置目录."""
    tmpdir = tempfile.mkdtemp()
    config_file = os.path.join(tmpdir, "config.json")
    progress_file = os.path.join(tmpdir, "progress.json")
    skip_file = os.path.join(tmpdir, "skip_ver")

    import hnust_exam.utils.constants as const
    monkeypatch.setattr(const, "_CONFIG_DIR", tmpdir)
    monkeypatch.setattr(const, "CONFIG_FILE", config_file)
    monkeypatch.setattr(const, "PROGRESS_FILE", progress_file)
    monkeypatch.setattr(const, "SKIP_VERSION_FILE", skip_file)

    yield tmpdir, config_file, progress_file, skip_file

    shutil.rmtree(tmpdir, ignore_errors=True)


class TestConfigManager:
    """配置读写测试."""

    def test_load_defaults(self, temp_config):
        cm = ConfigManager()
        config = cm.load_config()
        assert config["font_scale"] == 1.0
        assert config["dark_mode"] is False
        assert config["show_answer_immediately"] is False

    def test_save_and_load(self, temp_config):
        cm = ConfigManager()
        cm.save_config({"font_scale": 1.2, "dark_mode": True})
        config = cm.load_config()
        assert config["font_scale"] == 1.2
        assert config["dark_mode"] is True

    def test_preserve_unknown_keys(self, temp_config):
        tmpdir, config_file, _, _ = temp_config
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump({"font_scale": 1.3, "custom_key": "hello"}, f)
        cm = ConfigManager()
        config = cm.load_config()
        assert config["font_scale"] == 1.3
        assert config["custom_key"] == "hello"


class TestProgress:
    """进度读写测试."""

    def test_load_empty_progress(self, temp_config):
        cm = ConfigManager()
        progress = cm.load_progress()
        assert progress == {}

    def test_save_and_load_progress(self, temp_config):
        cm = ConfigManager()
        data = {"exam1.xlsx": {"status": "completed", "best_score": 95.0}}
        cm.save_progress(data)
        loaded = cm.load_progress()
        assert loaded["exam1.xlsx"]["status"] == "completed"
        assert loaded["exam1.xlsx"]["best_score"] == 95.0
