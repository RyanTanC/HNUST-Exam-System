"""配置管理原子写入测试."""

from __future__ import annotations

import json
import os

from hnust_exam.services.config_manager import ConfigManager


def test_save_config_uses_atomic_replace(tmp_path, monkeypatch):
    config_dir = tmp_path / "cfg"
    config_dir.mkdir()
    config_file = config_dir / "config.json"
    monkeypatch.setattr("hnust_exam.utils.constants._CONFIG_DIR", str(config_dir))
    monkeypatch.setattr("hnust_exam.utils.constants.CONFIG_FILE", str(config_file))

    cm = ConfigManager()
    cm.save_config({"font_scale": 1.5})

    assert config_file.exists()
    assert not (config_dir / "config.json.tmp").exists()
    assert json.loads(config_file.read_text(encoding="utf-8"))["font_scale"] == 1.5


def test_load_config_reads_existing_file(tmp_path, monkeypatch):
    config_dir = tmp_path / "cfg"
    config_dir.mkdir()
    config_file = config_dir / "config.json"
    config_file.write_text('{"font_scale": 1.25}', encoding="utf-8")
    monkeypatch.setattr("hnust_exam.utils.constants._CONFIG_DIR", str(config_dir))
    monkeypatch.setattr("hnust_exam.utils.constants.CONFIG_FILE", str(config_file))

    cm = ConfigManager()
    assert cm.load_config()["font_scale"] == 1.25
