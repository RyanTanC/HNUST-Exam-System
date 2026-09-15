"""配置管理服务：线程安全的 JSON 配置读写."""

from __future__ import annotations

import json
import os
import tempfile
import threading
from typing import Any

from hnust_exam.utils import constants


def _ensure_config_dir() -> None:
    os.makedirs(constants._CONFIG_DIR, exist_ok=True)


def _atomic_write_text(path: str, content: str) -> None:
    """原子写入文本文件，避免崩溃时截断目标文件."""
    directory = os.path.dirname(path)
    os.makedirs(directory, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix=os.path.basename(path), suffix=".tmp", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(content)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, path)
    except Exception:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise


class ConfigManager:
    """线程安全的 JSON 配置管理器."""

    DEFAULTS: dict[str, Any] = {
        "font_scale": 1.0,
        "dark_mode": False,
        "show_answer_immediately": False,
        "auto_advance_on_choice": False,
        "user_python_path": "",
        "student_name": "",
        "student_id": "",
        "grading_strictness": "normal",
        "github_token": "",
    }

    def __init__(self) -> None:
        self._lock = threading.Lock()

    def _ensure_dir(self) -> None:
        _ensure_config_dir()

    # ── 配置 ──────────────────────────────────────────────────────

    def load_config(self) -> dict[str, Any]:
        """加载配置，合并默认值."""
        self._ensure_dir()
        defaults = dict(self.DEFAULTS)
        with self._lock:
            try:
                if os.path.exists(constants.CONFIG_FILE):
                    with open(constants.CONFIG_FILE, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    defaults.update(data)
            except Exception:
                pass
        return defaults

    def save_config(self, config: dict[str, Any]) -> None:
        """保存配置."""
        self._ensure_dir()
        with self._lock:
            try:
                content = json.dumps(config, ensure_ascii=False, indent=2)
                _atomic_write_text(constants.CONFIG_FILE, content)
            except Exception:
                pass

    # ── 进度 ──────────────────────────────────────────────────────

    def load_progress(self) -> dict[str, Any]:
        """加载考试进度."""
        self._ensure_dir()
        with self._lock:
            try:
                if os.path.exists(constants.PROGRESS_FILE):
                    with open(constants.PROGRESS_FILE, "r", encoding="utf-8") as f:
                        return json.load(f)
            except Exception:
                pass
        return {}

    def save_progress(self, progress: dict[str, Any]) -> None:
        """保存考试进度."""
        self._ensure_dir()
        with self._lock:
            try:
                content = json.dumps(progress, ensure_ascii=False, indent=2)
                _atomic_write_text(constants.PROGRESS_FILE, content)
            except Exception:
                pass

    # ── 跳过版本 ──────────────────────────────────────────────────

    def load_skip_version(self) -> str:
        """加载跳过的版本号."""
        with self._lock:
            try:
                if os.path.exists(constants.SKIP_VERSION_FILE):
                    with open(constants.SKIP_VERSION_FILE, "r", encoding="utf-8") as f:
                        return f.read().strip()
            except Exception:
                pass
        return ""

    def save_skip_version(self, ver: str) -> None:
        """保存跳过的版本号."""
        self._ensure_dir()
        with self._lock:
            try:
                _atomic_write_text(constants.SKIP_VERSION_FILE, ver)
            except Exception:
                pass
