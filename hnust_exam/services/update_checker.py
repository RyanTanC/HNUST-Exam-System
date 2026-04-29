"""更新检查服务：异步检查 GitHub 最新版本."""

from __future__ import annotations

from datetime import datetime
from threading import Thread
from typing import Callable

import requests
from PySide6.QtCore import QObject, Signal

from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.constants import (
    CURRENT_VERSION,
    GITHUB_USERNAME,
    GITHUB_REPO_NAME,
)
from hnust_exam.utils.helpers import version_tuple


class _UpdateSignal(QObject):
    """内部信号，用于将回调派发到主线程."""
    result = Signal(object)


def fetch_update_info(
    config_manager: ConfigManager | None = None,
) -> dict | None:
    """同步获取更新信息，无更新返回 None."""
    try:
        repo_api_url = (
            f"https://api.github.com/repos/{GITHUB_USERNAME}"
            f"/{GITHUB_REPO_NAME}/releases/latest"
        )
        response = requests.get(repo_api_url, timeout=5)
        response.raise_for_status()
        data = response.json()

        latest_version = data.get("tag_name", "")
        if not latest_version:
            return None
        if version_tuple(latest_version) <= version_tuple(CURRENT_VERSION):
            return None

        # 检查是否跳过此版本
        if config_manager:
            if config_manager.load_skip_version() == latest_version:
                return None

        release_notes = (data.get("body", "") or "").strip() or "暂无更新日志"

        download_url = ""
        assets = data.get("assets", [])
        if assets:
            download_url = assets[0].get("browser_download_url", "")
        if not download_url:
            download_url = data.get("html_url", "")

        published = data.get("published_at", "")
        if published:
            try:
                dt = datetime.strptime(published, "%Y-%m-%dT%H:%M:%SZ")
                published = dt.strftime("%Y年%m月%d日 %H:%M")
            except Exception:
                pass

        return {
            "latest_ver": latest_version,
            "release_notes": release_notes,
            "download_url": download_url,
            "published_at": published,
        }
    except Exception:
        pass
    return None


def check_update_async(
    callback: Callable[[dict | None], None],
    config_manager: ConfigManager | None = None,
) -> None:
    """异步检查更新，通过回调函数通知结果.

    回调通过 Qt Signal 派发到主线程执行，确保线程安全。
    """

    sig = _UpdateSignal()
    sig.result.connect(callback)

    def _worker() -> None:
        try:
            info = fetch_update_info(config_manager)
            sig.result.emit(info)
        except Exception:
            pass

    Thread(target=_worker, daemon=True).start()
