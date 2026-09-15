"""更新检查服务：GitHub 查版本号，Gitee 下载 exe."""

from __future__ import annotations

from datetime import datetime
from threading import Thread
from typing import Callable

import requests
from PySide6.QtCore import QObject, Signal

import logging

from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.constants import (
    CURRENT_VERSION,
    FALLBACK_DOWNLOAD_URL,
    GITEE_OWNER,
    GITEE_UPDATE_REPO,
    GITHUB_USERNAME,
    GITHUB_REPO_NAME,
)
from hnust_exam.utils.helpers import version_tuple

logger = logging.getLogger(__name__)

# 保留信号引用，防止被 GC 回收后跨线程信号丢失
# 见 check_update_async 中的使用和清理逻辑
_pending_signals: list[_UpdateSignal] = []


class _UpdateSignal(QObject):
    """内部信号，用于将回调派发到主线程."""
    result = Signal(object)


def fetch_latest_release_info(github_token: str = "") -> dict | None:
    """同步获取 GitHub 最新发布信息，获取失败返回 None.

    github_token: GitHub Personal Access Token，用于提升 API 速率限制（5000次/小时）。
    """
    try:
        repo_api_url = (
            f"https://api.github.com/repos/{GITHUB_USERNAME}"
            f"/{GITHUB_REPO_NAME}/releases/latest"
        )
        headers = {}
        if github_token:
            headers["Authorization"] = f"Bearer {github_token}"
        response = requests.get(repo_api_url, headers=headers, timeout=5)
        response.raise_for_status()
        data = response.json()

        latest_version = data.get("tag_name", "")
        if not latest_version:
            return None

        release_notes = (data.get("body", "") or "").strip() or "暂无更新日志"

        download_url = ""
        expected_size = 0
        download_available = False
        assets = data.get("assets", [])
        if assets:
            # 优先选择 .exe 文件（Windows 安装包）
            exe_asset = next(
                (a for a in assets if a.get("name", "").lower().endswith(".exe")),
                None,
            )
            chosen = exe_asset or assets[0]
            download_url = chosen.get("browser_download_url", "")
            expected_size = chosen.get("size", 0)
            if download_url:
                download_available = True

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
            "download_available": download_available,
            "published_at": published,
            "expected_size": expected_size,
            "release_url": data.get("html_url", ""),
        }
    except Exception as e:
        logger.warning("获取 Release 信息失败: %s", e)
    return None


def fetch_gitee_release_info(tag_name: str) -> dict | None:
    """同步获取 Gitee 指定 tag 发布信息（仅用于下载），获取失败返回 None.

    注意：不通过列表接口（排序不稳定），直接用 tag 精确查询。
    """
    try:
        repo_api_url = (
            f"https://gitee.com/api/v5/repos/{GITEE_OWNER}"
            f"/{GITEE_UPDATE_REPO}/releases/tags/{tag_name}"
        )
        response = requests.get(repo_api_url, timeout=5)
        response.raise_for_status()
        release = response.json()

        latest_version = release.get("tag_name", "")
        if not latest_version:
            return None

        download_url = ""
        expected_size = 0
        download_available = False
        assets = release.get("assets", [])
        if assets:
            exe_asset = next(
                (a for a in assets if a.get("name", "").lower().endswith(".exe")),
                None,
            )
            chosen = exe_asset or assets[0]
            download_url = chosen.get("browser_download_url", "")
            expected_size = chosen.get("size", 0)
            if download_url:
                download_available = True

        return {
            "latest_ver": latest_version,
            "download_url": download_url,
            "download_available": download_available,
            "expected_size": expected_size,
        }
    except Exception as e:
        logger.warning("获取 Gitee Release 信息失败: %s", e)
    return None


def fetch_latest_gitee_release_info() -> dict | None:
    """从 Gitee 获取最新发布信息（遍历全部 release，取最高版本号）.

    Gitee 的 release 列表排序不可靠，因此遍历全部 tag 用版本号比较。
    获取失败返回 None。
    """
    try:
        repo_api_url = (
            f"https://gitee.com/api/v5/repos/{GITEE_OWNER}"
            f"/{GITEE_UPDATE_REPO}/releases"
        )
        response = requests.get(repo_api_url, timeout=5)
        response.raise_for_status()
        releases = response.json()
        if not releases:
            return None

        best = None
        best_version = None

        for release in releases:
            tag_name = release.get("tag_name", "")
            if not tag_name:
                continue
            try:
                current = version_tuple(tag_name)
            except Exception:
                continue
            if best_version is None or current > best_version:
                best_version = current
                best = release

        if best is None:
            return None

        latest_version = best.get("tag_name", "")
        if not latest_version:
            return None

        release_notes = (best.get("body", "") or "").strip() or "暂无更新日志"

        download_url = ""
        expected_size = 0
        download_available = False
        assets = best.get("assets", [])
        if assets:
            exe_asset = next(
                (a for a in assets if a.get("name", "").lower().endswith(".exe")),
                None,
            )
            chosen = exe_asset or assets[0]
            download_url = chosen.get("browser_download_url", "")
            expected_size = chosen.get("size", 0)
            if download_url:
                download_available = True

        published = best.get("created_at", "")
        if published:
            try:
                # Gitee 格式: 2026-05-29T20:00:11+08:00
                # GitHub 格式: 2026-05-27T13:21:30Z
                dt = datetime.fromisoformat(published.replace("Z", "+00:00"))
                published = dt.strftime("%Y年%m月%d日 %H:%M")
            except Exception:
                pass

        return {
            "latest_ver": latest_version,
            "release_notes": release_notes,
            "download_url": download_url,
            "download_available": download_available,
            "published_at": published,
            "expected_size": expected_size,
            "release_url": best.get("html_url", ""),
        }
    except Exception as e:
        logger.warning("获取 Gitee 最新 Release 信息失败: %s", e)
    return None


def _is_update_needed(
    latest_version: str,
    config_manager: ConfigManager | None = None,
) -> bool:
    """检查是否确实需要更新（版本比较 + 跳过标记)."""
    if version_tuple(latest_version) <= version_tuple(CURRENT_VERSION):
        return False
    if config_manager and config_manager.load_skip_version() == latest_version:
        return False
    return True


def fetch_update_info(
    config_manager: ConfigManager | None = None,
) -> dict | None:
    """同步获取更新信息，无更新返回 None.

    流程：
    1. 优先从 Gitee 获取最新版本号与下载信息（国内访问稳定）
    2. Gitee 失败时降级到 GitHub 获取版本号
    3. GitHub 降级时，尝试补充 Gitee 下载信息
    4. 均不可用时返回 None
    """
    token = ""
    if config_manager:
        token = config_manager.load_config().get("github_token", "")

    # ── 优先从 Gitee 获取 ──
    info = fetch_latest_gitee_release_info()
    if info:
        if _is_update_needed(info["latest_ver"], config_manager):
            return info
        return {"no_update": True, "latest_ver": info["latest_ver"]}

    # ── Gitee 不可用 → 降级到 GitHub ──
    gh_info = fetch_latest_release_info(github_token=token)
    if not gh_info:
        return None

    if not _is_update_needed(gh_info["latest_ver"], config_manager):
        return {"no_update": True, "latest_ver": gh_info["latest_ver"]}

    latest_version = gh_info["latest_ver"]

    # 尝试从 Gitee 补充下载信息
    gitee_info = fetch_gitee_release_info(latest_version)
    if (gitee_info
            and gitee_info["latest_ver"] == latest_version
            and gitee_info["download_available"]):
        gh_info["download_url"] = gitee_info["download_url"]
        gh_info["download_available"] = True
        gh_info["expected_size"] = gitee_info["expected_size"]
    else:
        gh_info["download_url"] = ""
        gh_info["download_available"] = False
        gh_info["fallback_url"] = FALLBACK_DOWNLOAD_URL

    return gh_info


def check_update_async(
    callback: Callable[[dict | None], None],
    config_manager: ConfigManager | None = None,
) -> None:
    """异步检查更新，通过回调函数通知结果.

    回调通过 Qt Signal 派发到主线程执行，确保线程安全。

    注意：_pending_signals 持有信号引用防止 GC 回收后跨线程投递丢失。
    """

    sig = _UpdateSignal()
    _pending_signals.append(sig)

    def _on_signal(info: object) -> None:
        """包装回调：执行后清理信号引用，允许 GC 回收."""
        try:
            callback(info)
        finally:
            try:
                _pending_signals.remove(sig)
            except ValueError:
                pass

    sig.result.connect(_on_signal)

    def _worker() -> None:
        info = None
        try:
            info = fetch_update_info(config_manager)
        except Exception as e:
            logger.warning("更新检查线程异常: %s", e)
        finally:
            # 必须保证回调一定触发（异常时投递 None，调用方按"网络不可用"处理），
            # 否则调用方拿不到任何结果，界面会一直停在"检查中"。
            sig.result.emit(info)

    Thread(target=_worker, daemon=True).start()
