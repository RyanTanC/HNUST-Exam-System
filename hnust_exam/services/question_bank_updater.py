"""题库静默更新服务

从 GitHub 镜像源下载 manifest.json，对比本地缓存，下载差异文件。
支持：镜像源 fallback、单文件重试、SHA256 校验、原子写入、去抖。
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from threading import Thread, Lock
from typing import Callable
from urllib.parse import quote

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from PySide6.QtCore import QObject, Qt, Signal

from hnust_exam.utils.constants import (
    GITHUB_REPO_NAME,
    GITHUB_USERNAME,
    MANIFEST_FILE,
    QUESTION_BANK_DIR,
    QUESTION_BANK_FILES_DIR,
)

logger = logging.getLogger(__name__)

# ── 常量 ──────────────────────────────────────────────────────────────

_LAST_CHECK_FILE = os.path.join(QUESTION_BANK_DIR, "last_check")
_CHECK_INTERVAL = 300          # 去抖间隔（秒）
_API_TIMEOUT = 20              # manifest 请求超时
_DOWNLOAD_TIMEOUT = 60         # 文件下载超时
_MAX_WORKERS = 4               # 并行下载线程数
_MAX_RETRIES = 3               # HTTP 重试次数（Session 级别）
_FILE_RETRIES = 2              # 单文件下载失败重试次数

_REMOTE_DIR = "题库"
_MANIFEST_URL = f"{_REMOTE_DIR}/manifest.json"

_MIRRORS = [
    f"https://raw.githubusercontent.com/{GITHUB_USERNAME}/{GITHUB_REPO_NAME}/master",
    f"https://cdn.jsdelivr.net/gh/{GITHUB_USERNAME}/{GITHUB_REPO_NAME}@master",
]


# ── 数据类 ─────────────────────────────────────────────────────────────

@dataclass
class UpdateResult:
    """更新结果"""
    success: bool
    message: str
    diffs: list[dict]
    error_type: str = ""  # "network", "download", "duplicate", "unknown"


# ── Session 工厂 ──────────────────────────────────────────────────────

def _make_session() -> requests.Session:
    """创建带重试策略的 Session。"""
    session = requests.Session()
    retry = Retry(
        total=_MAX_RETRIES,
        backoff_factor=1.0,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    if token:
        session.headers["Authorization"] = f"token {token}"
    session.headers["Accept"] = "application/vnd.github.v3+json"

    proxy = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
    if proxy:
        session.proxies = {"https": proxy, "http": proxy}

    return session


# ── 镜像源下载 ─────────────────────────────────────────────────────────

def _encode_path(rel_path: str) -> str:
    """逐段 URL 编码路径（保留 / 分隔符）。"""
    return "/".join(quote(p, safe="") for p in rel_path.split("/"))


def _fetch_from_mirrors(
    session: requests.Session, relative_path: str,
) -> requests.Response | None:
    """依次尝试镜像源，返回第一个成功的响应。"""
    encoded = _encode_path(relative_path)
    for mirror in _MIRRORS:
        url = f"{mirror}/{encoded}"
        try:
            resp = session.get(url, timeout=_API_TIMEOUT)
            if resp.status_code == 200:
                return resp
        except requests.RequestException:
            continue
    return None


# ── Manifest 管理 ──────────────────────────────────────────────────────

def _load_local_manifest() -> dict:
    """加载本地 manifest.json。"""
    try:
        if os.path.isfile(MANIFEST_FILE):
            with open(MANIFEST_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        logger.warning("加载本地 manifest 失败: %s", e)
    return {"version": 0, "files": {}}


def _save_local_manifest(manifest: dict) -> None:
    """原子写入 manifest.json（先写 .tmp 再 replace）。"""
    try:
        os.makedirs(os.path.dirname(MANIFEST_FILE), exist_ok=True)
        tmp = MANIFEST_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        os.replace(tmp, MANIFEST_FILE)
    except Exception as e:
        logger.error("保存 manifest 失败: %s", e)


def _fetch_remote_manifest(session: requests.Session) -> dict | None:
    """从镜像源下载远程 manifest.json 并解析。"""
    resp = _fetch_from_mirrors(session, _MANIFEST_URL)
    if resp:
        try:
            return resp.json()
        except Exception as e:
            logger.error("解析远程 manifest 失败: %s", e)
    return None


# ── 哈希计算 ───────────────────────────────────────────────────────────

def _sha256(file_path: str) -> str:
    """计算文件 SHA256。"""
    h = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception as e:
        logger.error("计算哈希失败: %s -> %s", file_path, e)
        return ""


# ── 文件下载 ───────────────────────────────────────────────────────────

def _download_file(
    session: requests.Session, rel_path: str, expected_hash: str,
) -> bool:
    """下载单个文件，带重试和哈希校验。成功返回 True。

    rel_path 是 manifest 中的 key，形如 "PY程序设计模拟题1.xlsx"（无题库前缀）。
    远程路径 = _REMOTE_DIR + "/" + rel_path（即 "题库/PY程序设计模拟题1.xlsx"）。
    本地目标 = QUESTION_BANK_FILES_DIR + "/" + rel_path（即 files/PY程序设计模拟题1.xlsx）。
    """
    target = os.path.join(QUESTION_BANK_FILES_DIR, rel_path)

    for attempt in range(1, _FILE_RETRIES + 1):
        ok = _try_download(session, rel_path, target, expected_hash)
        if ok:
            return True
        if attempt < _FILE_RETRIES:
            logger.info("重试下载 %s（第 %d 次）", rel_path, attempt + 1)
            time.sleep(1)

    logger.error("下载失败（已重试 %d 次）: %s", _FILE_RETRIES, rel_path)
    return False


def _try_download(
    session: requests.Session,
    rel_path: str,
    target: str,
    expected_hash: str,
) -> bool:
    """单次下载尝试。"""
    tmp = target + ".tmp"
    try:
        os.makedirs(os.path.dirname(target), exist_ok=True)

        resp = _fetch_from_mirrors(session, f"{_REMOTE_DIR}/{rel_path}")
        if not resp:
            return False

        with open(tmp, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        if expected_hash:
            actual = _sha256(tmp)
            if actual != expected_hash:
                logger.error("哈希校验失败: %s (期望 %s, 实际 %s)",
                             rel_path, expected_hash[:16], actual[:16])
                _safe_remove(tmp)
                return False

        os.replace(tmp, target)
        return True

    except Exception as e:
        logger.error("下载异常: %s -> %s", rel_path, e)
        _safe_remove(tmp)
        return False


def _safe_remove(path: str) -> None:
    """安全删除文件。"""
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass


# ── 去抖 ───────────────────────────────────────────────────────────────

def _should_check() -> bool:
    """距上次检测超过 _CHECK_INTERVAL 秒才允许检查。"""
    try:
        if os.path.isfile(_LAST_CHECK_FILE):
            with open(_LAST_CHECK_FILE, "r") as f:
                last = float(f.read().strip())
            if time.time() - last < _CHECK_INTERVAL:
                return False
    except Exception:
        pass
    return True


def _mark_checked() -> None:
    """记录本次检测时间。"""
    try:
        os.makedirs(os.path.dirname(_LAST_CHECK_FILE), exist_ok=True)
        with open(_LAST_CHECK_FILE, "w") as f:
            f.write(str(time.time()))
    except Exception as e:
        logger.warning("记录检测时间失败: %s", e)


# ── Qt 信号 ────────────────────────────────────────────────────────────

class _UpdateSignal(QObject):
    """将更新结果派发到主线程。"""
    result = Signal(object)


# ── 核心同步 ───────────────────────────────────────────────────────────

_CHECK_LOCK = Lock()
_IS_CHECKING = False


def is_updating() -> bool:
    """增量更新是否正在进行."""
    with _CHECK_LOCK:
        return _IS_CHECKING


def _sync_question_bank() -> UpdateResult:
    """对比远程和本地 manifest，下载差异文件。"""
    global _IS_CHECKING

    with _CHECK_LOCK:
        if _IS_CHECKING:
            return UpdateResult(False, "更新检查正在进行中...", [], "duplicate")
        _IS_CHECKING = True

    try:
        session = _make_session()

        # 下载远程 manifest
        remote = _fetch_remote_manifest(session)
        if not remote:
            return UpdateResult(
                False, "无法获取远程题库信息（网络问题）", [], "network",
            )

        local = _load_local_manifest()

        # 版本对比
        remote_ver = remote.get("version", 0)
        local_ver = local.get("version", 0)
        if remote_ver <= local_ver:
            return UpdateResult(True, "题库已是最新", [])

        # 差异对比
        remote_files = remote.get("files", {})
        local_files = local.get("files", {})
        to_download: list[str] = []
        diffs: list[dict] = []

        to_delete: list[str] = []

        for fname in sorted(set(remote_files) | set(local_files)):
            rinfo = remote_files.get(fname)
            linfo = local_files.get(fname)
            if rinfo is None:
                # 远程已删除，标记本地文件待删除
                if linfo is not None:
                    diffs.append({"file": fname, "change_type": "deleted"})
                    to_delete.append(fname)
            elif linfo is None:
                diffs.append({"file": fname, "change_type": "added"})
                to_download.append(fname)
            elif linfo.get("hash") != rinfo.get("hash"):
                diffs.append({"file": fname, "change_type": "modified"})
                to_download.append(fname)

        if not to_download and not to_delete:
            return UpdateResult(True, "题库已是最新", [])

        logger.info("发现 %d 个文件需要更新", len(to_download))

        # 并行下载
        downloaded: set[str] = set()
        failed: list[str] = []

        with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as pool:
            futures = {
                pool.submit(
                    _download_file, session, f, remote_files[f].get("hash", ""),
                ): f
                for f in to_download
            }
            for future in as_completed(futures):
                fname = futures[future]
                try:
                    if future.result():
                        downloaded.add(fname)
                    else:
                        failed.append(fname)
                except Exception as e:
                    logger.error("下载线程异常: %s -> %s", fname, e)
                    failed.append(fname)

        # 删除远程已移除的本地文件
        deleted: list[str] = []
        for fname in to_delete:
            target = os.path.join(QUESTION_BANK_FILES_DIR, fname)
            try:
                if os.path.isfile(target):
                    os.remove(target)
                    deleted.append(fname)
                    logger.info("已删除本地多余文件: %s", fname)
            except OSError as e:
                logger.warning("删除文件失败: %s -> %s", fname, e)

        # 更新本地 manifest（只记录成功的）
        if downloaded or deleted:
            merged = {**local_files}
            for f in downloaded:
                merged[f] = remote_files[f]
            for f in deleted:
                merged.pop(f, None)
            _save_local_manifest({"version": remote_ver, "files": merged})

        success_diffs = [d for d in diffs if d["file"] in downloaded or d["file"] in deleted]

        if failed:
            return UpdateResult(
                False,
                f"下载完成：成功 {len(downloaded)}/{len(to_download)}，失败 {len(failed)}",
                success_diffs,
                "download",
            )
        return UpdateResult(
            True,
            f"下载完成：成功 {len(downloaded)}/{len(to_download)}",
            success_diffs,
        )

    except Exception as e:
        logger.error("同步异常: %s", e, exc_info=True)
        return UpdateResult(False, f"更新失败: {e}", [], "unknown")
    finally:
        _IS_CHECKING = False


# ── 公开入口 ───────────────────────────────────────────────────────────

def check_question_bank_update_async(
    callback: Callable[[UpdateResult], None] | None = None,
    force: bool = False,
) -> None:
    """静默检查并更新题库（后台线程，幂等安全）。

    Parameters
    ----------
    callback : 可选回调，接收 UpdateResult（通过 Qt Signal 派发到主线程）。
    force : True 跳过去抖，立即检查（设置页手动按钮用）。
    """
    if not force and not _should_check():
        if callback:
            callback(UpdateResult(True, "距上次检测不足，跳过检查", []))
        return

    sig: _UpdateSignal | None = None
    if callback is not None:
        sig = _UpdateSignal()
        sig.result.connect(callback, Qt.ConnectionType.QueuedConnection)

    def _worker() -> None:
        try:
            result = _sync_question_bank()
            _mark_checked()
            if sig is not None:
                sig.result.emit(result)
        except Exception as e:
            logger.error("同步线程崩溃: %s", e, exc_info=True)
            if sig is not None:
                sig.result.emit(
                    UpdateResult(False, f"更新失败: {e}", [], "unknown"),
                )

    Thread(target=_worker, daemon=True).start()
