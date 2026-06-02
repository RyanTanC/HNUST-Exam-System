"""题库整包热更新服务.

从 Gitee Releases 下载 question_bank.zip，校验后原子替换本地题库。
支持：SHA256 校验、staging 暂存、原子替换、失败回滚、非阻塞 UI。
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import shutil
import time
import zipfile
from dataclasses import dataclass
from threading import Thread, Lock
from typing import Callable

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from PySide6.QtCore import QObject, Qt, Signal

from hnust_exam.utils.constants import (
    GITEE_REPO_NAME,
    GITEE_USERNAME,
    QUESTION_BANK_DIR,
    QUESTION_BANK_FILES_DIR,
)

logger = logging.getLogger(__name__)

# ── 常量 ──────────────────────────────────────────────────────────────

_STAGING_DIR = os.path.join(QUESTION_BANK_DIR, "staging")
_BACKUP_DIR = os.path.join(QUESTION_BANK_DIR, "files_backup")
_CURRENT_VERSION_FILE = os.path.join(QUESTION_BANK_DIR, "current_version")
_ZIP_FILENAME = "question_bank.zip"
_HASH_FILENAME = "question_bank.zip.sha256"
_DOWNLOAD_TIMEOUT = 120
_API_TIMEOUT = 15
_MAX_RETRIES = 3


# ── 数据类 ─────────────────────────────────────────────────────────────

@dataclass
class PackUpdateResult:
    """整包更新结果."""
    success: bool
    message: str
    new_version: str = ""
    error_type: str = ""  # "network", "download", "verify", "extract", "unknown"


# ── Session 工厂 ──────────────────────────────────────────────────────

def _make_session() -> requests.Session:
    """创建带重试策略的 Session."""
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

    proxy = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
    if proxy:
        session.proxies = {"https": proxy, "http": proxy}

    return session


# ── 哈希计算 ───────────────────────────────────────────────────────────

def _sha256_file(file_path: str) -> str:
    """计算文件 SHA256."""
    h = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception as e:
        logger.error("计算哈希失败: %s -> %s", file_path, e)
        return ""


def _parse_sha256_file(content: str) -> str:
    """解析 sha256 校验文件内容，返回哈希值.

    支持格式：
    - 纯哈希字符串
    - "hash  filename" 格式
    """
    content = content.strip()
    if not content:
        return ""
    # 取第一行，按空格分割，第一段是哈希
    first_line = content.splitlines()[0].strip()
    return first_line.split()[0] if first_line else ""


# ── Gitee Release 查询 ───────────────────────────────────────────────

def _fetch_release_assets(session: requests.Session) -> dict | None:
    """获取最新 Release 的 asset 信息（从 Gitee）.

    返回 {"zip_url": ..., "zip_size": ..., "hash_url": ..., "tag_name": ...} 或 None。
    """
    url = f"https://gitee.com/api/v5/repos/{GITEE_USERNAME}/{GITEE_REPO_NAME}/releases/latest"
    try:
        resp = session.get(url, timeout=_API_TIMEOUT)
        resp.raise_for_status()
        data = resp.json()

        tag_name = data.get("tag_name", "")
        assets = data.get("assets", [])

        zip_url = ""
        zip_size = 0
        hash_url = ""

        for asset in assets:
            name = asset.get("name", "")
            download_url = asset.get("browser_download_url", "")
            size = asset.get("size", 0)

            if name == _ZIP_FILENAME:
                zip_url = download_url
                zip_size = size
            elif name == _HASH_FILENAME:
                hash_url = download_url

        if not zip_url:
            logger.info("最新 Release 中未找到 %s", _ZIP_FILENAME)
            return None

        return {
            "zip_url": zip_url,
            "zip_size": zip_size,
            "hash_url": hash_url,
            "tag_name": tag_name,
        }
    except Exception as e:
        logger.error("查询 Release 失败: %s", e)
        return None


# ── 版本比较 ───────────────────────────────────────────────────────────

def _get_local_version() -> str:
    """读取本地题库版本号."""
    try:
        if os.path.isfile(_CURRENT_VERSION_FILE):
            with open(_CURRENT_VERSION_FILE, "r", encoding="utf-8") as f:
                return f.read().strip()
    except Exception:
        pass
    return ""


def _save_local_version(version: str) -> None:
    """保存本地题库版本号."""
    try:
        os.makedirs(os.path.dirname(_CURRENT_VERSION_FILE), exist_ok=True)
        tmp = _CURRENT_VERSION_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(version)
        os.replace(tmp, _CURRENT_VERSION_FILE)
    except Exception as e:
        logger.error("保存版本号失败: %s", e)


# ── 文件下载 ───────────────────────────────────────────────────────────

def _download_to_file(
    session: requests.Session,
    url: str,
    target: str,
    expected_size: int = 0,
    progress_callback: Callable[[int, int], None] | None = None,
) -> bool:
    """流式下载文件到目标路径.

    Returns True on success.
    """
    try:
        resp = session.get(url, stream=True, timeout=_DOWNLOAD_TIMEOUT)
        resp.raise_for_status()
    except requests.RequestException as e:
        logger.error("下载请求失败: %s -> %s", url, e)
        return False

    content_length = resp.headers.get("content-length")
    total = int(content_length) if content_length else expected_size

    downloaded = 0
    try:
        os.makedirs(os.path.dirname(target), exist_ok=True)
        with open(target, "wb") as f:
            for chunk in resp.iter_content(chunk_size=64 * 1024):
                if not chunk:
                    continue
                f.write(chunk)
                downloaded += len(chunk)
                if progress_callback and total > 0:
                    progress_callback(downloaded, total)
    except OSError as e:
        logger.error("写入文件失败: %s -> %s", target, e)
        _safe_remove(target)
        return False

    if expected_size > 0 and downloaded != expected_size:
        logger.error("文件大小不匹配: 期望 %d, 实际 %d", expected_size, downloaded)
        _safe_remove(target)
        return False

    if downloaded == 0:
        logger.error("下载内容为空: %s", url)
        _safe_remove(target)
        return False

    return True


# ── 安全删除 ───────────────────────────────────────────────────────────

def _safe_remove(path: str) -> None:
    """安全删除文件."""
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass


def _safe_rmtree(path: str) -> None:
    """安全删除目录树."""
    try:
        if os.path.isdir(path):
            shutil.rmtree(path)
    except OSError:
        pass


# ── 核心更新流程 ───────────────────────────────────────────────────────

_CHECK_LOCK = Lock()
_IS_CHECKING = False


def is_updating() -> bool:
    """整包更新是否正在进行."""
    with _CHECK_LOCK:
        return _IS_CHECKING


def _do_update(
    progress_callback: Callable[[int, int], None] | None = None,
) -> PackUpdateResult:
    """执行整包更新：查询 → 下载 → 校验 → 解压 → 替换."""
    global _IS_CHECKING

    with _CHECK_LOCK:
        if _IS_CHECKING:
            return PackUpdateResult(False, "整包更新正在进行中...", error_type="duplicate")
        _IS_CHECKING = True

    try:
        session = _make_session()

        # 1. 查询最新 Release
        logger.info("开始查询整包更新...")
        release = _fetch_release_assets(session)
        if not release:
            return PackUpdateResult(
                False, "无法获取 Release 信息（网络问题）", error_type="network",
            )

        remote_tag = release["tag_name"]
        local_ver = _get_local_version()
        if remote_tag == local_ver:
            logger.info("题库已是最新版本: %s", remote_tag)
            return PackUpdateResult(True, "题库已是最新", new_version=remote_tag)

        logger.info("发现新版题库: %s -> %s", local_ver or "(首次)", remote_tag)

        # 2. 下载 zip 到 staging
        os.makedirs(_STAGING_DIR, exist_ok=True)
        zip_path = os.path.join(_STAGING_DIR, _ZIP_FILENAME)
        _safe_remove(zip_path)

        logger.info("开始下载 %s (%d bytes)...", _ZIP_FILENAME, release["zip_size"])
        ok = _download_to_file(
            session, release["zip_url"], zip_path,
            expected_size=release["zip_size"],
            progress_callback=progress_callback,
        )
        if not ok:
            _safe_rmtree(_STAGING_DIR)
            return PackUpdateResult(False, "下载失败", error_type="download")

        logger.info("下载完成: %s", zip_path)

        # 3. SHA256 校验
        if release["hash_url"]:
            logger.info("下载 SHA256 校验文件...")
            hash_path = zip_path + ".sha256"
            hash_ok = _download_to_file(session, release["hash_url"], hash_path)
            if hash_ok:
                with open(hash_path, "r", encoding="utf-8") as f:
                    expected_hash = _parse_sha256_file(f.read())
                actual_hash = _sha256_file(zip_path)
                _safe_remove(hash_path)

                if expected_hash and actual_hash != expected_hash:
                    logger.error(
                        "SHA256 校验失败: 期望 %s..., 实际 %s...",
                        expected_hash[:16], actual_hash[:16],
                    )
                    _safe_rmtree(_STAGING_DIR)
                    return PackUpdateResult(
                        False, "SHA256 校验失败，文件可能损坏", error_type="verify",
                    )
                logger.info("SHA256 校验通过")
            else:
                logger.warning("SHA256 校验文件下载失败，跳过校验")
        else:
            logger.warning("Release 中无 SHA256 校验文件，跳过校验")

        # 4. 解压到 staging
        extract_dir = os.path.join(_STAGING_DIR, "extracted")
        _safe_rmtree(extract_dir)

        logger.info("解压 %s...", _ZIP_FILENAME)
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(extract_dir)
        except (zipfile.BadZipFile, OSError) as e:
            logger.error("解压失败: %s", e)
            _safe_rmtree(_STAGING_DIR)
            return PackUpdateResult(False, f"解压失败: {e}", error_type="extract")

        # 5. 验证解压内容（至少有一个 xlsx）
        xlsx_found = False
        for root, dirs, files in os.walk(extract_dir):
            for f in files:
                if f.endswith(".xlsx"):
                    xlsx_found = True
                    break
            if xlsx_found:
                break

        if not xlsx_found:
            logger.error("解压内容中未找到 xlsx 文件")
            _safe_rmtree(_STAGING_DIR)
            return PackUpdateResult(
                False, "压缩包中未找到试卷文件", error_type="verify",
            )

        # 6. 原子替换：清缓存 → 备份旧目录 → 移入新目录
        logger.info("原子替换题库目录...")
        # 释放 pandas/openpyxl 可能持有的文件句柄，避免 Windows 文件锁
        import gc
        gc.collect()
        _safe_rmtree(_BACKUP_DIR)

        if os.path.isdir(QUESTION_BANK_FILES_DIR):
            try:
                os.replace(QUESTION_BANK_FILES_DIR, _BACKUP_DIR)
            except OSError as e:
                logger.error("备份旧题库失败: %s", e)
                _safe_rmtree(_STAGING_DIR)
                return PackUpdateResult(
                    False, f"备份旧题库失败: {e}", error_type="unknown",
                )

        # 将解压内容移到 files/
        # zip 内可能有一层根目录（如 question_bank/），也可能是直接的文件
        # 检测逻辑：如果解压目录下只有一个子目录且包含 xlsx，则用该子目录
        source_dir = extract_dir
        entries = os.listdir(extract_dir)
        if len(entries) == 1 and os.path.isdir(os.path.join(extract_dir, entries[0])):
            candidate = os.path.join(extract_dir, entries[0])
            candidate_xlsx = any(
                f.endswith(".xlsx") for f in os.listdir(candidate)
                if os.path.isfile(os.path.join(candidate, f))
            )
            if candidate_xlsx:
                source_dir = candidate

        try:
            os.replace(source_dir, QUESTION_BANK_FILES_DIR)
        except OSError as e:
            logger.error("替换题库目录失败，尝试回滚: %s", e)
            # 回滚：恢复备份
            if os.path.isdir(_BACKUP_DIR):
                try:
                    os.replace(_BACKUP_DIR, QUESTION_BANK_FILES_DIR)
                    logger.info("回滚成功")
                except OSError:
                    logger.error("回滚失败！题库可能不可用")
            _safe_rmtree(_STAGING_DIR)
            return PackUpdateResult(False, f"替换失败: {e}", error_type="unknown")

        # 7. 清理：强制删除旧备份和暂存目录
        gc.collect()
        _safe_rmtree(_BACKUP_DIR)
        _safe_rmtree(_STAGING_DIR)

        # 8. 保存版本号
        _save_local_version(remote_tag)

        # 9. 更新 manifest.json（供 question_bank_updater 增量更新使用）
        _regenerate_manifest(QUESTION_BANK_FILES_DIR)

        logger.info("题库整包更新完成: %s", remote_tag)
        return PackUpdateResult(True, f"题库已更新至 {remote_tag}", new_version=remote_tag)

    except Exception as e:
        logger.error("整包更新异常: %s", e, exc_info=True)
        _safe_rmtree(_STAGING_DIR)
        # 尝试回滚
        if os.path.isdir(_BACKUP_DIR):
            try:
                if not os.path.isdir(QUESTION_BANK_FILES_DIR):
                    os.replace(_BACKUP_DIR, QUESTION_BANK_FILES_DIR)
                    logger.info("异常回滚成功")
            except OSError:
                pass
        return PackUpdateResult(False, f"更新失败: {e}", error_type="unknown")
    finally:
        _IS_CHECKING = False


def _regenerate_manifest(files_dir: str) -> None:
    """更新后重新生成本地 manifest.json（简化版，只记录文件列表）."""
    from hnust_exam.utils.constants import MANIFEST_FILE
    try:
        files_info = {}
        for fname in sorted(os.listdir(files_dir)):
            fpath = os.path.join(files_dir, fname)
            if os.path.isfile(fpath) and fname.endswith(".xlsx"):
                files_info[fname] = {
                    "hash": _sha256_file(fpath),
                    "size": os.path.getsize(fpath),
                }
            elif os.path.isdir(fpath):
                for sub in sorted(os.listdir(fpath)):
                    subpath = os.path.join(fpath, sub)
                    if os.path.isfile(subpath):
                        key = f"{fname}/{sub}"
                        files_info[key] = {
                            "hash": _sha256_file(subpath),
                            "size": os.path.getsize(subpath),
                        }

        manifest = {"version": 1, "files": files_info}
        os.makedirs(os.path.dirname(MANIFEST_FILE), exist_ok=True)
        tmp = MANIFEST_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        os.replace(tmp, MANIFEST_FILE)
        logger.info("manifest.json 已更新（%d 个文件）", len(files_info))
    except Exception as e:
        logger.error("更新 manifest 失败: %s", e)


# ── Qt 信号 ────────────────────────────────────────────────────────────

class _PackUpdateSignal(QObject):
    """将更新结果派发到主线程."""
    result = Signal(object)
    progress = Signal(int, int)  # downloaded, total


# ── 公开入口 ───────────────────────────────────────────────────────────

def check_pack_update_async(
    callback: Callable[[PackUpdateResult], None] | None = None,
    progress_callback: Callable[[int, int], None] | None = None,
) -> None:
    """异步检查并执行整包更新（后台线程，幂等安全）.

    Parameters
    ----------
    callback : 可选回调，接收 PackUpdateResult（通过 Qt Signal 派发到主线程）。
    progress_callback : 可选进度回调，接收 (downloaded_bytes, total_bytes)。
    """
    sig: _PackUpdateSignal | None = None
    if callback is not None:
        sig = _PackUpdateSignal()
        sig.result.connect(callback, Qt.ConnectionType.QueuedConnection)

    def _worker() -> None:
        try:
            result = _do_update(progress_callback=progress_callback)
            if sig is not None:
                sig.result.emit(result)
        except Exception as e:
            logger.error("整包更新线程崩溃: %s", e, exc_info=True)
            if sig is not None:
                sig.result.emit(
                    PackUpdateResult(False, f"更新失败: {e}", error_type="unknown"),
                )

    Thread(target=_worker, daemon=True).start()
