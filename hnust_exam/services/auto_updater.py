"""自动更新模块：下载新版本 exe → 校验 → 替换 → 重启."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import time

import requests


def get_app_exe_path() -> str:
    """获取当前应用 exe 的真实路径.

    - frozen（PyInstaller 打包）：返回用户双击的那个原始 exe 的绝对路径
    - 开发环境：返回 Python 解释器路径
    """
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.argv[0])
    return sys.executable


def check_write_permission(path: str) -> bool:
    """检查指定路径所在目录是否可写."""
    return os.access(os.path.dirname(path), os.W_OK)


def get_temp_download_path() -> str:
    """获取临时下载路径，避免污染安装目录."""
    return os.path.join(tempfile.gettempdir(), "HNUST_update_tmp.exe")


def download_file(
    url: str,
    save_path: str,
    progress_callback=None,
    expected_size: int = 0,
) -> tuple[bool, str]:
    """流式下载文件到临时路径，带进度回调.

    Args:
        url: 下载地址
        save_path: 保存路径（临时目录）
        progress_callback: 回调函数 (percent, speed_mb, remaining_sec)
        expected_size: 预期文件大小（字节），用于完整性校验

    Returns:
        (success, error_msg)
    """
    try:
        resp = requests.get(url, stream=True, timeout=(10, 60))
        resp.raise_for_status()
    except requests.RequestException as e:
        return False, f"无法连接服务器：{e}"

    total_length = 0
    content_length = resp.headers.get("content-length")
    if content_length:
        total_length = int(content_length)
    elif expected_size > 0:
        total_length = expected_size

    start_time = time.time()
    downloaded = 0

    try:
        with open(save_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=64 * 1024):
                if not chunk:
                    continue
                f.write(chunk)
                downloaded += len(chunk)

                if progress_callback and total_length > 0:
                    percent = min(int(downloaded * 100 / total_length), 100)
                    elapsed = time.time() - start_time
                    if elapsed > 0.5 and downloaded > 0:
                        speed_mb = round(downloaded / elapsed / 1024 / 1024, 2)
                        remaining = min(
                            int((total_length - downloaded) / (downloaded / elapsed)),
                            3599,  # 上限 99:59
                        )
                    else:
                        speed_mb = 0.0
                        remaining = 0
                    progress_callback(percent, speed_mb, remaining)
    except OSError as e:
        _cleanup_temp(save_path)
        return False, f"写入文件失败：{e}"

    # 完整性校验
    if expected_size > 0:
        actual_size = os.path.getsize(save_path)
        if actual_size != expected_size:
            _cleanup_temp(save_path)
            return False, f"文件大小不匹配（预期 {expected_size}，实际 {actual_size}）"

    if downloaded == 0:
        _cleanup_temp(save_path)
        return False, "下载内容为空"

    return True, ""


def replace_and_restart(new_exe_path: str) -> tuple[bool, str]:
    """用新 exe 替换当前 exe 并重启（仅 frozen 环境）.

    流程：
    1. 将当前 exe 重命名为 .old.bak
    2. 将新 exe 移动到当前 exe 位置
    3. 启动新 exe，退出当前进程

    失败时自动回滚。

    Returns:
        (success, error_msg) — 成功时不会返回（进程已退出）
    """
    if not getattr(sys, "frozen", False):
        return False, "非打包环境，无法自动替换"

    exe_path = get_app_exe_path()
    backup_path = exe_path + ".old.bak"

    # 检查写入权限
    if not check_write_permission(exe_path):
        return False, "权限不足，请以管理员身份运行或将程序移至非系统目录"

    # 步骤1：备份当前 exe
    try:
        if os.path.exists(backup_path):
            os.remove(backup_path)
        os.replace(exe_path, backup_path)
    except OSError as e:
        return False, f"无法备份当前版本：{e}"

    # 步骤2：移动新 exe 到原位置
    try:
        shutil.move(new_exe_path, exe_path)
    except OSError as e:
        # 回滚：恢复备份
        try:
            os.replace(backup_path, exe_path)
        except OSError:
            pass
        return False, f"替换文件失败（可能被杀毒软件拦截）：{e}"

    # 步骤3：启动新版本并退出
    try:
        subprocess.Popen([exe_path], close_fds=True)
    except OSError as e:
        # 启动失败，回滚：将备份恢复到原位置
        try:
            os.replace(backup_path, exe_path)
        except OSError:
            pass
        return False, f"启动新版本失败（已回滚）：{e}"

    sys.exit(0)


def clean_old_backup() -> None:
    """启动时清理上次更新留下的 .old.bak 残留文件."""
    if not getattr(sys, "frozen", False):
        return
    exe_path = get_app_exe_path()
    backup_path = exe_path + ".old.bak"
    if os.path.exists(backup_path):
        try:
            os.remove(backup_path)
        except OSError:
            pass


def _cleanup_temp(path: str) -> None:
    """清理临时文件."""
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass
