"""跨进程文件锁管理.

使用 filelock 实现，确保同一时间只有一个进程执行更新操作。
支持过期锁检测（进程崩溃后残留的锁文件）。
"""

from __future__ import annotations

import logging
import os
import time
from typing import Optional

from filelock import FileLock, Timeout

from hnust_exam.utils.constants import QUESTION_BANK_DIR

logger = logging.getLogger(__name__)

# 锁文件目录
_LOCK_DIR = os.path.join(QUESTION_BANK_DIR, ".locks")


class UpdateLock:
    """跨进程更新锁.

    Parameters
    ----------
    resource_name : str
        资源名称（如 "question_bank_update"），用于生成唯一的锁文件路径
    stale_timeout : int
        过期锁超时时间（秒）。如果锁文件的修改时间超过此值，认为是过期锁。
        默认 300 秒（5 分钟）
    """

    def __init__(
        self,
        resource_name: str,
        stale_timeout: float = 300,
        timeout: float = 60,
    ) -> None:
        self._resource_name = resource_name
        self._stale_timeout = stale_timeout
        self._acquire_timeout = timeout
        self._lock_path = os.path.join(
            _LOCK_DIR, f"{resource_name}.lock"
        )
        self._lock: Optional[FileLock] = None

    def acquire(
        self,
        timeout: float = 60,
        blocking: bool = True,
    ) -> bool:
        """获取锁.

        Parameters
        ----------
        timeout : float
            等待锁的超时时间（秒）
        blocking : bool
            是否阻塞等待。False 时立即返回。

        Returns
        -------
        bool
            是否成功获取锁
        """
        os.makedirs(_LOCK_DIR, exist_ok=True)
        self._cleanup_stale_lock()
        self._lock = FileLock(self._lock_path)

        try:
            if blocking:
                self._lock.acquire(timeout=timeout)
            else:
                self._lock.acquire(timeout=0)
            return True
        except Timeout:
            logger.warning(
                "获取更新锁超时: %s (timeout=%s, blocking=%s)",
                self._resource_name, timeout, blocking,
            )
            return False

    def release(self) -> None:
        """释放锁."""
        if self._lock is not None:
            try:
                self._lock.release()
            except Exception as e:
                logger.warning("释放更新锁失败: %s -> %s", self._resource_name, e)
            self._lock = None

    def is_locked(self) -> bool:
        """检查锁是否被持有."""
        return self._lock is not None and self._lock.is_locked

    def _cleanup_stale_lock(self) -> None:
        """清理过期锁文件."""
        if not os.path.exists(self._lock_path):
            return

        try:
            mtime = os.path.getmtime(self._lock_path)
            age = time.time() - mtime
            if age > self._stale_timeout:
                logger.info(
                    "检测到过期锁文件 (age=%.0fs): %s, 尝试清理",
                    age, self._lock_path,
                )
                os.remove(self._lock_path)
        except OSError as e:
            logger.warning("清理过期锁文件失败: %s -> %s", self._lock_path, e)

    def __enter__(self) -> UpdateLock:
        if not self.acquire(timeout=self._acquire_timeout):
            raise RuntimeError(
                f"无法获取更新锁: {self._resource_name}"
            )
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.release()
        return False

    def __del__(self) -> None:
        try:
            self.release()
        except Exception:
            pass
