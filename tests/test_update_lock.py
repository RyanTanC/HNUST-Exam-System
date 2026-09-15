"""测试跨进程文件锁."""
import os
import tempfile
from unittest.mock import patch

from hnust_exam.services.update_lock import UpdateLock, _LOCK_DIR as ORIG_LOCK_DIR


def test_update_lock_acquire_release():
    """测试锁的获取和释放."""
    with tempfile.TemporaryDirectory() as tmpdir:
        with patch("hnust_exam.services.update_lock._LOCK_DIR", tmpdir):
            lock = UpdateLock("test_resource")
            assert lock.acquire(timeout=5) is True
            assert lock.is_locked() is True
            lock.release()
            assert lock.is_locked() is False


def test_update_lock_context_manager():
    """测试上下文管理器."""
    with tempfile.TemporaryDirectory() as tmpdir:
        with patch("hnust_exam.services.update_lock._LOCK_DIR", tmpdir):
            with UpdateLock("test_resource") as lock:
                assert lock.is_locked() is True
            assert lock.is_locked() is False


def test_update_lock_non_blocking():
    """测试非阻塞获取."""
    with tempfile.TemporaryDirectory() as tmpdir:
        with patch("hnust_exam.services.update_lock._LOCK_DIR", tmpdir):
            lock1 = UpdateLock("test_resource")
            lock2 = UpdateLock("test_resource")

            assert lock1.acquire(timeout=5) is True
            assert lock2.acquire(blocking=False) is False

            lock1.release()
            assert lock2.acquire(blocking=False) is True
            lock2.release()


def test_update_lock_context_manager_fail_raises():
    """测试上下文管理器在获取锁失败时抛出异常."""
    with tempfile.TemporaryDirectory() as tmpdir:
        with patch("hnust_exam.services.update_lock._LOCK_DIR", tmpdir):
            lock1 = UpdateLock("test_resource")
            lock1.acquire(timeout=5)

            import pytest
            with pytest.raises(RuntimeError, match="无法获取更新锁"):
                with UpdateLock("test_resource", timeout=1) as lock:
                    pass  # 不应该进入这里

            lock1.release()


def test_update_lock_stale_cleanup():
    """测试过期锁文件被自动清理."""
    with tempfile.TemporaryDirectory() as tmpdir:
        with patch("hnust_exam.services.update_lock._LOCK_DIR", tmpdir):
            lock_path = os.path.join(tmpdir, "test_resource.lock")

            # 创建一个旧的锁文件
            with open(lock_path, "w") as f:
                f.write("stale")

            # 设置文件的修改时间为很久以前
            old_time = 1000000  # 1970年代的某个时间
            os.utime(lock_path, (old_time, old_time))

            # 获取锁时会触发过期检测并清理
            lock = UpdateLock("test_resource", stale_timeout=10)
            assert lock.acquire(timeout=5) is True
            # 过期锁文件应先被清理，然后 filelock 重新创建锁文件
            assert os.path.exists(lock_path)
            lock.release()
