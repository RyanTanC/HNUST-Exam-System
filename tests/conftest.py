"""pytest 全局配置：把用户数据目录重定向到临时目录.

``~/.hnust_exam`` 里放的是用户真实的配置、答题进度和题库缓存。历史上已经出现过
测试直接往里写、把假数据盖到真实文件上的事故（配置被重置、题库 manifest 被污染）。
这里统一加一道护栏：任何用例只要通过 ``hnust_exam.utils.constants`` 或各服务模块
里的派生常量访问文件，都会被引到临时目录，跑测试不可能再动到用户数据。

需要更细粒度控制的用例，可以在这之上再 monkeypatch 一层，互不冲突。
"""

from __future__ import annotations

import os
import tempfile

import pytest

from hnust_exam.services import (
    question_bank_updater,
    resource_pack_updater,
    update_lock,
)
from hnust_exam.utils import constants


@pytest.hookimpl(tryfirst=True)
def pytest_configure(config) -> None:
    """给 pytest 一个**固定**的 basetemp，别用 numbered 临时目录.

    真实踩到的坑：pytest 退出时会在 ``atexit`` 里清理历史 numbered 临时目录
    （``cleanup_numbered_dir`` → ``shutil.rmtree``）。某些环境装了**批量删除
    安全钩子**，会把这个清理判定为可疑批量删除、直接拒绝并 ``raise SystemExit(1)``。
    表现是**测试全绿但退出码变成 1**：

        273 passed, 1 skipped in 46.72s
        [safe-delete][SAFE_DELETE_BULK_CONFIRM_REQUIRED] {...}
        Exception ignored in atexit callback <function cleanup_numbered_dir ...>
        SystemExit: 1

    很容易被误判成"某两个用例偶发失败"。指定固定 basetemp 后 pytest 不再创建
    numbered 目录，atexit 那段清理自然也不会跑。
    """
    if not config.option.basetemp:
        config.option.basetemp = os.path.join(
            tempfile.gettempdir(), "_hnust_pytest_base"
        )


@pytest.fixture(autouse=True)
def isolate_user_data(tmp_path_factory, monkeypatch):
    """把用户数据路径全部指向临时目录（每个用例一份，互不干扰）."""
    root = tmp_path_factory.mktemp("hnust_userdata")
    qb = root / "question_bank"

    # constants 里的路径常量（config_manager / telemetry / app 直接引用）
    monkeypatch.setattr(constants, "_CONFIG_DIR", str(root))
    monkeypatch.setattr(constants, "CONFIG_FILE", str(root / "config.json"))
    monkeypatch.setattr(constants, "PROGRESS_FILE", str(root / "progress.json"))
    monkeypatch.setattr(constants, "SKIP_VERSION_FILE", str(root / "skip_ver"))
    monkeypatch.setattr(
        constants, "TELEMETRY_QUEUE_FILE", str(root / "telemetry_queue.json")
    )
    monkeypatch.setattr(constants, "QUESTION_BANK_DIR", str(qb))
    monkeypatch.setattr(constants, "QUESTION_BANK_FILES_DIR", str(qb / "files"))
    monkeypatch.setattr(constants, "MANIFEST_FILE", str(qb / "manifest.json"))
    monkeypatch.setattr(constants, "LOG_DIR", str(root / "logs"))

    # 下面这些是在模块导入时从 QUESTION_BANK_DIR 算出来的，改 constants 已经来不及
    monkeypatch.setattr(resource_pack_updater, "_BACKUP_DIR", str(qb / "files_backup"))
    monkeypatch.setattr(
        resource_pack_updater, "_CURRENT_VERSION_FILE", str(qb / "current_version")
    )
    monkeypatch.setattr(update_lock, "_LOCK_DIR", str(qb / ".locks"))
    monkeypatch.setattr(question_bank_updater, "_LAST_CHECK_FILE", str(qb / "last_check"))

    yield root
