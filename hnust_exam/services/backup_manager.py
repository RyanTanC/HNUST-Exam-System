"""备份管理服务：程序文件的备份、恢复和清理."""

from __future__ import annotations

import logging
import os
import shutil
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from hnust_exam.models.exam import Exam

logger = logging.getLogger(__name__)


def _validate_program_file_path(program_file: str) -> str:
    """校验程序文件名，拒绝绝对路径和目录穿越."""
    normalized = os.path.normpath(program_file.strip())
    drive, _ = os.path.splitdrive(normalized)
    if not normalized or normalized in (".", os.pardir):
        raise ValueError("不允许的程序文件路径")
    if drive or os.path.isabs(normalized):
        raise ValueError("不允许的程序文件路径")
    parts = normalized.replace("/", os.sep).split(os.sep)
    if os.pardir in parts:
        raise ValueError("不允许的程序文件路径")
    return normalized


class BackupManager:
    """管理程序文件的备份和恢复."""

    def __init__(self) -> None:
        self._backup_dir: str | None = None

    def init_backup(
        self, exam_file_path: str, exam: Exam
    ) -> None:
        """初始化备份：将所有引用的程序文件备份到 _backup_programs 目录."""
        from hnust_exam.services.resource_manager import get_program_dir, get_question_bank_root

        source_dir = get_program_dir()
        if not os.path.isdir(source_dir):
            source_dir = get_question_bank_root()

        self._backup_dir = os.path.join(get_question_bank_root(), "_backup_programs")
        if os.path.exists(self._backup_dir):
            try:
                shutil.rmtree(self._backup_dir)
            except Exception as e:
                logger.warning("重置备份目录失败 %s: %s", self._backup_dir, e)
        os.makedirs(self._backup_dir, exist_ok=True)

        referenced_files: set[str] = set()
        for q in exam.questions:
            raw_program_file = q.program_file.strip()
            if raw_program_file:
                referenced_files.add(_validate_program_file_path(raw_program_file))

        for pf in referenced_files:
            src = os.path.join(source_dir, pf)
            if os.path.isfile(src):
                dst = os.path.join(self._backup_dir, pf)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(src, dst)

    def restore_file(self, program_file: str, exam_file_path: str) -> None:
        """从备份恢复单个程序文件."""
        if not self._backup_dir or not os.path.exists(self._backup_dir):
            return

        program_file = _validate_program_file_path(program_file)
        backup_file = os.path.join(self._backup_dir, program_file)
        if not os.path.exists(backup_file):
            return

        from hnust_exam.services.resource_manager import get_program_dir
        target_path = os.path.join(get_program_dir(), program_file)
        os.makedirs(os.path.dirname(target_path), exist_ok=True)

        try:
            shutil.copy2(backup_file, target_path)
        except Exception as e:
            logger.warning("恢复程序文件失败 %s: %s", program_file, e)

    def cleanup(self) -> None:
        """清理备份目录."""
        if self._backup_dir and os.path.exists(self._backup_dir):
            try:
                shutil.rmtree(self._backup_dir)
            except Exception as e:
                logger.warning("清理备份目录失败 %s: %s", self._backup_dir, e)
        self._backup_dir = None
