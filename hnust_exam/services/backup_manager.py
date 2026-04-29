"""备份管理服务：程序文件的备份、恢复和清理."""

from __future__ import annotations

import os
import shutil
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from hnust_exam.models.exam import Exam


class BackupManager:
    """管理程序文件的备份和恢复."""

    def __init__(self) -> None:
        self._backup_dir: str | None = None

    def init_backup(
        self, exam_file_path: str, exam: Exam
    ) -> None:
        """初始化备份：将所有引用的程序文件备份到 _backup_programs 目录."""
        exam_dir = os.path.dirname(exam_file_path)
        source_dir = os.path.join(exam_dir, "试题文件夹")
        if not os.path.exists(source_dir):
            source_dir = exam_dir

        self._backup_dir = os.path.join(exam_dir, "_backup_programs")
        if os.path.exists(self._backup_dir):
            shutil.rmtree(self._backup_dir)
        os.makedirs(self._backup_dir, exist_ok=True)

        referenced_files: set[str] = set()
        for q in exam.questions:
            pf = q.program_file.strip()
            if pf:
                referenced_files.add(pf)

        for pf in referenced_files:
            src = os.path.join(source_dir, pf)
            if os.path.isfile(src):
                shutil.copy2(src, os.path.join(self._backup_dir, pf))

    def restore_file(self, program_file: str, exam_file_path: str) -> None:
        """从备份恢复单个程序文件."""
        if not self._backup_dir or not os.path.exists(self._backup_dir):
            return

        backup_file = os.path.join(self._backup_dir, program_file)
        if not os.path.exists(backup_file):
            return

        exam_dir = os.path.dirname(exam_file_path)
        target_path = os.path.join(exam_dir, "试题文件夹", program_file)
        if not os.path.exists(target_path):
            target_path = os.path.join(exam_dir, program_file)

        try:
            shutil.copy2(backup_file, target_path)
        except Exception:
            pass

    def cleanup(self) -> None:
        """清理备份目录."""
        if self._backup_dir and os.path.exists(self._backup_dir):
            try:
                shutil.rmtree(self._backup_dir)
            except Exception:
                pass
        self._backup_dir = None
