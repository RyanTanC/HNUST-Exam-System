"""备份管理服务：程序文件的备份、恢复和清理."""

from __future__ import annotations

import os
import shutil
import logging
import tempfile
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from hnust_exam.models.exam import Exam

logger = logging.getLogger(__name__)


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

        self._backup_dir = tempfile.mkdtemp(prefix="_backup_programs_", dir=exam_dir)

        referenced_files: set[str] = set()
        for q in exam.questions:
            pf = q.program_file.strip()
            if pf:
                referenced_files.add(pf)

        for pf in referenced_files:
            src = self.resolve_program_path(pf, exam_dir, must_exist=True)
            if os.path.isfile(src):
                dst = self._safe_join(self._backup_dir, pf)
                if dst:
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.copy2(src, dst)

    @staticmethod
    def is_unsafe_program_path(program_file: str) -> bool:
        normalized = os.path.normpath(program_file.strip())
        parts = normalized.split(os.sep)
        return (
            not normalized
            or os.path.isabs(normalized)
            or normalized.startswith("..")
            or any(part == ".." for part in parts)
        )

    @staticmethod
    def _safe_join(root: str, relative_path: str) -> str | None:
        if BackupManager.is_unsafe_program_path(relative_path):
            return None
        root_abs = os.path.abspath(root)
        candidate = os.path.abspath(os.path.join(root_abs, os.path.normpath(relative_path)))
        if os.path.commonpath([root_abs, candidate]) != root_abs:
            return None
        return candidate

    def resolve_program_path(
        self,
        program_file: str,
        exam_dir: str,
        *,
        must_exist: bool = False,
    ) -> str | None:
        if self.is_unsafe_program_path(program_file):
            return None
        roots = [os.path.join(exam_dir, "试题文件夹"), exam_dir]
        for root in roots:
            candidate = self._safe_join(root, program_file)
            if candidate and (not must_exist or os.path.exists(candidate)):
                return candidate
        return None

    def restore_file(self, program_file: str, exam_file_path: str) -> None:
        """从备份恢复单个程序文件."""
        if not self._backup_dir or not os.path.exists(self._backup_dir):
            return

        backup_file = self._safe_join(self._backup_dir, program_file)
        if not backup_file:
            logger.warning("拒绝恢复不安全的程序文件路径: %s", program_file)
            return
        if not os.path.exists(backup_file):
            return

        exam_dir = os.path.dirname(exam_file_path)
        target_path = self.resolve_program_path(program_file, exam_dir)
        if not target_path:
            logger.warning("拒绝写入不安全的程序文件路径: %s", program_file)
            return

        try:
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            shutil.copy2(backup_file, target_path)
        except Exception:
            logger.exception("恢复程序文件失败: %s", program_file)

    def cleanup(self) -> None:
        """释放当前备份引用，备份目录请确认后手动清理."""
        if self._backup_dir and os.path.exists(self._backup_dir):
            logger.info("保留备份目录，请确认后手动清理: %s", self._backup_dir)
        self._backup_dir = None
