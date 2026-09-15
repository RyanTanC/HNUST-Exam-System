"""备份管理安全测试."""

from __future__ import annotations

from dataclasses import dataclass

import pytest

from hnust_exam.services.backup_manager import BackupManager


@dataclass
class _Question:
    program_file: str


@dataclass
class _Exam:
    questions: list[_Question]


def test_init_backup_rejects_parent_traversal(tmp_path, monkeypatch):
    root = tmp_path / "bank"
    program_dir = root / "试题文件夹"
    program_dir.mkdir(parents=True)

    monkeypatch.setattr("hnust_exam.services.resource_manager.get_question_bank_root", lambda: str(root))
    monkeypatch.setattr("hnust_exam.services.resource_manager.get_program_dir", lambda: str(program_dir))

    exam = _Exam([_Question("../config.json")])

    with pytest.raises(ValueError, match="不允许的程序文件路径"):
        BackupManager().init_backup("exam.xlsx", exam)


def test_restore_file_rejects_absolute_target(tmp_path, monkeypatch):
    root = tmp_path / "bank"
    program_dir = root / "试题文件夹"
    program_dir.mkdir(parents=True)

    monkeypatch.setattr("hnust_exam.services.resource_manager.get_program_dir", lambda: str(program_dir))

    manager = BackupManager()
    manager._backup_dir = str(tmp_path / "backup")
    (tmp_path / "backup").mkdir()
    with pytest.raises(ValueError, match="不允许的程序文件路径"):
        manager.restore_file(str(tmp_path / "evil.py"), "exam.xlsx")
