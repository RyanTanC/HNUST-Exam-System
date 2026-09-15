"""考试页判分设置测试."""

from __future__ import annotations

from hnust_exam.views.exam_page import ExamPage


class _ConfigManager:
    def load_config(self):
        return {"grading_strictness": "strict"}


class _MainWindow:
    config_mgr = _ConfigManager()


class _Exam:
    def __init__(self):
        self.strictness = None

    def grade(self, strictness="normal"):
        self.strictness = strictness
        return []


def test_grade_exam_uses_configured_strictness():
    page = ExamPage.__new__(ExamPage)
    page.main_window = _MainWindow()
    page.exam = _Exam()

    assert page._grade_exam() == []
    assert page.exam.strictness == "strict"
