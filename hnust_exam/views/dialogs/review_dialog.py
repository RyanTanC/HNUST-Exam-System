"""题目回顾对话框."""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QFrame,
    QWidget,
)

from hnust_exam.models.result import Result
from hnust_exam.utils.theme import Theme

if TYPE_CHECKING:
    from hnust_exam.models.exam import Exam


class ReviewDialog(QDialog):
    """单题回顾对话框."""

    def __init__(self, result: Result, exam: Exam | None, parent=None) -> None:
        super().__init__(parent)
        self._result = result
        self._exam = exam
        self.setWindowTitle(f"题目回顾 - {result.question_number}")
        self.setMinimumSize(500, 350)
        self.resize(700, 500)
        self._build_ui()

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 顶部
        header = QFrame()
        header.setStyleSheet(f"background-color: {c['PRIMARY']}; padding: 10px 20px;")
        icon_txt = "✓" if self._result.is_correct else "✗"
        title = QLabel(
            f"{icon_txt}  {self._result.q_type} - 题号 {self._result.question_number} - {self._result.score}分"
        )
        title.setStyleSheet("color: white; font-size: 12pt; font-weight: bold;")
        header_layout = QHBoxLayout(header)
        header_layout.addWidget(title)
        layout.addWidget(header)

        # 内容滚动区
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        content = QWidget()
        content.setStyleSheet(f"background-color: {c['WHITE']}; padding: 20px;")
        content_layout = QVBoxLayout(content)

        # 题目
        q_title = QLabel("题目")
        q_title.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold; padding-bottom: 5px;")
        content_layout.addWidget(q_title)

        q_text = QLabel(self._result.question_text or "暂无")
        q_text.setWordWrap(True)
        q_text.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; padding-bottom: 15px;")
        content_layout.addWidget(q_text)

        # 选项（如果是单选题）
        if self._result.q_type == "单选" and self._exam:
            for q in self._exam.questions:
                if q.number == self._result.question_number and q.options:
                    opt_title = QLabel("选项：")
                    opt_title.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold; padding-bottom: 5px;")
                    content_layout.addWidget(opt_title)
                    for letter, text in q.options.items():
                        opt = QLabel(f"  ({letter}) {text}")
                        opt.setWordWrap(True)
                        opt.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; padding: 2px 20px;")
                        content_layout.addWidget(opt)
                    break

        # 答案区
        color = c["SUCCESS"] if self._result.is_correct else c["DANGER"]
        ans_frame = QFrame()
        ans_frame.setStyleSheet(
            f"background-color: {c['SURFACE']}; border: 1px solid {c['BORDER']}; border-radius: 4px;"
        )
        ans_layout = QVBoxLayout(ans_frame)

        user_ans = QLabel(f"你的答案：{self._result.user_answer}")
        user_ans.setWordWrap(True)
        user_ans.setStyleSheet(f"color: {color}; font-size: 11pt; font-weight: bold; padding: 8px 15px;")
        ans_layout.addWidget(user_ans)

        correct_ans = QLabel(f"正确答案：{self._result.correct_answer}")
        correct_ans.setWordWrap(True)
        correct_ans.setStyleSheet(
            f"color: {c['SUCCESS']}; font-size: 11pt; font-weight: bold; "
            f"font-family: Consolas; padding: 0 15px 8px 15px;"
        )
        ans_layout.addWidget(correct_ans)

        content_layout.addWidget(ans_frame)
        content_layout.addStretch()

        scroll.setWidget(content)
        layout.addWidget(scroll, 1)

        # 底部
        btn_frame = QFrame()
        btn_frame.setStyleSheet(f"background-color: {c['BG']}; padding: 10px;")
        btn_layout = QHBoxLayout(btn_frame)
        btn_layout.addStretch()
        close_btn = QPushButton("关闭")
        close_btn.setStyleSheet(
            f"background-color: {c['BG']}; color: {c['TEXT']}; font-size: 11pt; "
            f"padding: 5px 20px; border: none; border-radius: 4px;"
        )
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)
        layout.addWidget(btn_frame)
