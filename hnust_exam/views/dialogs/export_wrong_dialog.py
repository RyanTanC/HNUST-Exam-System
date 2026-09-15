"""导出错题对话框：预览格式化文本并复制到剪贴板."""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QFrame,
    QMessageBox,
)

from hnust_exam.utils.theme import Theme
from hnust_exam.utils.ui_helpers import themed_question

if TYPE_CHECKING:
    from hnust_exam.models.exam import Exam
    from hnust_exam.models.result import Result


_AI_PROMPT_PREFIX = (
    "你是一位有 10 年编程教学经验的编程导师，擅长通过错题精准诊断学生的认知误区，"
    "并用启发式对话帮助学生真正理解。\n\n"
    "现在学生递交了错题记录，请严格按照以下框架完成教学辅导，语言要亲切、耐心：\n\n"
    "【1. 错题重放与自我觉察】\n"
    "- 复述原题，问 2 个引导性问题，帮学生暴露当时的思考过程\n"
    "- 给出最常见的错误推理路径，问学生是否这样想的，然后分析问题\n\n"
    "【2. 认知误区诊断】\n"
    "- 判断学生最可能属于哪类错误（可多选）：概念混淆型/字面量识记不清/"
    "对函数返回值理解有误/粗心/知识负迁移\n"
    "- 用比喻或生活化例子纠正误区（如列表比作购物清单，元组比作不能涂改的收据）\n"
    "- 指出不解决此误区未来还哪些知识点会反复出错\n\n"
    "【3. 实验式学习活动】\n"
    "- 设计 1-2 个 Python 交互小实验，让学生先猜输出再运行验证\n"
    "- 要求完成填空总结\n\n"
    "【4. 分层练习（最近发展区）】\n"
    "- 4 道练习题：基础题×1、变式题×1、易混题×1、综合应用题×1\n"
    "- 每道题写明考查点，详细解析中穿插反问\n\n"
    "【5. 元认知反思】\n"
    "- 让学生用自己的话回答关键区别\n"
    "- 提供避坑口诀或对比表格\n"
    "- 用鼓励的话结束，告诉学生：如果这几道练习全对，这个知识坑就算填平了\n\n"
    "以下是学生的题目：\n"
)


class ExportWrongDialog(QDialog):
    """导出错题预览对话框."""

    def __init__(
        self,
        results: list[Result],
        exam: Exam,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._results = results
        self._exam = exam
        self.setWindowTitle("导出错题")
        self.setMinimumSize(560, 420)
        self.resize(640, 520)
        self._text = self._build_text()
        self._build_ui()

    # ─────────────── 文本生成 ───────────────

    def _build_text(self) -> str:
        """生成格式化的错题和标记文本."""
        q_map = {q.number: q for q in self._exam.questions}
        sections: list[str] = []

        wrong_results = [r for r in self._results if not r.is_correct]
        if wrong_results:
            lines = ["我的错题有："]
            for r in wrong_results:
                lines.append(self._format_question(r, q_map))
            sections.append("\n".join(lines))

        marked_indices = self._exam.marked_indices
        if marked_indices:
            marked_qs = [
                self._results[i]
                for i in sorted(marked_indices)
                if i < len(self._results)
            ]
            if marked_qs:
                lines = ["我的标记的是："]
                for r in marked_qs:
                    lines.append(self._format_question(r, q_map))
                sections.append("\n".join(lines))

        return "\n\n".join(sections)

    def _format_question(self, r: Result, q_map: dict) -> str:
        """格式化单道题目."""
        parts: list[str] = []
        parts.append(f"题目：{r.question_number}.")

        q = q_map.get(r.question_number)
        if q:
            parts.append(q.text.strip())
            if q.options:
                for letter in ("A", "B", "C", "D", "E", "F"):
                    if letter in q.options:
                        parts.append(f"{letter}. {q.options[letter]}")
        elif r.question_text:
            parts.append(r.question_text.strip())

        parts.append(f"题型：{r.q_type}")
        parts.append(f"参考答案：{r.correct_answer}")
        parts.append(f"我的选择：{r.user_answer}")

        return "\n".join(parts)

    # ─────────────── UI 构建 ───────────────

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 顶部标题栏
        header = QFrame()
        header.setStyleSheet(f"background:{c['PRIMARY']};")
        header_lay = QHBoxLayout(header)
        header_lay.setContentsMargins(20, 12, 20, 12)
        title = QLabel("导出错题")
        title.setStyleSheet("color:#fff;font-size:14pt;font-weight:bold;")
        header_lay.addWidget(title)
        header_lay.addStretch()
        layout.addWidget(header)

        # 预览区
        body = QFrame()
        body.setStyleSheet(f"background:{c['BG']};")
        body_lay = QVBoxLayout(body)
        body_lay.setContentsMargins(20, 16, 20, 16)

        preview_label = QLabel("预览：")
        preview_label.setStyleSheet(
            f"color:{c['TEXT']};font-size:10pt;font-weight:600;"
        )
        body_lay.addWidget(preview_label)

        self._text_edit = QTextEdit()
        self._text_edit.setReadOnly(True)
        self._text_edit.setPlainText(self._text)
        self._text_edit.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._text_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._text_edit.setStyleSheet(
            f"background:{c['SURFACE']};color:{c['TEXT']};"
            f"border:1px solid {c['BORDER']};border-radius:6px;"
            f"padding:12px;font-size:10pt;"
            f"font-family:'Consolas','Courier New',monospace;"
            f"QScrollBar:vertical {{background:{c['BG']};width:8px;"
            f"border:none;border-radius:4px;margin:2px;}}"
            f"QScrollBar::handle:vertical {{background:{c['MUTED']};"
            f"border:none;border-radius:4px;min-height:30px;}}"
            f"QScrollBar::handle:vertical:hover {{background:{c['TEXT']};}}"
            f"QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical "
            f"{{height:0;border:none;}}"
            f"QScrollBar::add-page:vertical,QScrollBar::sub-page:vertical "
            f"{{background:none;}}"
        )
        body_lay.addWidget(self._text_edit, 1)
        layout.addWidget(body, 1)

        # 底部按钮栏
        footer = QFrame()
        footer.setStyleSheet(f"background:{c['BG']};")
        footer_lay = QHBoxLayout(footer)
        footer_lay.setContentsMargins(20, 12, 20, 16)
        footer_lay.addStretch()

        copy_btn = QPushButton("复制到剪贴板")
        copy_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        copy_btn.setStyleSheet(
            f"background:{c['PRIMARY']};color:#fff;font-size:11pt;"
            f"font-weight:bold;padding:10px 28px;border:none;border-radius:6px;"
        )
        copy_btn.clicked.connect(self._copy_to_clipboard)
        footer_lay.addWidget(copy_btn)

        ai_btn = QPushButton("复制到AI分析定制学习计划")
        ai_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        ai_btn.setStyleSheet(
            f"background:{c['PRIMARY']};color:#fff;font-size:11pt;"
            f"font-weight:bold;padding:10px 24px;border:none;border-radius:6px;"
        )
        ai_btn.clicked.connect(self._copy_to_ai_learn_plan)
        footer_lay.addWidget(ai_btn)

        close_btn = QPushButton("关闭")
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setStyleSheet(
            f"background:{c['SURFACE']};color:{c['TEXT']};font-size:11pt;"
            f"font-weight:bold;padding:10px 28px;"
            f"border:1px solid {c['BORDER']};border-radius:6px;"
        )
        close_btn.clicked.connect(self.accept)
        footer_lay.addWidget(close_btn)

        layout.addWidget(footer)

    def _copy_to_clipboard(self) -> None:
        """复制文本到剪贴板."""
        clipboard = QApplication.clipboard()
        if clipboard:
            clipboard.setText(self._text)
        self.accept()

    def _copy_to_ai_learn_plan(self) -> None:
        """复制带 AI 分析前缀的文本到剪贴板，询问是否跳转 AI 页面."""
        full_text = _AI_PROMPT_PREFIX + self._text
        clipboard = QApplication.clipboard()
        if clipboard:
            clipboard.setText(full_text)
        reply = themed_question(
            self,
            "提示",
            "已复制错题以及标记题，直接粘贴AI即可，是否跳转到AI页面？",
        )
        if reply == QMessageBox.StandardButton.Yes:
            QDesktopServices.openUrl(QUrl("https://chat.deepseek.com"))
