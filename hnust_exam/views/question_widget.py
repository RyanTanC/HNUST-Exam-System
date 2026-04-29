"""题目展示与答题控件."""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QScroller,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QLineEdit,
    QTextEdit,
    QRadioButton,
    QButtonGroup,
    QFrame,
    QScrollArea,
    QSizePolicy,
)

from hnust_exam.models.question import Question
from hnust_exam.utils.helpers import normalize_answer
from hnust_exam.utils.theme import Theme

if TYPE_CHECKING:
    from hnust_exam.views.exam_page import ExamPage


class QuestionWidget(QWidget):
    """题目展示与答题控件."""

    def __init__(self, exam_page: ExamPage, parent=None) -> None:
        super().__init__(parent)
        self.exam_page = exam_page
        self._choice_buttons: dict[str, QPushButton] = {}
        self._answer_text: QTextEdit | None = None
        self._answer_entry: QLineEdit | None = None
        self._answer_card: QFrame | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 滚动区域
        self._scroll_area = QScrollArea()
        self._scroll_area.setWidgetResizable(True)
        self._scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll_area.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        # 启用系统级惯性滚动
        QScroller.grabGesture(self._scroll_area.viewport(), QScroller.ScrollerGestureType.TouchGesture)

        self._content = QWidget()
        self._content.setStyleSheet(f"background-color: {c['WHITE']};")
        self._content_layout = QVBoxLayout(self._content)
        self._content_layout.setContentsMargins(20, 20, 20, 20)

        # 题目标题栏
        self._title_bar = QLabel("")
        self._title_bar.setStyleSheet(
            f"background-color: {c['PRIMARY']}; color: white; font-size: 10pt; "
            f"font-weight: bold; padding: 8px 10px; border-radius: 4px;"
        )
        self._title_bar.setWordWrap(True)
        self._content_layout.addWidget(self._title_bar)

        # 答题说明（程序题等）
        self._instruction_label = QLabel("")
        self._instruction_label.setWordWrap(True)
        self._instruction_label.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 9pt; padding: 5px 0 10px 0;"
        )
        self._instruction_label.hide()
        self._content_layout.addWidget(self._instruction_label)

        # 题目内容
        self._question_text = QLabel("")
        self._question_text.setWordWrap(True)
        self._question_text.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self._question_text.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 11pt; font-weight: 500; padding: 10px 0;"
        )
        self._content_layout.addWidget(self._question_text)

        # 选项区域
        self._options_frame = QFrame()
        self._options_layout = QVBoxLayout(self._options_frame)
        self._options_layout.setContentsMargins(0, 0, 0, 0)
        self._content_layout.addWidget(self._options_frame)

        # 答题提示
        self._hint_label = QLabel("")
        self._hint_label.setStyleSheet(
            f"background-color: {c['HINT_BG']}; color: {c['HINT_TEXT']}; "
            f"font-size: 9pt; padding: 5px 10px; border-radius: 4px;"
        )
        self._hint_label.hide()
        self._content_layout.addWidget(self._hint_label)

        # 答案输入区
        self._answer_area = QFrame()
        self._answer_layout = QHBoxLayout(self._answer_area)
        self._answer_layout.setContentsMargins(0, 10, 0, 0)
        self._content_layout.addWidget(self._answer_area)

        # 反馈区
        self._feedback_label = QLabel("")
        self._feedback_label.setWordWrap(True)
        self._feedback_label.setStyleSheet(
            f"font-size: 11pt; font-weight: bold; padding: 5px 0;"
        )
        self._content_layout.addWidget(self._feedback_label)

        self._content_layout.addStretch()

        self._scroll_area.setWidget(self._content)
        layout.addWidget(self._scroll_area)

        # 快捷键提示
        self._kb_hint = QLabel(
            "快捷键  ← → 切换题目  |  Ctrl+N 下一未答  |  "
            "Ctrl+A 查看答案  |  A B C D 选择选项"
        )
        self._kb_hint.setStyleSheet(
            f"background-color: {c['KB_HINT_BG']}; color: {c['MUTED']}; "
            f"font-size: 8pt; padding: 5px; min-height: 15px;"
        )
        layout.addWidget(self._kb_hint)

    def refresh_theme(self) -> None:
        """刷新主题颜色."""
        c = Theme.get_current_colors()
        self._content.setStyleSheet(f"background-color: {c['WHITE']};")
        self._title_bar.setStyleSheet(
            f"background-color: {c['PRIMARY']}; color: white; font-size: 10pt; "
            f"font-weight: bold; padding: 8px 10px; border-radius: 4px;"
        )
        self._instruction_label.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 9pt; padding: 5px 0 10px 0;"
        )
        self._question_text.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 11pt; font-weight: 500; padding: 10px 0;"
        )
        self._hint_label.setStyleSheet(
            f"background-color: {c['HINT_BG']}; color: {c['HINT_TEXT']}; "
            f"font-size: 9pt; padding: 5px 10px; border-radius: 4px;"
        )
        self._kb_hint.setStyleSheet(
            f"background-color: {c['KB_HINT_BG']}; color: {c['MUTED']}; "
            f"font-size: 8pt; padding: 5px; min-height: 15px;"
        )
        # 刷新动态内容（选项按钮、输入框等）
        if self.exam_page.exam:
            self.show_question()

    def show_question(self) -> None:
        """显示当前题目."""
        exam = self.exam_page.exam
        if not exam:
            return

        q = exam.get_question(exam.current_index)
        if not q:
            return

        c = Theme.get_current_colors()
        global_num = q.number

        # 计算类型内序号
        type_questions = exam.question_groups.get(q.q_type, [])
        type_idx = type_questions.index(q) + 1 if q in type_questions else 0

        # 标题
        total_score = sum(x.score for x in exam.questions)
        self._title_bar.setText(
            f"[{exam.current_index + 1}/{exam.total_count}] "
            f"{q.q_type} - 第{type_idx}题（题号：{global_num}）- {q.score}分"
            f"（共{exam.total_count}题，共{total_score}分）"
        )

        # 题目内容
        self._question_text.setText(q.text)

        # 清除旧控件
        self._clear_layout(self._options_layout)
        self._clear_layout(self._answer_layout)
        self._choice_buttons.clear()
        self._answer_text = None
        self._answer_entry = None
        if self._answer_card is not None:
            self._answer_card.deleteLater()
            self._answer_card = None
        self._feedback_label.setText("")
        self._feedback_label.setStyleSheet(f"font-size: 11pt; font-weight: bold; padding: 5px 0;")

        current_answer = exam.get_answer(global_num)

        if q.q_type == "单选":
            self._build_single_choice(q, global_num, current_answer, c)
        elif q.q_type == "判断":
            self._build_judge(q, global_num, current_answer, c)
        elif q.q_type in ("填空", "程序填空", "程序改错", "程序设计"):
            self._build_text_input(q, global_num, current_answer, q.q_type, c)

        # 答题说明
        if q.q_type == "程序设计":
            self._instruction_label.setText(
                "<<答题说明>>\n"
                '1. 点击下方"答题"按钮，系统会用IDLE打开对应的程序文件\n'
                "2. 在程序中完成需求（如补全代码、修复bug）\n"
                "3. 修改完成后按Ctrl+S保存文件\n"
                "4. 回到本系统，在答案框中输入核心代码\n"
                "注意：需按题目要求编写代码，保存后再提交答案。"
            )
            self._instruction_label.show()
        elif q.q_type in ("程序改错", "程序填空"):
            self._instruction_label.setText(
                "<<答题说明>>\n"
                '1. 点击下方"答题"按钮，系统会用IDLE打开对应的程序文件\n'
                "   （无需安装PyCharm，Python自带IDLE即可）\n"
                "2. 在**********FOUND**********语句的下一行修改程序\n"
                "3. 修改完成后按Ctrl+S保存文件\n"
                "4. 回到本系统，在答案框中输入修改后的内容\n"
                "注意：不可以增加或删除程序行，也不可以更改程序的结构。"
            )
            self._instruction_label.show()
        else:
            self._instruction_label.hide()

        # 答题按钮可见性
        program_types = {"程序设计", "程序填空", "程序改错"}
        self.exam_page._buttons["答题"].setVisible(q.q_type in program_types)

        # 标记按钮文本
        if exam.is_marked(exam.current_index):
            self.exam_page._buttons["标记试题"].setText("取消标记")
        else:
            self.exam_page._buttons["标记试题"].setText("标记试题")

        # 滚动到顶部
        self._scroll_area.verticalScrollBar().setValue(0)

    def _build_single_choice(self, q: Question, global_num: str, current_answer: str, c: dict) -> None:
        """构建单选题 UI."""
        # 显示选项文字
        for letter, text in q.options.items():
            opt_label = QLabel(f"({letter}){text}")
            opt_label.setWordWrap(True)
            opt_label.setFrameShape(QFrame.Shape.NoFrame)
            opt_label.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; padding: 5px 0;")
            self._options_layout.addWidget(opt_label)

        self._hint_label.setText("在下面选择答案（点击选项字母进行选择）")
        self._hint_label.show()

        for letter in q.options.keys():
            btn = QPushButton(letter)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            btn.setFixedSize(50, 40)
            if current_answer == letter:
                btn.setStyleSheet(
                    f"background-color: {c['PRIMARY']}; color: white; "
                    f"font-size: 16pt; font-weight: bold; border: none; border-radius: 4px;"
                )
            else:
                btn.setStyleSheet(
                    f"background-color: {c['BG']}; color: {c['TEXT']}; "
                    f"font-size: 16pt; font-weight: bold; border: none; border-radius: 4px;"
                )
            btn.clicked.connect(lambda checked, o=letter: self._on_choice(global_num, o))
            self._answer_layout.addWidget(btn)
            self._choice_buttons[letter] = btn

        self._answer_layout.addStretch()

    def _build_judge(self, q: Question, global_num: str, current_answer: str, c: dict) -> None:
        """构建判断题 UI."""
        self._hint_label.setText("在下面选择答案（点击对或错进行选择）")
        self._hint_label.show()

        for label, value in [("对", "A"), ("错", "B")]:
            btn = QPushButton(label)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            btn.setFixedSize(80, 40)
            if current_answer == value:
                btn.setStyleSheet(
                    f"background-color: {c['PRIMARY']}; color: white; "
                    f"font-size: 16pt; font-weight: bold; border: none; border-radius: 4px;"
                )
            else:
                btn.setStyleSheet(
                    f"background-color: {c['BG']}; color: {c['TEXT']}; "
                    f"font-size: 16pt; font-weight: bold; border: none; border-radius: 4px;"
                )
            btn.clicked.connect(lambda checked, v=value: self._on_choice(global_num, v))
            self._answer_layout.addWidget(btn)
            self._choice_buttons[value] = btn

        self._answer_layout.addStretch()

    def _build_text_input(
        self, q: Question, global_num: str, current_answer: str, q_type: str, c: dict
    ) -> None:
        """构建文本输入 UI."""
        self._hint_label.setText("在下面输入答案")
        self._hint_label.show()

        ans_label = QLabel("答案：")
        ans_label.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt;")
        self._answer_layout.addWidget(ans_label)

        if q_type in ("程序填空", "程序改错", "程序设计"):
            self._answer_text = QTextEdit()
            self._answer_text.setFixedHeight(120)
            self._answer_text.setPlainText(current_answer)
            self._answer_text.setStyleSheet(
                f"background-color: {c['INPUT_BG']}; color: {c['TEXT']}; "
                f"border: 1px solid {c['BORDER']}; border-radius: 4px; "
                f"font-family: Consolas, monospace; font-size: 11pt; padding: 4px;"
            )
            self._answer_layout.addWidget(self._answer_text)
        else:
            self._answer_entry = QLineEdit()
            self._answer_entry.setText(current_answer)
            self._answer_entry.setStyleSheet(
                f"background-color: {c['INPUT_BG']}; color: {c['TEXT']}; "
                f"border: 1px solid {c['BORDER']}; border-radius: 4px; "
                f"font-size: 11pt; padding: 4px 8px;"
            )
            self._answer_layout.addWidget(self._answer_entry)

    def _on_choice(self, global_num: str, option: str) -> None:
        """选择题/判断题选择事件."""
        exam = self.exam_page.exam
        if not exam:
            return

        exam.set_answer(global_num, option)
        q = exam.get_question(exam.current_index)
        if not q:
            return

        c = Theme.get_current_colors()

        if self.exam_page.show_answer_immediately:
            correct = normalize_answer(q.correct_answer, q.q_type)
            chosen = normalize_answer(option, q.q_type)

            for letter, btn in self._choice_buttons.items():
                normalized = normalize_answer(letter, q.q_type)
                if normalized == correct:
                    btn.setStyleSheet(
                        f"background-color: {c['SUCCESS']}; color: white; "
                        f"font-size: 16pt; font-weight: bold; border: none; border-radius: 4px;"
                    )
                elif letter == option and chosen != correct:
                    btn.setStyleSheet(
                        f"background-color: {c['DANGER']}; color: white; "
                        f"font-size: 16pt; font-weight: bold; border: none; border-radius: 4px;"
                    )
                else:
                    btn.setStyleSheet(
                        f"background-color: {c['BG']}; color: {c['TEXT']}; "
                        f"font-size: 16pt; font-weight: bold; border: none; border-radius: 4px;"
                    )

            if chosen == correct:
                self._feedback_label.setText("回答正确！")
                self._feedback_label.setStyleSheet(
                    f"color: {c['SUCCESS']}; font-size: 11pt; font-weight: bold; padding: 5px 0;"
                )
            else:
                self._feedback_label.setText(f"回答错误，正确答案是：{q.correct_answer}")
                self._feedback_label.setStyleSheet(
                    f"color: {c['DANGER']}; font-size: 11pt; font-weight: bold; padding: 5px 0;"
                )
        else:
            for letter, btn in self._choice_buttons.items():
                if letter == option:
                    btn.setStyleSheet(
                        f"background-color: {c['PRIMARY']}; color: white; "
                        f"font-size: 16pt; font-weight: bold; border: none; border-radius: 4px;"
                    )
                else:
                    btn.setStyleSheet(
                        f"background-color: {c['BG']}; color: {c['TEXT']}; "
                        f"font-size: 16pt; font-weight: bold; border: none; border-radius: 4px;"
                    )
            self._feedback_label.setText("")

        self.exam_page.nav_panel.refresh()
        self.exam_page._update_progress()

    def save_current_answer(self) -> None:
        """保存当前文本输入的答案."""
        exam = self.exam_page.exam
        if not exam:
            return

        q = exam.get_question(exam.current_index)
        if not q:
            return

        if self._answer_text is not None:
            answer = self._answer_text.toPlainText().strip()
            exam.set_answer(q.number, answer)
        elif self._answer_entry is not None:
            answer = self._answer_entry.text().strip()
            exam.set_answer(q.number, answer)

    def display_correct_answer(self) -> None:
        """显示正确答案."""
        exam = self.exam_page.exam
        if not exam:
            return
        q = exam.get_question(exam.current_index)
        if not q:
            return

        # 再次点击则隐藏
        if self._answer_card is not None:
            self._answer_card.deleteLater()
            self._answer_card = None
            self._feedback_label.setText("")
            self._feedback_label.setStyleSheet("font-size: 11pt; font-weight: bold; padding: 5px 0;")
            return

        c = Theme.get_current_colors()

        ans_frame = QFrame()
        ans_frame.setStyleSheet(
            f"background-color: {c['ANSWER_BG']}; border: 1px solid {c['SUCCESS']}; "
            f"border-radius: 4px; padding: 10px;"
        )
        ans_layout = QVBoxLayout(ans_frame)

        title = QLabel("标准答案")
        title.setStyleSheet(
            f"color: {c['SUCCESS']}; font-size: 10pt; font-weight: bold; "
            f"border: none;"
        )
        ans_layout.addWidget(title)

        ans_text = QLabel(q.correct_answer)
        ans_text.setWordWrap(True)
        ans_text.setStyleSheet(
            f"color: {c['SUCCESS']}; font-size: 17pt; font-weight: bold; "
            f"font-family: Consolas; border: none;"
        )
        ans_layout.addWidget(ans_text)

        self._content_layout.insertWidget(self._content_layout.count() - 1, ans_frame)
        self._answer_card = ans_frame

        self._feedback_label.setText("")
        self._feedback_label.setStyleSheet(
            f"color: {c['SUCCESS']}; font-size: 11pt; font-weight: bold; padding: 5px 0;"
        )
    @staticmethod
    def _clear_layout(layout) -> None:
        """清除布局中所有子控件."""
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
