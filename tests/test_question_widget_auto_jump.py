"""题目控件自动跳转到下一未答行为测试."""

from __future__ import annotations

from PySide6.QtWidgets import QPushButton, QWidget

from hnust_exam.models.exam import Exam
from hnust_exam.models.question import Question
from hnust_exam.views.question_widget import QuestionWidget


class _NavPanel:
    def __init__(self) -> None:
        self.refresh_count = 0

    def refresh(self) -> None:
        self.refresh_count += 1


class _ExamPage(QWidget):
    def __init__(
        self,
        exam: Exam,
        auto_advance_on_choice: bool = True,
        show_answer_immediately: bool = False,
    ) -> None:
        super().__init__()
        self.exam = exam
        self.show_answer_immediately = show_answer_immediately
        self.auto_advance_on_choice = auto_advance_on_choice
        self.nav_panel = _NavPanel()
        self.progress_updates = 0
        self.jump_count = 0
        self._buttons = {"答题": QWidget(), "标记试题": QPushButton()}

    def _update_progress(self) -> None:
        self.progress_updates += 1

    def jump_next_unanswered(self) -> None:
        self.jump_count += 1


def _exam_with(*questions: Question) -> Exam:
    exam = Exam()
    exam.questions = list(questions)
    exam.question_groups = {}
    for q in questions:
        exam.question_groups.setdefault(q.q_type, []).append(q)
    exam.active_type_order = list(exam.question_groups)
    return exam


def _question(index: int, number: str, q_type: str, correct_answer: str = "A") -> Question:
    return Question(
        index=index,
        number=number,
        q_type=q_type,
        text=f"题目{number}",
        options={"A": "选项A", "B": "选项B"} if q_type in ("单选", "判断") else {},
        correct_answer=correct_answer,
        score=1,
    )


def test_choice_triggers_next_unanswered_after_normal_delay(qtbot):
    exam = _exam_with(_question(0, "1", "单选"), _question(1, "2", "单选"))
    page = _ExamPage(exam)
    qtbot.addWidget(page)
    widget = QuestionWidget(page)
    qtbot.addWidget(widget)
    widget.show_question()

    widget._on_choice("1", "A")

    assert exam.get_answer("1") == "A"
    assert page.jump_count == 0
    assert page.nav_panel.refresh_count == 1
    assert page.progress_updates == 1

    qtbot.wait(250)
    assert page.jump_count == 1


def test_choice_triggers_next_unanswered_after_feedback_delay(qtbot):
    exam = _exam_with(_question(0, "1", "单选"), _question(1, "2", "单选"))
    page = _ExamPage(exam, show_answer_immediately=True)
    qtbot.addWidget(page)
    widget = QuestionWidget(page)
    qtbot.addWidget(widget)
    widget.show_question()

    widget._on_choice("1", "A")

    assert page.jump_count == 0
    qtbot.wait(250)
    assert page.jump_count == 0
    qtbot.wait(500)
    assert page.jump_count == 1


def test_choice_respects_auto_advance_setting(qtbot):
    exam = _exam_with(_question(0, "1", "单选"), _question(1, "2", "单选"))
    page = _ExamPage(exam, auto_advance_on_choice=False)
    qtbot.addWidget(page)
    widget = QuestionWidget(page)
    qtbot.addWidget(widget)
    widget.show_question()

    widget._on_choice("1", "A")
    qtbot.wait(250)

    assert exam.get_answer("1") == "A"
    assert page.jump_count == 0


def test_fill_in_correct_answer_triggers_next_unanswered_after_normal_delay(qtbot):
    exam = _exam_with(_question(0, "1", "填空", "Python"), _question(1, "2", "单选"))
    page = _ExamPage(exam)
    qtbot.addWidget(page)
    widget = QuestionWidget(page)
    qtbot.addWidget(widget)
    widget.show_question()

    widget._answer_entry.setText(" python ")
    widget._answer_entry.editingFinished.emit()

    assert exam.get_answer("1") == "python"
    assert page.jump_count == 0
    assert page.nav_panel.refresh_count == 1
    assert page.progress_updates == 1

    qtbot.wait(250)
    assert page.jump_count == 1


def test_fill_in_correct_answer_triggers_next_unanswered_after_feedback_delay(qtbot):
    exam = _exam_with(_question(0, "1", "填空", "Python"), _question(1, "2", "单选"))
    page = _ExamPage(exam, show_answer_immediately=True)
    qtbot.addWidget(page)
    widget = QuestionWidget(page)
    qtbot.addWidget(widget)
    widget.show_question()

    widget._answer_entry.setText("python")
    widget._answer_entry.editingFinished.emit()

    assert page.jump_count == 0
    qtbot.wait(250)
    assert page.jump_count == 0
    qtbot.wait(500)
    assert page.jump_count == 1


def test_fill_in_correct_answer_respects_auto_advance_setting(qtbot):
    exam = _exam_with(_question(0, "1", "填空", "Python"), _question(1, "2", "单选"))
    page = _ExamPage(exam, auto_advance_on_choice=False)
    qtbot.addWidget(page)
    widget = QuestionWidget(page)
    qtbot.addWidget(widget)
    widget.show_question()

    widget._answer_entry.setText("python")
    widget._answer_entry.editingFinished.emit()
    qtbot.wait(250)

    assert exam.get_answer("1") == "python"
    assert page.jump_count == 0


def test_fill_in_wrong_answer_does_not_trigger_next_unanswered(qtbot):
    exam = _exam_with(_question(0, "1", "填空", "Python"), _question(1, "2", "单选"))
    page = _ExamPage(exam)
    qtbot.addWidget(page)
    widget = QuestionWidget(page)
    qtbot.addWidget(widget)
    widget.show_question()

    widget._answer_entry.setText("java")
    widget._answer_entry.editingFinished.emit()
    qtbot.wait(250)

    assert exam.get_answer("1") == "java"
    assert page.jump_count == 0
    assert page.nav_panel.refresh_count == 1
    assert page.progress_updates == 1


def test_program_text_question_does_not_auto_jump(qtbot):
    exam = _exam_with(_question(0, "1", "程序填空", "print('ok')"), _question(1, "2", "单选"))
    page = _ExamPage(exam)
    qtbot.addWidget(page)
    widget = QuestionWidget(page)
    qtbot.addWidget(widget)
    widget.show_question()

    widget._answer_text.setPlainText("print('ok')")
    widget.save_current_answer()
    qtbot.wait(250)

    assert exam.get_answer("1") == "print('ok')"
    assert page.jump_count == 0
