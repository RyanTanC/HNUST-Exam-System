"""判分服务测试."""

import pytest

from hnust_exam.services.grader import (
    check_fill_in,
    check_program_answer,
    grade_exam,
)
from hnust_exam.models.exam import Exam
from hnust_exam.models.question import Question


class TestCheckFillIn:
    """填空题判分测试."""

    def test_exact_match(self):
        assert check_fill_in("abc", "abc") is True

    def test_case_insensitive(self):
        assert check_fill_in("Abc", "abc") is True

    def test_multi_blank_semicolon(self):
        assert check_fill_in("a;b;c", "a;b;c") is True

    def test_multi_blank_comma(self):
        assert check_fill_in("a,b,c", "a,b,c") is True

    def test_multi_blank_mismatch_count(self):
        assert check_fill_in("a;b", "a;b;c") is False

    def test_whitespace_trimmed(self):
        assert check_fill_in("  a ; b  ", "a;b") is True

    def test_wrong_answer(self):
        assert check_fill_in("xyz", "abc") is False

    def test_answer_with_space_not_split(self):
        # 答案包含空格时不应被拆分（如 "for loop" 作为单个答案）
        assert check_fill_in("for loop", "for loop") is True
        assert check_fill_in("for loop", "for") is False


class TestCheckProgramAnswer:
    """程序题判分测试."""

    def test_exact_match(self):
        assert check_program_answer("print(1)", "print(1)") is True

    def test_whitespace_normalized(self):
        assert check_program_answer("  print(1)  \n", "print(1)") is True

    def test_quote_normalization(self):
        assert check_program_answer("print('hello')", 'print("hello")') is True

    def test_multiline_exact(self):
        code = "for i in range(5):\n    print(i)"
        assert check_program_answer(code, code) is True

    def test_similar_code(self):
        # 差异极小的代码应通过（相似度>=0.9）
        user = "print('hello world')"
        correct = "print( 'hello world' )"
        assert check_program_answer(user, correct) is True

    def test_completely_wrong(self):
        assert check_program_answer("import os", "print('hello')") is False


class TestGradeExam:
    """整卷判分测试."""

    def _make_exam(self, questions_data: list[dict]) -> Exam:
        exam = Exam()
        exam.questions = []
        for idx, qd in enumerate(questions_data):
            q = Question(
                index=idx,
                number=qd["number"],
                q_type=qd["type"],
                text=qd.get("text", ""),
                correct_answer=qd["correct"],
                score=qd.get("score", 2),
            )
            exam.questions.append(q)
            exam.answer_map[qd["number"]] = qd.get("user", "")
        return exam

    def test_single_choice_all_correct(self):
        exam = self._make_exam([
            {"number": "1", "type": "单选", "correct": "A", "user": "A", "score": 2},
            {"number": "2", "type": "单选", "correct": "B", "user": "B", "score": 3},
        ])
        results = grade_exam(exam)
        assert len(results) == 2
        assert results[0].is_correct is True
        assert results[1].is_correct is True

    def test_judge_type(self):
        exam = self._make_exam([
            {"number": "1", "type": "判断", "correct": "A", "user": "对", "score": 2},
            {"number": "2", "type": "判断", "correct": "B", "user": "错", "score": 2},
        ])
        results = grade_exam(exam)
        assert results[0].is_correct is True
        assert results[1].is_correct is True

    def test_fill_in_multi_blank(self):
        exam = self._make_exam([
            {"number": "1", "type": "填空", "correct": "a;b", "user": "a;b", "score": 4},
        ])
        results = grade_exam(exam)
        assert results[0].is_correct is True

    def test_unanswered(self):
        exam = self._make_exam([
            {"number": "1", "type": "单选", "correct": "A", "user": "", "score": 2},
        ])
        results = grade_exam(exam)
        assert results[0].is_correct is False
        assert results[0].user_answer == "未作答"
