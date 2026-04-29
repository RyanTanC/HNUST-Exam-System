"""判分服务：对考试答案进行评判."""

from __future__ import annotations

import re
import difflib
from typing import TYPE_CHECKING

from hnust_exam.models.result import Result
from hnust_exam.utils.helpers import (
    normalize_answer,
    normalize_code,
    normalize_code_flexible,
)

if TYPE_CHECKING:
    from hnust_exam.models.exam import Exam


def check_fill_in(user_ans: str, correct_ans: str) -> bool:
    """检查填空题答案（支持多空，用分号/逗号分隔）."""
    user_parts = re.split(r"[,;，；]+", user_ans.strip())
    correct_parts = re.split(r"[,;，；]+", correct_ans.strip())
    user_parts = [p for p in user_parts if p]
    correct_parts = [p for p in correct_parts if p]
    if len(user_parts) != len(correct_parts):
        return False
    return all(
        u.strip().lower() == c.strip().lower()
        for u, c in zip(user_parts, correct_parts)
    )


def check_program_answer(user_ans: str, correct_ans: str) -> bool:
    """检查程序题答案：精确比对 + 灵活比对."""
    if normalize_code(user_ans) == normalize_code(correct_ans):
        return True
    if normalize_code_flexible(user_ans) == normalize_code_flexible(correct_ans):
        return True
    # 相似度阈值判定（用于程序改错等）
    seq = difflib.SequenceMatcher(
        None, normalize_code_flexible(user_ans), normalize_code_flexible(correct_ans)
    )
    return seq.ratio() >= 0.9


def grade_exam(exam: Exam) -> list[Result]:
    """对整份试卷判分，返回每题结果."""
    results = []
    for q in exam.questions:
        q_type = q.q_type
        score = q.score
        user_ans_raw = exam.get_answer(q.number)
        is_correct = False

        if user_ans_raw:
            user_ans = normalize_answer(user_ans_raw, q_type)
            correct_ans = normalize_answer(q.correct_answer, q_type)

            if q_type in ("填空", "程序填空"):
                is_correct = check_fill_in(user_ans, correct_ans)
            elif q_type in ("程序设计", "程序改错"):
                is_correct = check_program_answer(user_ans_raw, q.correct_answer)
            else:
                is_correct = user_ans == correct_ans

        results.append(
            Result(
                question_number=q.number,
                q_type=q_type,
                score=score,
                user_answer=user_ans_raw if user_ans_raw else "未作答",
                correct_answer=q.correct_answer,
                is_correct=is_correct,
                question_text=q.text,
            )
        )
    return results
