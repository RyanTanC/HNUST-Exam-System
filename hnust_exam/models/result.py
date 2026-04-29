"""判分结果数据模型."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Result:
    """单题判分结果."""

    question_number: str
    q_type: str
    score: int
    user_answer: str
    correct_answer: str
    is_correct: bool
    question_text: str
