"""题目数据模型."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Question:
    """一道题目的数据模型."""

    index: int  # 全局序号（从0开始）
    number: str  # 题号（如 "1", "T1"）
    q_type: str  # 题型（单选/填空/判断/程序填空/程序改错/程序设计）
    text: str  # 题目内容
    options: dict[str, str] = field(default_factory=dict)  # {"A": "选项内容", ...}
    correct_answer: str = ""  # 正确答案
    score: int = 0  # 分值
    program_file: str = ""  # 程序文件名（如有）

    @classmethod
    def from_dict(cls, data: dict, global_index: int) -> Question:
        """从字典创建 Question（兼容旧 pandas 行格式）."""
        options = {}
        for letter in ["A", "B", "C", "D", "E", "F"]:
            opt_key_1 = f"选项{letter}"
            opt_key_2 = f"选项 {letter}"
            text = data.get(opt_key_1, data.get(opt_key_2, "")).strip()
            if text:
                options[letter] = text

        try:
            score = int(float(data.get("分值", 0)))
        except (ValueError, TypeError):
            score = 0

        return cls(
            index=global_index,
            number=str(data.get("题号", "")).strip(),
            q_type=str(data.get("题型", "")).strip(),
            text=str(data.get("题目", "")).strip(),
            options=options,
            correct_answer=str(data.get("正确答案", "")).strip(),
            score=score,
            program_file=str(data.get("程序文件", "")).strip(),
        )
