"""考试模型：加载试卷、管理答案/标记、提供题目操作."""

from __future__ import annotations

import os
from typing import Optional

import pandas as pd

from hnust_exam.models.question import Question
from hnust_exam.models.result import Result
from hnust_exam.services.grader import grade_exam
from hnust_exam.utils.constants import REQUIRED_COLUMNS
from hnust_exam.utils.helpers import get_resource_path


class Exam:
    """考试模型，管理题目列表、用户答案、标记等状态."""

    def __init__(self) -> None:
        self.questions: list[Question] = []
        self.question_groups: dict[str, list[Question]] = {}
        self.active_type_order: list[str] = []
        self.answer_map: dict[str, str] = {}  # {题号: 用户答案}
        self.marked_indices: set[int] = set()  # 全局索引集合
        self.current_index: int = 0
        self.is_pure_program_exam: bool = False

    def load_from_excel(self, file_path: str) -> str:
        """从 Excel 文件加载试卷。成功返回空字符串，失败返回错误信息."""
        if not os.path.exists(file_path):
            return f"找不到文件：{file_path}"

        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            return f"读取Excel失败：{e}"

        df.columns = df.columns.str.strip()
        df = df.fillna("")
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()
        df = df[df["题号"] != ""]
        df = df[df["题目"] != ""]

        missing = REQUIRED_COLUMNS - set(df.columns)
        if missing:
            return f"Excel 缺少必要列：{', '.join(missing)}\n当前列：{', '.join(df.columns)}"

        if "程序文件" in df.columns:
            df["程序文件"] = df["程序文件"].astype(str).str.strip()
        else:
            df["程序文件"] = ""

        # 重复题号警告
        dup_nums = df[df.duplicated(subset=["题号"], keep=False)]["题号"].unique()
        dup_warning = ""
        if len(dup_nums) > 0:
            preview = ", ".join(dup_nums[:5].tolist())
            dup_warning = (
                f"发现重复题号：{preview}"
                f"{'...' if len(dup_nums) > 5 else ''}\n"
                "可能导致答案被覆盖，建议检查Excel文件。"
            )

        # 创建 Question 列表
        self.questions = []
        for idx, (_, row) in enumerate(df.iterrows()):
            q = Question.from_dict(row.to_dict(), idx)
            self.questions.append(q)

        # 按题型分组
        all_types = {q.q_type for q in self.questions}
        if all_types == {"程序设计"}:
            self.is_pure_program_exam = True
            self.question_groups = {"程序设计": list(self.questions)}
        else:
            self.is_pure_program_exam = False
            self.question_groups = {}
            for q in self.questions:
                self.question_groups.setdefault(q.q_type, []).append(q)

        # 题型显示顺序（按试卷中首次出现的顺序）
        self.active_type_order = []
        for q in self.questions:
            if q.q_type not in self.active_type_order:
                self.active_type_order.append(q.q_type)

        # 重置状态
        self.answer_map = {}
        self.marked_indices = set()
        self.current_index = 0

        return dup_warning

    def get_question(self, index: int) -> Optional[Question]:
        """按全局索引获取题目."""
        if 0 <= index < len(self.questions):
            return self.questions[index]
        return None

    def set_answer(self, question_number: str, answer: str) -> None:
        """设置用户答案."""
        if answer:
            self.answer_map[question_number] = answer
        else:
            self.answer_map.pop(question_number, None)

    def get_answer(self, question_number: str) -> str:
        """获取用户答案."""
        return self.answer_map.get(question_number, "")

    def toggle_mark(self, index: int) -> bool:
        """切换标记状态，返回是否已标记."""
        if index in self.marked_indices:
            self.marked_indices.remove(index)
            return False
        else:
            self.marked_indices.add(index)
            return True

    def is_marked(self, index: int) -> bool:
        """判断题目是否已标记."""
        return index in self.marked_indices

    def next_unanswered_index(self) -> Optional[int]:
        """查找下一未答题的索引，支持循环."""
        total = len(self.questions)
        for i in range(self.current_index + 1, total):
            if self.questions[i].number not in self.answer_map:
                return i
        for i in range(self.current_index + 1):
            if self.questions[i].number not in self.answer_map:
                return i
        return None

    @property
    def answered_count(self) -> int:
        return len(self.answer_map)

    @property
    def total_count(self) -> int:
        return len(self.questions)

    @property
    def marked_count(self) -> int:
        return len(self.marked_indices)

    @property
    def unanswered_count(self) -> int:
        return self.total_count - self.answered_count

    def grade(self) -> list[Result]:
        """判分，返回每题结果."""
        return grade_exam(self)

    @property
    def score_percentage(self) -> float:
        """计算得分百分比."""
        results = self.grade()
        total = sum(r.score for r in results)
        if total == 0:
            return 0.0
        earned = sum(r.score for r in results if r.is_correct)
        return earned / total * 100

    def find_exam_file(self, exam_name: str) -> Optional[str]:
        """查找试卷文件路径."""
        internal_path = get_resource_path(
            os.path.join("题库", exam_name + ".xlsx")
        )
        external_path = os.path.join("题库", exam_name + ".xlsx")

        if os.path.exists(internal_path):
            return internal_path
        elif os.path.exists(external_path):
            return external_path
        return None

    @staticmethod
    def list_exam_files() -> list[str]:
        """列出所有可用的试卷文件名（含扩展名）."""
        exam_dir = get_resource_path("题库")
        if not os.path.exists(exam_dir):
            external = "题库"
            if not os.path.exists(external):
                os.makedirs(external)
            exam_dir = external

        return [f for f in os.listdir(exam_dir) if f.endswith(".xlsx")]

    def get_options_for_question(self, q: Question) -> list[tuple[str, str]]:
        """获取题目的选项列表."""
        return list(q.options.items())
