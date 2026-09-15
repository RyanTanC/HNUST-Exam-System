"""试卷 Excel 加载的健壮性回归测试.

背景（真实 bug）：``Exam.load_from_excel()`` 原来把「缺列检查」放在了
按列名过滤**之后**：

    df = df[df["题号"] != ""]      # ← 缺「题号」列时这里直接 KeyError
    df = df[df["题目"] != ""]
    missing = REQUIRED_COLUMNS - set(df.columns)   # ← 永远走不到

于是用户拿到一份列名不对的 Excel 时，看到的是 ``KeyError: '题号'``
这种栈追踪，而不是"缺少必要列"的提示。检查必须提前到过滤之前。
"""

from __future__ import annotations

import pandas as pd
import pytest

from hnust_exam.models.exam import Exam

REQUIRED = ["题号", "题型", "题目", "正确答案", "分值"]


def _write_excel(tmp_path, columns: dict, name: str = "exam.xlsx") -> str:
    """把 {列名: [值]} 写成 xlsx，返回路径."""
    path = tmp_path / name
    pd.DataFrame(columns).to_excel(path, index=False)
    return str(path)


def _valid_columns() -> dict:
    return {
        "题号": [1],
        "题型": ["单选"],
        "题目": ["1+1=?"],
        "正确答案": ["B"],
        "分值": [2],
    }


class TestMissingColumns:
    """缺列时必须返回可读提示，不能抛异常."""

    @pytest.mark.parametrize("drop", REQUIRED)
    def test_each_missing_column_returns_message(self, tmp_path, drop):
        """缺任何一列都应返回友好提示，而不是 KeyError."""
        cols = _valid_columns()
        cols.pop(drop)
        path = _write_excel(tmp_path, cols)

        exam = Exam()
        result = exam.load_from_excel(path)  # 不应抛异常

        assert isinstance(result, str)
        assert "缺少必要列" in result, f"缺 {drop} 时应提示缺列，实际：{result!r}"
        assert drop in result, f"提示里应点名缺的是 {drop}"

    def test_missing_title_column_does_not_crash(self, tmp_path):
        """缺「题号」是最典型的场景（旧代码在这里崩）."""
        cols = _valid_columns()
        cols.pop("题号")
        path = _write_excel(tmp_path, cols)

        exam = Exam()
        result = exam.load_from_excel(path)

        assert "缺少必要列" in result
        assert "题号" in result
        assert exam.questions == [], "加载失败时不应留下半成品题目"

    def test_message_lists_actual_columns(self, tmp_path):
        """提示里要带上实际读到的列，方便用户对照排查."""
        cols = _valid_columns()
        cols.pop("分值")
        path = _write_excel(tmp_path, cols)

        result = Exam().load_from_excel(path)

        assert "当前列" in result
        assert "题号" in result


class TestValidAndEdgeCases:
    """正常与边界输入."""

    def test_valid_file_loads(self, tmp_path):
        path = _write_excel(tmp_path, _valid_columns())
        exam = Exam()

        result = exam.load_from_excel(path)

        assert result == "", f"正常文件应返回空串，实际 {result!r}"
        assert len(exam.questions) == 1
        assert exam.questions[0].number == "1"

    def test_column_names_are_stripped(self, tmp_path):
        """列名带空格要能容错（Excel 里很常见）."""
        cols = {f"  {k}  ": v for k, v in _valid_columns().items()}
        path = _write_excel(tmp_path, cols)

        result = Exam().load_from_excel(path)

        assert result == "", f"列名空格应被 strip，实际 {result!r}"

    def test_missing_file_returns_message(self):
        result = Exam().load_from_excel("C:/definitely/not/here.xlsx")

        assert "找不到文件" in result

    def test_blank_rows_are_dropped(self, tmp_path):
        """空题号 / 空题目的行应被忽略，而不是变成空题目."""
        cols = _valid_columns()
        cols["题号"] = [1, "", 2]
        cols["题型"] = ["单选", "单选", "单选"]
        cols["题目"] = ["1+1=?", "", "2+2=?"]
        cols["正确答案"] = ["B", "A", "C"]
        cols["分值"] = [2, 2, 2]
        path = _write_excel(tmp_path, cols)

        exam = Exam()
        result = exam.load_from_excel(path)

        assert result == ""
        assert len(exam.questions) == 2, f"空行应被丢弃，实际 {len(exam.questions)} 题"
