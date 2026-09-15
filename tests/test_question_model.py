"""题目模型字段清洗测试."""

from __future__ import annotations

from hnust_exam.models.question import Question


def test_from_dict_converts_nan_core_fields_to_defaults():
    question = Question.from_dict({
        "题号": "nan",
        "题型": "nan",
        "题目": "nan",
        "正确答案": "nan",
        "分值": "2",
        "程序文件": "nan",
        "语言": "nan",
    }, 0)

    assert question.number == ""
    assert question.q_type == ""
    assert question.text == ""
    assert question.correct_answer == ""
    assert question.program_file == ""
    assert question.language == "python"
