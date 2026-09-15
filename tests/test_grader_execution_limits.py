"""程序执行限制测试."""

from __future__ import annotations

from hnust_exam.services.grader import _execute_code


def test_execute_code_times_out_and_returns_none():
    assert _execute_code("while True:\n    pass", "python") is None


def test_execute_code_returns_stdout_for_normal_program():
    assert _execute_code("print('ok')", "python").strip() == "ok"
