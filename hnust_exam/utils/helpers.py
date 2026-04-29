"""无状态工具函数."""

import os
import sys
import re


def get_resource_path(relative_path: str) -> str:
    """获取打包后的资源路径."""
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def version_tuple(v: str) -> tuple:
    """将版本号字符串转为可比较的元组."""
    v = v.lstrip("vV")
    parts = v.split("-")
    main = tuple(int(x) for x in parts[0].split(".") if x.isdigit())
    return main


def normalize_answer(ans: str, q_type: str) -> str:
    """标准化答案字符串用于比较.

    判断题将中文/布尔/数字形式统一映射为 a/b。
    """
    ans = str(ans).strip().lower()
    if q_type == "判断":
        mapping = {
            "对": "a", "错": "b", "t": "a", "f": "b",
            "true": "a", "false": "b", "1": "a", "0": "b",
            "y": "a", "n": "b", "yes": "a", "no": "b",
            "正确": "a", "错误": "b", "√": "a", "×": "b",
        }
        ans = mapping.get(ans, ans)
    return ans


def normalize_code(code: str) -> str:
    """标准化代码：strip每行，去除空行."""
    lines = [line.strip() for line in str(code).strip().splitlines()]
    lines = [line for line in lines if line]
    return "\n".join(lines)


def normalize_code_flexible(code: str) -> str:
    """灵活标准化代码：统一换行符、去除空行、统一引号."""
    code = str(code).strip()
    code = code.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in code.split("\n")]
    lines = [line for line in lines if line]
    normalized = [line.replace("'", '"') for line in lines]
    return "\n".join(normalized)


def get_log_dir() -> str:
    """获取日志目录."""
    try:
        docs = os.path.expanduser("~/Documents")
        if not os.path.isdir(docs):
            docs = os.path.expanduser("~")
        log_dir = os.path.join(docs, "HNUST_Exam_Logs")
        os.makedirs(log_dir, exist_ok=True)
        return log_dir
    except Exception:
        return os.path.dirname(os.path.abspath(sys.argv[0]))
