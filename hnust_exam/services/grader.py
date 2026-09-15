"""判分服务：对考试答案进行评判.

复刻万维考试系统（VFGRADE.EXE）核心判分逻辑，支持：
- 单选/判断：归一化精确匹配
- 多选：全对满分，部分对半分，超集0分
- 填空：两阶段匹配（一次匹配 + 断词匹配）
- 程序设计：分阶段评分（编译门槛 → 输出+关键字+代码加权）
"""

from __future__ import annotations

import os
import re
import sys
import tokenize
import io
import difflib
import subprocess
import tempfile
import shutil
from typing import TYPE_CHECKING

from hnust_exam.models.result import Result
from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.helpers import (
    normalize_answer,
    normalize_code,
    normalize_code_flexible,
)

if TYPE_CHECKING:
    from hnust_exam.models.exam import Exam


def normalize_punctuation(s: str) -> str:
    """将全角标点转为半角，合并连续空白为一个空格，统一小写。

    用于填空题等文本答案的归一化比较。
    """
    # 全角标点 → 半角（用 dict 逐字符替换，避免 maketrans 引号拼接问题）
    _punct_map = {
        "，": ",",   # ，
        "。": ".",   # 。
        "！": "!",   # ！
        "？": "?",   # ？
        "；": ";",   # ；
        "：": ":",   # ：
        "“": '"',   # "
        "”": '"',   # "
        "‘": "'",   # '
        "’": "'",   # '
        "（": "(",   # （
        "）": ")",   # ）
        "【": "[",   # 【
        "】": "]",   # 】
        "《": "<",   # 《
        "》": ">",   # 》
    }
    s = "".join(_punct_map.get(ch, ch) for ch in s)
    # 合并连续空白为一个空格
    s = re.sub(r"\s+", " ", s)
    # 统一小写
    s = s.lower().strip()
    return s


def strip_comments(code: str, lang: str) -> str:
    """移除代码中的注释。

    Python: 移除 ``#`` 行注释。
    C: 移除 ``//`` 行注释和 ``/* */`` 块注释。
    """
    if lang.lower() == "python":
        try:
            lines = code.splitlines(keepends=True)
            readline = io.StringIO(code).readline
            for tok in tokenize.generate_tokens(readline):
                if tok.type == tokenize.COMMENT:
                    line_idx = tok.start[0] - 1
                    lines[line_idx] = lines[line_idx][:tok.start[1]]
            return "".join(lines)
        except tokenize.TokenError:
            return code
    elif lang.lower() == "c":
        # 先移除块注释 /* ... */
        code = re.sub(r"/\*.*?\*/", "", code, flags=re.DOTALL)
        # 再移除行注释 //
        code = re.sub(r"//.*$", "", code, flags=re.MULTILINE)
        return code
    return code


def get_program_tokens(code: str, lang: str) -> list[str]:
    """提取程序的核心 token 序列，忽略注释、空白、缩进。

    Python: 使用 ``tokenize`` 模块。
    C: 使用增强归一化（去注释 + normalize_code）。
    """
    lang = lang.lower()
    if lang == "python":
        code = strip_comments(code, "python")
        tokens: list[str] = []
        try:
            readline = io.StringIO(code).readline
            for tok in tokenize.generate_tokens(readline):
                if tok.type not in (
                    tokenize.COMMENT,
                    tokenize.NL,
                    tokenize.NEWLINE,
                    tokenize.INDENT,
                    tokenize.DEDENT,
                    tokenize.ENCODING,
                ):
                    tokens.append(tok.string)
        except tokenize.TokenError:
            # tokenize 失败时回退到归一化
            return normalize_code(code).split()
        return tokens
    elif lang == "c":
        cleaned = strip_comments(code, "c")
        normalized = normalize_code(cleaned)
        return normalized.split()
    # 未知语言，回退
    return normalize_code(code).split()


def _sanitize_answer(ans: str, q_type: str = "") -> str:
    """预清洗标准答案，修正题库录入时的人为错误。

    只应用于 correct_ans，不要用于 user_ans。
    """
    s = str(ans)

    # 通用：中文标点 → 英文
    s = s.replace("“", '"').replace("”", '"')   # " "
    s = s.replace("‘", "'").replace("’", "'")   # ' '
    s = s.replace("，", ",")                           # ，
    s = s.replace("；", ";")                           # ；
    s = s.replace("（", "(").replace("）", ")")   # （）

    # 引号配对修复（双引号 + 单引号）
    s = _fix_quote_pairs(s, '"')
    s = _fix_quote_pairs(s, "'")

    # 去掉首尾多余分隔符（逗号总是去除，分号仅去除连续 2+ 个的）
    s = re.sub(r"^[,]+|[,]+$", "", s.strip())
    s = re.sub(r"^[;]{2,}|[;]{2,}$", "", s.strip())

    return s


def _fix_quote_pairs(s: str, q: str) -> str:
    """修复字符串中指定引号的配对问题。"""
    count = s.count(q)
    if count % 2 == 0:
        return s
    # 奇数个引号：先删尾部，再删头部
    if s.endswith(q):
        s = s[:-1]
    if s.count(q) % 2 != 0 and s.startswith(q):
        s = s[1:]
    if s.count(q) % 2 != 0:
        # 仍未配对，删除最后一个
        idx = s.rfind(q)
        if idx >= 0:
            s = s[:idx] + s[idx + 1:]
    return s


def check_fill_in(user_ans: str, correct_ans: str, is_code: bool = False) -> tuple[bool, int, int]:
    """检查填空题答案，复刻万维两阶段匹配逻辑.

    阶段1（一次匹配）：normalize_punctuation → 分割 → 逐空精确比较
    阶段2（断词匹配）：一次匹配失败时触发
        - is_code=True：编辑距离比较（适用于代码填空）
        - is_code=False：断词 + Jaccard 相似度（适用于自然语言填空）

    Returns
    -------
    (all_correct, matched_count, total_blanks)
        all_correct: 是否全部正确
        matched_count: 匹配的空数
        total_blanks: 总空数
    """
    user_norm = normalize_punctuation(user_ans)
    correct_norm = normalize_punctuation(correct_ans)
    user_parts = re.split(r"[,;]+", user_norm)
    correct_parts = re.split(r"[,;]+", correct_norm)
    user_parts = [p.strip() for p in user_parts if p.strip()]
    correct_parts = [p.strip() for p in correct_parts if p.strip()]
    # 最多支持10个空
    if len(user_parts) > 10 or len(correct_parts) > 10:
        return False, 0, max(len(user_parts), len(correct_parts))
    if len(user_parts) != len(correct_parts):
        return False, 0, max(len(user_parts), len(correct_parts))

    # 阶段1：一次匹配（精确比较）
    matched = sum(1 for u, c in zip(user_parts, correct_parts) if u == c)
    total = len(correct_parts)
    if matched == total:
        return True, matched, total

    # 阶段2：断词匹配（一次匹配失败后触发）
    if is_code:
        # 代码填空：编辑距离比较
        threshold = _get_config_threshold("split_word_threshold", 0.7)
        matched = 0
        for u, c in zip(user_parts, correct_parts):
            if u == c:
                matched += 1
            else:
                ratio = difflib.SequenceMatcher(None, u, c).ratio()
                if ratio >= threshold:
                    matched += 1
    else:
        # 自然语言填空：断词 + Jaccard 相似度
        threshold = _get_config_threshold("split_word_threshold", 0.6)
        matched = 0
        for u, c in zip(user_parts, correct_parts):
            if u == c:
                matched += 1
            else:
                u_tokens = set(re.split(r"\s+", u))
                c_tokens = set(re.split(r"\s+", c))
                if c_tokens and len(u_tokens & c_tokens) / len(u_tokens | c_tokens) >= threshold:
                    matched += 1

    all_correct = matched == total
    return all_correct, matched, total


def check_program_answer(user_ans: str, correct_ans: str, lang: str = "python") -> bool:
    """检查程序题答案：token 比对 + 归一化 + 灵活比对.

    优先用 ``get_program_tokens`` 提取 token 序列比对；
    失败时用归一化 + 剥注释后 ``SequenceMatcher``，阈值从 config_manager 读取（默认 0.85）。
    """
    correct_ans = _sanitize_answer(correct_ans)

    # 优先：token 级比对
    try:
        user_tokens = get_program_tokens(user_ans, lang)
        correct_tokens = get_program_tokens(correct_ans, lang)
        if user_tokens == correct_tokens:
            return True
    except Exception:
        pass

    # 后备1：精确比对
    if normalize_code(user_ans) == normalize_code(correct_ans):
        return True
    if normalize_code_flexible(user_ans) == normalize_code_flexible(correct_ans):
        return True

    # 后备2：归一化 + 剥注释后 SequenceMatcher
    user_clean = strip_comments(normalize_code_flexible(user_ans), lang)
    correct_clean = strip_comments(normalize_code_flexible(correct_ans), lang)
    seq = difflib.SequenceMatcher(None, user_clean, correct_clean)

    # 从配置读取阈值，默认 0.85
    try:
        cfg = ConfigManager().load_config()
        threshold = cfg.get("grading_threshold", 0.85)
    except Exception:
        threshold = 0.85

    return seq.ratio() >= threshold


def _get_config_threshold(key: str, default: float) -> float:
    """从配置读取阈值，读取失败返回默认值."""
    try:
        cfg = ConfigManager().load_config()
        return cfg.get(key, default)
    except Exception:
        return default


def check_multi_select(user_ans: str, correct_ans: str, partial_credit: bool = True) -> tuple[bool, bool]:
    """多选题判分，复刻万维 MultiSelectTypeAll/MultiSelectTypeHalf 逻辑.

    Returns
    -------
    (is_correct, is_half)
        is_correct: 全对
        is_half: 半对（部分正确，可得一半分数）
    """
    # 支持 | 和 # 两种分隔符（兼容万维题库格式）
    def _parse_options(s: str) -> set[str]:
        s = str(s).strip().upper()
        s = s.replace("#", "|")
        parts = re.split(r"[|,;]+", s)
        return {p.strip() for p in parts if p.strip()}

    user_set = _parse_options(user_ans)
    correct_set = _parse_options(correct_ans)

    if not correct_set:
        return False, False

    # 全对
    if user_set == correct_set:
        return True, False

    # 超集（选了额外错误选项）→ 0分
    if user_set > correct_set:
        return False, False

    # 子集或有交集 → 半对（如果允许部分得分）
    if partial_credit and user_set & correct_set:
        return False, True

    return False, False


def check_compilation(code: str, lang: str) -> tuple[bool, str]:
    """检查程序是否能通过编译.

    复刻万维 NeedCheckProgramCompilerPass / ProgramCompilerPass 逻辑。
    编译失败 → 该题直接 0 分（"程序编译没有通过，判分结束"）。

    Returns
    -------
    (passed, message)
    """
    lang = lang.lower().strip()
    if lang == "python":
        try:
            compile(code, "<exam>", "exec")
            return True, "编译通过"
        except SyntaxError as e:
            return False, f"语法错误: {e}"
    elif lang == "c":
        gcc = shutil.which("gcc")
        if not gcc:
            return True, "编译器不可用，跳过检查"
        try:
            proc = subprocess.run(
                [gcc, "-fsyntax-only", "-x", "c", "-"],
                input=code.encode("utf-8"),
                capture_output=True,
                timeout=5,
            )
            if proc.returncode == 0:
                return True, "编译通过"
            stderr = proc.stderr.decode("utf-8", errors="replace")
            return False, f"编译错误: {stderr[:200]}"
        except subprocess.TimeoutExpired:
            return False, "编译超时"
        except Exception as e:
            return True, f"编译检查异常，跳过: {e}"
    # 未知语言，跳过
    return True, "未知语言，跳过编译检查"


def check_output(user_code: str, expected_output: str, lang: str) -> float:
    """执行用户代码并比对输出，返回正确率 0.0~1.0.

    复刻万维 NeedCheckProgramOutputContent / ProgramOutputContentRightPercent 逻辑。

    安全措施（已知安全债务）：
    - resource.setrlimit 限制 CPU/内存（仅 Unix）
    - subprocess timeout=5
    - 适用场景：可信学生代码（教学平台），不可信代码需 Docker 隔离
    """
    if not expected_output.strip():
        return 0.0

    user_output = _execute_code(user_code, lang)
    if user_output is None:
        return 0.0

    return _compare_outputs(user_output, expected_output)


def _run_limited(
    args: list[str],
    *,
    input_data: bytes | None = None,
    timeout: int = 5,
    preexec_fn=None,
) -> subprocess.CompletedProcess | None:
    """运行外部程序，超时时清理进程树."""
    creationflags = 0
    if os.name == "nt":
        creationflags = subprocess.CREATE_NEW_PROCESS_GROUP
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            creationflags |= subprocess.CREATE_NO_WINDOW

    proc = subprocess.Popen(
        args,
        stdin=subprocess.PIPE if input_data is not None else None,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        preexec_fn=preexec_fn,
        creationflags=creationflags,
    )
    try:
        stdout, stderr = proc.communicate(input=input_data, timeout=timeout)
    except subprocess.TimeoutExpired:
        try:
            if os.name == "nt":
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=2,
                )
            else:
                proc.kill()
        except Exception:
            pass
        try:
            proc.communicate(timeout=3)
        except Exception:
            pass
        return None
    return subprocess.CompletedProcess(args, proc.returncode, stdout, stderr)


def _execute_code(code: str, lang: str) -> str | None:
    """执行代码并返回 stdout，失败返回 None."""
    lang = lang.lower().strip()

    # 定义 preexec_fn 用于资源限制（仅 Unix）
    def _set_limits():
        try:
            import resource
            resource.setrlimit(resource.RLIMIT_CPU, (5, 5))
            resource.setrlimit(resource.RLIMIT_AS, (256 * 1024 * 1024, 256 * 1024 * 1024))
        except (ImportError, ValueError):
            pass  # Windows 或权限不足时跳过

    preexec = _set_limits if os.name != "nt" else None

    try:
        if lang == "python":
            proc = _run_limited(
                [sys.executable, "-c", code],
                timeout=5,
                preexec_fn=preexec,
            )
            if proc is not None and proc.returncode == 0:
                return proc.stdout.decode("utf-8", errors="replace")
            return None
        elif lang == "c":
            gcc = shutil.which("gcc")
            if not gcc:
                return None
            # 写临时文件编译执行
            with tempfile.NamedTemporaryFile(suffix=".c", delete=False, mode="w", encoding="utf-8") as f:
                f.write(code)
                c_path = f.name
            exe_path = c_path.replace(".c", ".exe" if os.name == "nt" else "")
            try:
                comp = subprocess.run(
                    [gcc, "-o", exe_path, c_path],
                    capture_output=True,
                    timeout=5,
                )
                if comp.returncode != 0:
                    return None
                proc = _run_limited(
                    [exe_path],
                    timeout=5,
                    preexec_fn=preexec,
                )
                if proc is not None and proc.returncode == 0:
                    return proc.stdout.decode("utf-8", errors="replace")
                return None
            finally:
                for p in (c_path, exe_path):
                    try:
                        os.unlink(p)
                    except OSError:
                        pass
    except (subprocess.TimeoutExpired, Exception):
        return None
    return None


def _compare_outputs(user_output: str, expected_output: str) -> float:
    """比较用户输出与预期输出，返回正确率.

    策略：去行尾空格 → 去空行 → 逐行比较（含浮点容差）。
    """
    user_lines = _normalize_output(user_output)
    expected_lines = _normalize_output(expected_output)

    if not expected_lines:
        return 1.0 if not user_lines else 0.0

    matched = 0
    for u, e in zip(user_lines, expected_lines):
        if _output_line_equal(u, e):
            matched += 1

    return matched / max(len(expected_lines), 1)


def _normalize_output(s: str) -> list[str]:
    """标准化输出：去行尾空格，去空行."""
    lines = s.strip().splitlines()
    lines = [line.rstrip() for line in lines]
    lines = [line for line in lines if line.strip()]
    return lines


def _output_line_equal(user_line: str, expected_line: str) -> bool:
    """比较两行输出是否相等（含浮点容差）."""
    if user_line == expected_line:
        return True
    # 浮点数容差比较
    try:
        return abs(float(user_line) - float(expected_line)) < 1e-6
    except (ValueError, TypeError):
        return False


def check_keywords(user_code: str, keywords: list[str]) -> float:
    """检查用户代码中的关键字匹配率.

    复刻万维 NeedCheckProgramKeyWord / ProgramKeyWordRightPercent 逻辑。
    在去注释后的代码中搜索每个关键字。

    Returns
    -------
    匹配率 0.0~1.0
    """
    if not keywords:
        return 1.0

    # 去注释后搜索
    cleaned = strip_comments(user_code, "python")  # 通用处理
    cleaned = strip_comments(cleaned, "c")

    matched = sum(1 for kw in keywords if kw in cleaned)
    return matched / max(len(keywords), 1)


def _get_program_weights() -> tuple[float, float, float]:
    """获取程序题评分权重（可配置）.

    Returns
    -------
    (output_weight, keyword_weight, code_weight)
    """
    try:
        cfg = ConfigManager().load_config()
        output_w = cfg.get("output_weight", 0.4)
        keyword_w = cfg.get("keyword_weight", 0.3)
        code_w = cfg.get("code_weight", 0.3)
        return output_w, keyword_w, code_w
    except Exception:
        return 0.4, 0.3, 0.3


def grade_exam(exam: Exam, strictness: str = "normal") -> list[Result]:
    """对整份试卷判分，返回每题结果.

    复刻万维分阶段评分逻辑：
    - 单选/判断：归一化精确匹配
    - 多选：全对满分，部分对半分，超集0分
    - 填空：两阶段匹配（一次匹配 + 断词匹配）
    - 程序设计：编译门槛 → 输出+关键字+代码加权

    Parameters
    ----------
    strictness:
        ``"strict"`` — 原始精确比对（不归一化）。
        ``"normal"`` — 归一化后精确比对（默认）。
        ``"lenient"`` — 归一化 + SequenceMatcher 阈值 0.8。
    """
    results = []
    for q in exam.questions:
        q_type = q.q_type
        score = q.score
        user_ans_raw = exam.get_answer(q.number)

        # 初始化结果字段
        is_correct = False
        is_half = False
        earned_score = 0.0

        if user_ans_raw:
            if strictness == "strict":
                is_correct, is_half, earned_score = _grade_strict(q, user_ans_raw)
            elif strictness == "lenient":
                is_correct, is_half, earned_score = _grade_lenient(q, user_ans_raw)
            else:
                is_correct, is_half, earned_score = _grade_normal(q, user_ans_raw)

        results.append(
            Result(
                question_number=q.number,
                q_type=q_type,
                score=score,
                user_answer=user_ans_raw if user_ans_raw else "未作答",
                correct_answer=q.correct_answer,
                is_correct=is_correct,
                question_text=q.text,
                earned_score=earned_score,
                is_half=is_half,
            )
        )
    return results


def _grade_strict(q, user_ans_raw: str) -> tuple[bool, bool, float]:
    """strict 模式：原始精确比对（不归一化，区分大小写）."""
    q_type = q.q_type
    score = q.score

    if q_type == "多选":
        is_correct, is_half = check_multi_select(user_ans_raw, q.correct_answer, False)
        if is_correct:
            return True, False, float(score)
        return False, False, 0.0

    if q_type in ("填空", "程序填空", "程序改错"):
        user_parts = re.split(r"[,;]+", user_ans_raw.strip())
        correct_parts = re.split(r"[,;]+", q.correct_answer.strip())
        user_parts = [p for p in user_parts if p]
        correct_parts = [p for p in correct_parts if p]
        if user_parts == correct_parts:
            return True, False, float(score)
        return False, False, 0.0

    if q_type in ("程序设计",):
        is_code_match = normalize_code(user_ans_raw) == normalize_code(q.correct_answer)
        if is_code_match:
            return True, False, float(score)
        return False, False, 0.0

    # 单选/判断
    if user_ans_raw.strip() == q.correct_answer.strip():
        return True, False, float(score)
    return False, False, 0.0


def _grade_lenient(q, user_ans_raw: str) -> tuple[bool, bool, float]:
    """lenient 模式：归一化 + SequenceMatcher 阈值 0.8."""
    q_type = q.q_type
    score = q.score
    user_ans = normalize_answer(user_ans_raw, q_type)
    correct_ans = normalize_answer(q.correct_answer, q_type)

    if q_type == "多选":
        is_correct, is_half = check_multi_select(user_ans_raw, q.correct_answer, q.partial_credit)
        if is_correct:
            return True, False, float(score)
        if is_half:
            return False, True, score / 2.0
        return False, False, 0.0

    if q_type in ("填空", "程序填空", "程序改错"):
        is_code = "程序" in q_type
        all_correct, matched, total = check_fill_in(user_ans, correct_ans, is_code=is_code)
        if all_correct:
            return True, False, float(score)
        if matched > 0 and q.partial_credit:
            return False, True, score * matched / max(total, 1)
        return False, False, 0.0

    if q_type in ("程序设计", "程序改错"):
        lang = getattr(q, "language", "python")
        user_clean = strip_comments(normalize_code_flexible(user_ans_raw), lang)
        correct_clean = strip_comments(normalize_code_flexible(q.correct_answer), lang)
        ratio = difflib.SequenceMatcher(None, user_clean, correct_clean).ratio()
        if ratio >= 0.8:
            return True, False, float(score)
        return False, False, 0.0

    # 单选/判断
    if user_ans == correct_ans:
        return True, False, float(score)
    return False, False, 0.0


def _grade_normal(q, user_ans_raw: str) -> tuple[bool, bool, float]:
    """normal 模式：归一化后精确比对（默认），含万维分阶段逻辑."""
    q_type = q.q_type
    score = q.score
    user_ans = normalize_answer(user_ans_raw, q_type)
    correct_ans = normalize_answer(q.correct_answer, q_type)

    # ── 单选/判断 ──
    if q_type in ("单选", "判断"):
        if user_ans == correct_ans:
            return True, False, float(score)
        return False, False, 0.0

    # ── 多选 ──
    if q_type == "多选":
        is_correct, is_half = check_multi_select(user_ans_raw, q.correct_answer, q.partial_credit)
        if is_correct:
            return True, False, float(score)
        if is_half:
            return False, True, score / 2.0
        return False, False, 0.0

    # ── 填空/程序填空/程序改错 ──
    if q_type in ("填空", "程序填空", "程序改错"):
        is_code = "程序" in q_type
        all_correct, matched, total = check_fill_in(user_ans, correct_ans, is_code=is_code)
        if all_correct:
            return True, False, float(score)
        if matched > 0 and q.partial_credit:
            return False, True, score * matched / max(total, 1)
        return False, False, 0.0

    # ── 程序设计（分阶段评分）──
    if q_type == "程序设计":
        lang = getattr(q, "language", "python")
        return _grade_program(q, user_ans_raw, lang)

    # 兜底
    if user_ans == correct_ans:
        return True, False, float(score)
    return False, False, 0.0


def _grade_program(q, user_ans_raw: str, lang: str) -> tuple[bool, bool, float]:
    """程序设计题分阶段评分，复刻万维编译→输出→关键字逻辑.

    阶段1: 编译检查（compile_check=True 时）→ 失败则 0 分
    阶段2: 多维度加权评分（输出 + 关键字 + 代码比对）
    """
    score = q.score

    # ── 阶段1: 编译检查 ──
    if q.compile_check:
        passed, msg = check_compilation(user_ans_raw, lang)
        if not passed:
            return False, False, 0.0

    # ── 阶段2: 多维度加权评分 ──
    output_w, keyword_w, code_w = _get_program_weights()

    # 如果没有输出和关键字检查，退化为纯代码比对
    has_output = q.output_check and q.expected_output.strip()
    has_keyword = q.keyword_check and q.keywords

    if not has_output and not has_keyword:
        # 纯代码比对（现有逻辑）
        is_correct = check_program_answer(user_ans_raw, q.correct_answer, lang)
        if is_correct:
            return True, False, float(score)
        return False, False, 0.0

    # 计算各维度得分
    earned = 0.0

    if has_output:
        output_rate = check_output(user_ans_raw, q.expected_output, lang)
        earned += output_rate * score * output_w

    if has_keyword:
        keyword_rate = check_keywords(user_ans_raw, q.keywords)
        earned += keyword_rate * score * keyword_w

    # 代码比对维度
    code_rate = 1.0 if check_program_answer(user_ans_raw, q.correct_answer, lang) else 0.0
    # 当有输出/关键字检查时，代码比对权重可能被压缩
    actual_code_w = code_w if (has_output or has_keyword) else 1.0
    earned += code_rate * score * actual_code_w

    # 限制不超过满分
    earned = min(earned, float(score))

    is_correct = earned >= score * 0.99  # 浮点容差
    is_half = not is_correct and earned > 0

    return is_correct, is_half, earned
