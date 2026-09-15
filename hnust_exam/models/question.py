"""题目数据模型."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Question:
    """一道题目的数据模型."""

    index: int  # 全局序号（从0开始）
    number: str  # 题号（如 "1", "T1"）
    q_type: str  # 题型（单选/多选/填空/判断/程序填空/程序改错/程序设计）
    text: str  # 题目内容
    options: dict[str, str] = field(default_factory=dict)  # {"A": "选项内容", ...}
    correct_answer: str = ""  # 正确答案
    score: int = 0  # 分值
    program_file: str = ""  # 程序文件名（如有）
    language: str = "python"  # 编程语言（python/c）
    images: dict[str, str] = field(default_factory=dict)  # {"图片1": "filename.gif", ...} 引自"图片"列
    partial_credit: bool = True  # 多选题是否允许半对半分
    compile_check: bool = False  # 程序题是否检查编译
    output_check: bool = False  # 程序题是否检查输出
    keyword_check: bool = False  # 程序题是否检查关键字
    expected_output: str = ""  # 程序题预期输出
    keywords: list[str] = field(default_factory=list)  # 程序题关键字列表

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

        # 解析"图片"列：格式 "图片1=file1.gif;图片2=file2.png"
        images = {}
        raw_images = str(data.get("图片", "")).strip()
        if raw_images and raw_images.lower() != "nan":
            for pair in raw_images.split(";"):
                pair = pair.strip()
                if "=" in pair:
                    ref, fname = pair.split("=", 1)
                    ref = ref.strip()
                    fname = fname.strip()
                    if ref and fname:
                        images[ref] = fname

        # 读取可选的新字段（向后兼容：列不存在时用默认值）
        def _get_bool(key_en: str, key_cn: str, default: bool = False) -> bool:
            val = str(data.get(key_en, data.get(key_cn, ""))).strip().lower()
            if val in ("1", "true", "是", "yes"):
                return True
            if val in ("0", "false", "否", "no"):
                return False
            return default

        def _get_str(key_en: str, key_cn: str, default: str = "") -> str:
            val = str(data.get(key_en, data.get(key_cn, default))).strip()
            return val if val.lower() != "nan" else default

        compile_check = _get_bool("CompileCheck", "编译检查", False)
        output_check = _get_bool("OutputCheck", "输出检查", False)
        keyword_check = _get_bool("KeywordCheck", "关键字检查", False)
        expected_output = _get_str("ExpectedOutput", "预期输出", "")
        keywords_raw = _get_str("Keywords", "关键字", "")
        keywords = [k.strip() for k in keywords_raw.replace("；", ";").split(";") if k.strip()]
        partial_credit = _get_bool("PartialCredit", "部分得分", True)

        return cls(
            index=global_index,
            number=_get_str("Number", "题号"),
            q_type=_get_str("Type", "题型"),
            text=_get_str("Text", "题目"),
            options=options,
            correct_answer=_get_str("CorrectAnswer", "正确答案"),
            score=score,
            program_file=_get_str("ProgramFile", "程序文件"),
            language=_get_str("Language", "语言", "python").lower() or "python",
            images=images,
            partial_credit=partial_credit,
            compile_check=compile_check,
            output_check=output_check,
            keyword_check=keyword_check,
            expected_output=expected_output,
            keywords=keywords,
        )
