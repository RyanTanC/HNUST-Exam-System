"""控制台输出编码的回归测试.

背景（真实 bug）：``python_env.py`` 里几处 ``subprocess.run(..., text=True)``
在中文 Windows 上会抛 ``UnicodeDecodeError``。

根因链：

1. Windows 控制台程序（``where`` / ``py``）的输出走 **OEM 代码页**
   （中文系统 cp936，英文 cp437）。
2. Python 3.15+ 默认开启 UTF-8 模式，此时
   ``locale.getpreferredencoding()`` 返回 ``utf-8``。
3. ``text=True`` 就是拿这个值去解码 → GBK 字节被当 utf-8 解 → 炸。
4. **最阴的一点**：异常发生在 ``subprocess`` 的读取线程
   （``_readerthread`` → ``fh.read()``）里，主流程不崩，
   但 ``result.stdout`` / ``result.stderr`` 会变成空串 —— **结果被静默丢弃**。

实测（Python 3.14.3 + 中文 Windows）：``where python`` 必崩。
后果是 ``find_system_python()`` 拿不到 ``where`` 的结果，Python 探测降级。

修法：显式指定 OEM 代码页并配合 ``errors="replace"``（见 ``_console_encoding()``）。
"""

from __future__ import annotations

import ast
import codecs
import os
import subprocess
from pathlib import Path

import pytest

from hnust_exam.services import python_env
from hnust_exam.services.python_env import _console_encoding, find_system_python


# ───────── 编码辅助函数本身 ─────────


class TestConsoleEncoding:
    """``_console_encoding()`` 的取值必须是合法可用的编码名."""

    def test_is_a_valid_codec(self):
        """返回的编码名必须能被 codecs 解析（拼错会抛 LookupError）."""
        codecs.lookup(_console_encoding())

    @pytest.mark.skipif(os.name != "nt", reason="只在 Windows 上校验 OEM 代码页")
    def test_matches_windows_oem_code_page(self):
        """Windows 上必须等于真实的 OEM 代码页."""
        import ctypes

        expected = f"cp{ctypes.windll.kernel32.GetOEMCP()}"
        assert _console_encoding() == expected

    @pytest.mark.skipif(os.name == "nt", reason="非 Windows 上应回退 utf-8")
    def test_falls_back_to_utf8_off_windows(self):
        assert _console_encoding() == "utf-8"


# ───────── 真实调用不再丢结果 ─────────


@pytest.mark.skipif(os.name != "nt", reason="where 是 Windows 命令")
class TestConsoleOutputIsDecodable:
    """中文控制台输出必须能完整解出来，不能被丢成空串."""

    def test_chinese_error_message_survives(self):
        """核心回归：``where`` 的中文报错信息必须保留下来."""
        result = subprocess.run(
            ["where", "definitely_not_here_xyz_12345"],
            capture_output=True,
            encoding=_console_encoding(),
            errors="replace",
            timeout=15,
        )

        assert result.returncode != 0
        assert result.stderr.strip(), (
            "中文报错信息被丢成空串了 —— 说明又退回成按 utf-8 解码 GBK 字节"
        )

    def test_python_probe_does_not_raise(self):
        """``find_system_python()`` 必须能正常返回，不能因编码问题抛异常."""
        found = find_system_python()

        assert found is None or isinstance(found, str)
        if found:
            assert os.path.isfile(found), f"返回的路径不存在：{found}"


# ───────── 防止再犯：源码里不许出现裸 text=True ─────────


def _bare_text_true_lines(source: str) -> list[int]:
    """用 AST 找出所有 ``xxx(text=True)`` 调用的行号.

    不用正则是因为 ``text=True`` 会出现在注释和 docstring 里
    （本模块的说明文字就提到了它），正则会把它们误报成违规。
    """
    offenders: list[int] = []
    for node in ast.walk(ast.parse(source)):
        if not isinstance(node, ast.Call):
            continue
        for kw in node.keywords:
            if (
                kw.arg == "text"
                and isinstance(kw.value, ast.Constant)
                and kw.value.value is True
            ):
                offenders.append(node.lineno)
    return offenders


class TestNoBareTextTrue:
    """回归：源码里不能再出现裸 ``text=True``."""

    def test_python_env_has_no_bare_text_true(self):
        src = Path(python_env.__file__).read_text(encoding="utf-8")
        offenders = _bare_text_true_lines(src)

        assert not offenders, (
            f"python_env.py 第 {offenders} 行还在用裸 text=True；"
            "中文 Windows 上会抛 UnicodeDecodeError 并把结果静默丢成空串。"
            "请改用 encoding=_console_encoding(), errors='replace'。"
        )

    def test_scanner_actually_detects_violations(self):
        """自检：AST 扫描器本身要能抓到违规，否则上面的用例是假绿."""
        fake = "import subprocess\nsubprocess.run(['x'], text=True)\n"
        assert _bare_text_true_lines(fake) == [2]

        # 注释 / docstring 里的提及不算违规
        benign = '"""说明：不要写 text=True。"""\n# 也别写 text=True\n'
        assert _bare_text_true_lines(benign) == []
