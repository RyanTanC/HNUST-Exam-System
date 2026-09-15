"""题库资源包构建脚本测试.

重点锁定版本号格式约定：构建脚本产出的 tag 必须能被客户端
``resource_pack_updater._validate_version_tag()`` 接受，
否则打出的包会被客户端以"版本号格式错误"拒绝更新。
"""

from __future__ import annotations

import importlib.util
from datetime import datetime
from pathlib import Path

import pytest

_SCRIPT_PATH = (
    Path(__file__).resolve().parent.parent / "scripts" / "build_resource_pack.py"
)


def _load_builder():
    """按文件路径加载构建脚本（scripts/ 不是包，无法直接 import）."""
    spec = importlib.util.spec_from_file_location("build_resource_pack", _SCRIPT_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


builder = _load_builder()


class TestValidateVersion:
    def test_accepts_four_segment(self):
        assert builder.validate_version("2026.05.29.1430") is True

    def test_accepts_midnight(self):
        assert builder.validate_version("2026.01.15.0000") is True

    def test_rejects_three_segment(self):
        """三段日期（旧格式）必须被拒绝 —— 这是热更新失效的根因."""
        assert builder.validate_version("2026.05.29") is False

    def test_rejects_app_version_style(self):
        assert builder.validate_version("v1.1.8") is False

    def test_rejects_garbage(self):
        assert builder.validate_version("not-a-version") is False

    def test_rejects_empty(self):
        assert builder.validate_version("") is False


class TestDefaultVersion:
    def test_default_version_passes_validation(self):
        """不传 --version 时自动生成的版本号必须合法."""
        version = datetime.now().strftime("%Y.%m.%d.%H%M")
        assert builder.validate_version(version) is True


class TestPatternSingleSource:
    def test_shares_pattern_with_client(self):
        """构建脚本与客户端必须共用同一个正则，避免两边漂移."""
        from hnust_exam.utils.constants import QUESTION_BANK_VERSION_PATTERN

        assert builder.QUESTION_BANK_VERSION_PATTERN is QUESTION_BANK_VERSION_PATTERN


class TestScanSourceExcludes:
    """题库/ 是嵌套 git 仓库，rglob 会深入 .git/ 把整个仓库打进包里."""

    @pytest.fixture
    def source_dir(self, tmp_path):
        root = tmp_path / "题库"
        root.mkdir()
        (root / "PY选择题.xlsx").write_text("real", encoding="utf-8")
        (root / "试题文件夹").mkdir()
        (root / "试题文件夹" / "Prog00001.py").write_text("print(1)", encoding="utf-8")

        # 嵌套 git 仓库：其中的文件名本身不以点开头，仅按 name 过滤会漏掉
        (root / ".git").mkdir()
        (root / ".git" / "config").write_text("[core]", encoding="utf-8")
        (root / ".git" / "COMMIT_EDITMSG").write_text("msg", encoding="utf-8")
        (root / "__pycache__").mkdir()
        (root / "__pycache__" / "mod.cpython-313.pyc").write_bytes(b"\x00")
        return root

    def test_excludes_git_internals(self, source_dir):
        files = builder.scan_source(source_dir)
        assert not [k for k in files if k.startswith(".git/")]

    def test_excludes_pycache(self, source_dir):
        files = builder.scan_source(source_dir)
        assert not [k for k in files if "__pycache__" in k]

    def test_keeps_real_files(self, source_dir):
        files = builder.scan_source(source_dir)
        assert sorted(files) == ["PY选择题.xlsx", "试题文件夹/Prog00001.py"]

    def test_excludes_manifest_itself(self, source_dir):
        (source_dir / "manifest.json").write_text("{}", encoding="utf-8")
        files = builder.scan_source(source_dir)
        assert "manifest.json" not in files
