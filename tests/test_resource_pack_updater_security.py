"""资源包更新安全测试."""

from __future__ import annotations

import os
import zipfile

import pytest

from hnust_exam.services import resource_pack_updater


def test_safe_extract_rejects_zip_slip(tmp_path):
    zip_path = tmp_path / "bad.zip"
    extract_dir = tmp_path / "extract"
    outside_path = tmp_path / "outside.txt"

    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.writestr("../outside.txt", "owned")
        zf.writestr("safe.xlsx", "ok")

    with zipfile.ZipFile(zip_path, "r") as zf:
        with pytest.raises(ValueError, match="Zip path traversal"):
            resource_pack_updater._safe_extract_zip(zf, str(extract_dir))

    assert not outside_path.exists()


def test_safe_extract_allows_normal_members(tmp_path):
    zip_path = tmp_path / "good.zip"
    extract_dir = tmp_path / "extract"

    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.writestr("folder/safe.xlsx", "ok")

    with zipfile.ZipFile(zip_path, "r") as zf:
        resource_pack_updater._safe_extract_zip(zf, str(extract_dir))

    assert (extract_dir / "folder" / "safe.xlsx").read_text(encoding="utf-8") == "ok"
