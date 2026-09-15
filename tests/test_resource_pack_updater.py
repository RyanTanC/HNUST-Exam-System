"""题库整包更新单元测试."""

from __future__ import annotations

import os
import hashlib
import zipfile
from unittest.mock import patch, MagicMock

import pytest

from hnust_exam.services import resource_pack_updater as rpu
from hnust_exam.services import update_lock
from hnust_exam.services.resource_pack_updater import (
    PackUpdateResult,
)
from hnust_exam.utils import constants


class TestValidateVersionTag:
    def test_valid_tag(self):
        assert rpu._validate_version_tag("2026.06.01.1823") is True

    def test_valid_tag_january(self):
        assert rpu._validate_version_tag("2026.01.15.0930") is True

    def test_invalid_tag_empty(self):
        assert rpu._validate_version_tag("") is False

    def test_invalid_tag_none(self):
        assert rpu._validate_version_tag(None) is False

    def test_invalid_tag_wrong_format(self):
        assert rpu._validate_version_tag("v1.0.0") is False

    def test_invalid_tag_no_time(self):
        assert rpu._validate_version_tag("2026.06.01") is False

    def test_tag_day_zero(self):
        assert rpu._validate_version_tag("2026.06.00.1823") is True

    def test_invalid_tag_garbage(self):
        assert rpu._validate_version_tag("not-a-version") is False


class TestParseSha256File:
    def test_plain_hash(self):
        h = "a" * 64
        assert rpu._parse_sha256_file(h) == h

    def test_hash_with_filename(self):
        h = "a" * 64
        assert rpu._parse_sha256_file(f"{h}  question_bank.zip") == h

    def test_multiline(self):
        h = "a" * 64
        content = f"{h}  question_bank.zip\nsome other line"
        assert rpu._parse_sha256_file(content) == h

    def test_empty(self):
        assert rpu._parse_sha256_file("") == ""

    def test_whitespace_only(self):
        assert rpu._parse_sha256_file("   \n  ") == ""


class TestSha256File:
    def test_computes_correct_hash(self, tmp_path):
        f = tmp_path / "test.bin"
        f.write_bytes(b"hello world")
        expected = hashlib.sha256(b"hello world").hexdigest()
        assert rpu._sha256_file(str(f)) == expected

    def test_empty_file(self, tmp_path):
        f = tmp_path / "empty.bin"
        f.write_bytes(b"")
        expected = hashlib.sha256(b"").hexdigest()
        assert rpu._sha256_file(str(f)) == expected

    def test_nonexistent_returns_empty(self):
        assert rpu._sha256_file("/nonexistent/path") == ""


class TestLocalVersion:
    def test_save_and_read(self, tmp_path):
        with patch.object(rpu, "_CURRENT_VERSION_FILE", str(tmp_path / "current_version")):
            rpu._save_local_version("2026.06.01.1823")
            assert rpu._get_local_version() == "2026.06.01.1823"

    def test_read_nonexistent(self, tmp_path):
        with patch.object(rpu, "_CURRENT_VERSION_FILE", str(tmp_path / "nonexistent")):
            assert rpu._get_local_version() == ""

    def test_overwrite(self, tmp_path):
        with patch.object(rpu, "_CURRENT_VERSION_FILE", str(tmp_path / "current_version")):
            rpu._save_local_version("2026.06.01.1823")
            rpu._save_local_version("2026.06.02.0915")
            assert rpu._get_local_version() == "2026.06.02.0915"


class TestMakeStagingDir:
    def test_creates_directory(self, tmp_path):
        staging = rpu._make_staging_dir(str(tmp_path))
        assert os.path.isdir(staging)
        assert staging.startswith(str(tmp_path))

    def test_staging_prefix(self, tmp_path):
        staging = rpu._make_staging_dir(str(tmp_path))
        assert "staging_" in os.path.basename(staging)


class TestSafeExtractZip:
    def test_rejects_path_traversal(self, tmp_path):
        zip_path = tmp_path / "bad.zip"
        extract_dir = tmp_path / "extract"

        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("../outside.txt", "owned")
            zf.writestr("safe.xlsx", "ok")

        with zipfile.ZipFile(zip_path, "r") as zf:
            with pytest.raises(ValueError, match="Zip path traversal"):
                rpu._safe_extract_zip(zf, str(extract_dir))

    def test_allows_normal_members(self, tmp_path):
        zip_path = tmp_path / "good.zip"
        extract_dir = tmp_path / "extract"

        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("folder/safe.xlsx", "ok")

        with zipfile.ZipFile(zip_path, "r") as zf:
            rpu._safe_extract_zip(zf, str(extract_dir))

        assert (extract_dir / "folder" / "safe.xlsx").read_text(encoding="utf-8") == "ok"


class TestPackUpdateResult:
    def test_default_values(self):
        r = PackUpdateResult(success=True, message="ok")
        assert r.success is True
        assert r.message == "ok"
        assert r.new_version == ""
        assert r.error_type == ""

    def test_error_result(self):
        r = PackUpdateResult(False, "失败", error_type="network")
        assert r.success is False
        assert r.error_type == "network"


class TestIsUpdating:
    def test_returns_false_when_not_locked(self):
        assert rpu.is_updating() is False


class TestForceUpdate:
    """验证 force_update 参数的行为."""

    def test_force_update_passes_through(self):
        """check_pack_update_async 接受 force_update 参数，不抛出异常."""
        with (
            patch.object(rpu, "_do_update", return_value=PackUpdateResult(True, "ok")) as mock,
            patch.object(rpu, "Thread", new=lambda target, daemon: MagicMock(
                start=lambda: mock(started=True),
            )),
        ):
            called = []

            def cb(result):
                called.append(result)

            rpu.check_pack_update_async(
                callback=cb, progress_callback=None, force_update=True,
            )

    def test_force_update_normal_check_no_side_effects(self):
        """不传 force_update 时默认 False，与现有行为一致."""
        with (
            patch.object(rpu, "_do_update", return_value=PackUpdateResult(True, "ok")) as mock,
            patch.object(rpu, "Thread", new=lambda target, daemon: MagicMock(
                start=lambda: mock(started=True),
            )),
        ):
            called = []

            def cb(result):
                called.append(result)

            rpu.check_pack_update_async(callback=cb)


class TestStatusCallback:
    """验证 status_callback 在 _do_update 各阶段被调用."""

    def test_status_called_during_update(self, tmp_path, monkeypatch):
        status_log = []

        monkeypatch.setattr(rpu, "QUESTION_BANK_DIR", str(tmp_path / "question_bank"))
        monkeypatch.setattr(rpu, "QUESTION_BANK_FILES_DIR", str(tmp_path / "question_bank" / "files"))
        monkeypatch.setattr(rpu, "_CURRENT_VERSION_FILE", str(tmp_path / "current_version"))
        # 下面三个常量默认指向真实用户目录 ~/.hnust_exam，不隔离会在测试中
        # 创建/删除用户的真实 files_backup 并把假数据写进真实 manifest.json。
        # 同时它们都是绝对路径，不隔离还会因跨盘符 os.replace 报 WinError 17。
        monkeypatch.setattr(
            rpu, "_BACKUP_DIR", str(tmp_path / "question_bank" / "files_backup")
        )
        monkeypatch.setattr(constants, "MANIFEST_FILE", str(tmp_path / "manifest.json"))
        monkeypatch.setattr(update_lock, "_LOCK_DIR", str(tmp_path / ".locks"))

        # mock 一个远程 release 和下载
        def mock_fetch(session):
            return {
                "zip_url": "http://example.com/qb.zip",
                "zip_size": 100,
                "hash_url": "",
                "tag_name": "2026.06.04.1200",
            }

        # 让下载创建一个小 zip 并写入目标目录
        _download_count = [0]

        def mock_download(session, url, path, expected_size=0, progress_callback=None):
            _download_count[0] += 1
            # 创建一个小 zip
            zip_path = path
            with zipfile.ZipFile(zip_path, "w") as zf:
                zf.writestr("test.xlsx", "hello")
            if progress_callback:
                progress_callback(100, 100)
            return True

        def mock_validate(tag):
            return True

        monkeypatch.setattr(rpu, "_fetch_release_assets", mock_fetch)
        monkeypatch.setattr(rpu, "_download_to_file", mock_download)
        monkeypatch.setattr(rpu, "_validate_version_tag", mock_validate)

        os.makedirs(str(tmp_path / "question_bank" / "files"), exist_ok=True)

        result = rpu._do_update(
            status_callback=status_log.append,
            force_update=False,
        )

        assert result.success is True
        assert "校验文件完整性..." in status_log
        assert "解压文件中..." in status_log
        assert "更新文件中..." in status_log
