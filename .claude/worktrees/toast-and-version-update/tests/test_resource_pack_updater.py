"""测试 resource_pack_updater 的版本检查逻辑."""
import os
from unittest.mock import patch
from hnust_exam.services.resource_pack_updater import _do_update


def _prepare_extracted_dir(tmp_path):
    """在 staging/extracted 下创建假 xlsx 文件，通过解压验证."""
    extract_dir = os.path.join(str(tmp_path), "staging", "extracted")
    os.makedirs(extract_dir, exist_ok=True)
    xlsx_path = os.path.join(extract_dir, "test.xlsx")
    with open(xlsx_path, "w") as f:
        f.write("")
    return extract_dir


def _setup_paths(monkeypatch, rpu_module, tmp_path):
    """替换路径常量为临时目录."""
    monkeypatch.setattr(rpu_module, "QUESTION_BANK_DIR", str(tmp_path))
    monkeypatch.setattr(rpu_module, "QUESTION_BANK_FILES_DIR", str(tmp_path / "files"))
    monkeypatch.setattr(rpu_module, "_STAGING_DIR", str(tmp_path / "staging"))
    monkeypatch.setattr(rpu_module, "_BACKUP_DIR", str(tmp_path / "files_backup"))
    monkeypatch.setattr(rpu_module, "_CURRENT_VERSION_FILE", str(tmp_path / "current_version"))


def test_do_update_local_older_than_remote(monkeypatch, tmp_path):
    """本地版本比远程旧 → 执行更新."""
    from hnust_exam.services import resource_pack_updater as rpu

    _setup_paths(monkeypatch, rpu, tmp_path)
    (tmp_path / "current_version").write_text("2026.05.20.0800", encoding="utf-8")

    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",
    }

    _prepare_extracted_dir(tmp_path)

    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        with patch.object(rpu, "_download_to_file", return_value=True):
            with patch.object(rpu, "zipfile"):
                with patch.object(rpu, "_safe_rmtree"):
                    with patch.object(rpu, "_safe_remove"):
                        with patch.object(rpu, "_regenerate_manifest"):
                            result = _do_update()

    assert result.success, f"更新应当成功，但得到: {result.message}"
    assert "更新" in result.message, f"消息应含'更新'，实际: {result.message}"
    assert result.new_version == "2026.05.21.1430"


def test_do_update_local_equal_to_remote(monkeypatch, tmp_path):
    """本地版本等于远程 → 不更新."""
    from hnust_exam.services import resource_pack_updater as rpu

    _setup_paths(monkeypatch, rpu, tmp_path)
    (tmp_path / "current_version").write_text("2026.05.21.1430", encoding="utf-8")

    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",
    }

    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        with patch.object(rpu, "_download_to_file", side_effect=AssertionError("不该下载")):
            result = _do_update()

    assert result.success, f"应当返回成功，但得到: {result.message}"
    assert "已是最新" in result.message, f"消息应含'已是最新'，实际: {result.message}"


def test_do_update_same_date_older_time(monkeypatch, tmp_path):
    """同日期但本地时间更旧 → 更新."""
    from hnust_exam.services import resource_pack_updater as rpu

    _setup_paths(monkeypatch, rpu, tmp_path)
    (tmp_path / "current_version").write_text("2026.05.21.0800", encoding="utf-8")

    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",
    }

    _prepare_extracted_dir(tmp_path)

    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        with patch.object(rpu, "_download_to_file", return_value=True):
            with patch.object(rpu, "zipfile"):
                with patch.object(rpu, "_safe_rmtree"):
                    with patch.object(rpu, "_safe_remove"):
                        with patch.object(rpu, "_regenerate_manifest"):
                            result = _do_update()

    assert result.success, f"更新应当成功，但得到: {result.message}"
    assert "更新" in result.message, f"消息应含'更新'，实际: {result.message}"


def test_do_update_same_date_newer_time(monkeypatch, tmp_path):
    """同日期但本地时间更新 → 不更新."""
    from hnust_exam.services import resource_pack_updater as rpu

    _setup_paths(monkeypatch, rpu, tmp_path)
    (tmp_path / "current_version").write_text("2026.05.21.1600", encoding="utf-8")

    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",
    }

    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        with patch.object(rpu, "_download_to_file", side_effect=AssertionError("不该下载")):
            result = _do_update()

    assert result.success, f"应当返回成功（已是最新），但得到: {result.message}"
    assert "已是最新" in result.message, f"消息应含'已是最新'，实际: {result.message}"


def test_do_update_old_format_compatibility(monkeypatch, tmp_path):
    """旧格式 yyyy.mm.dd vs 新格式 yyyy.mm.dd.tttt → 正确比较."""
    from hnust_exam.services import resource_pack_updater as rpu

    _setup_paths(monkeypatch, rpu, tmp_path)
    (tmp_path / "current_version").write_text("2026.05.21", encoding="utf-8")

    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",
    }

    _prepare_extracted_dir(tmp_path)

    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        with patch.object(rpu, "_download_to_file", return_value=True):
            with patch.object(rpu, "zipfile"):
                with patch.object(rpu, "_safe_rmtree"):
                    with patch.object(rpu, "_safe_remove"):
                        with patch.object(rpu, "_regenerate_manifest"):
                            result = _do_update()

    assert result.success, f"更新应当成功，但得到: {result.message}"
    assert "更新" in result.message, f"消息应含'更新'，实际: {result.message}"