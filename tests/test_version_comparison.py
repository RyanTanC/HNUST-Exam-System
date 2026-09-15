"""测试版本号解析与比较工具函数."""
import pytest
from hnust_exam.utils.version import parse_version, compare_versions


# ── parse_version ──────────────────────────────────────────────────────

def test_parse_full_format():
    assert parse_version("2026.05.21.1430") == (2026, 5, 21, 1430)


def test_parse_date_only():
    assert parse_version("2026.05.21") == (2026, 5, 21, 0)


def test_parse_empty_string():
    assert parse_version("") == (0, 0, 0, 0)


def test_parse_invalid_format():
    assert parse_version("invalid") == (0, 0, 0, 0)
    assert parse_version("v1.0.0") == (0, 0, 0, 0)


def test_parse_partial_digits():
    assert parse_version("2026.ab.21.1430") == (2026, 0, 21, 1430)


def test_parse_single_number():
    assert parse_version("2026") == (2026, 0, 0, 0)


# ── compare_versions ───────────────────────────────────────────────────

def test_compare_equal():
    assert compare_versions("2026.05.21.1430", "2026.05.21.1430") == 0
    assert compare_versions("2026.05.21", "2026.05.21") == 0


def test_compare_local_older_by_date():
    assert compare_versions("2026.05.20.1430", "2026.05.21.1430") == -1
    assert compare_versions("2026.04.30.1200", "2026.05.01.0800") == -1


def test_compare_local_newer_by_date():
    assert compare_versions("2026.05.22.0800", "2026.05.21.1430") == 1


def test_compare_same_date_older_time():
    assert compare_versions("2026.05.21.0800", "2026.05.21.1430") == -1


def test_compare_same_date_newer_time():
    assert compare_versions("2026.05.21.1600", "2026.05.21.1430") == 1


def test_compare_old_format_vs_new():
    assert compare_versions("2026.05.21", "2026.05.21.1430") == -1
    assert compare_versions("2026.05.22", "2026.05.21.1430") == 1


def test_compare_empty_local():
    assert compare_versions("", "2026.05.21.1430") == -1


def test_compare_both_empty():
    assert compare_versions("", "") == 0


def test_compare_same_date_empty_time_vs_zero_time():
    assert compare_versions("2026.05.21", "2026.05.21.0000") == 0
