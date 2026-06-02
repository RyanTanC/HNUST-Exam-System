#!/usr/bin/env python3
"""题库资源包构建工具.

扫描 题库/ 目录，生成 manifest.json 并打包为 dist/question_bank.zip，
同时输出 dist/question_bank.zip.sha256 校验文件。

用法：
    python scripts/build_resource_pack.py
    python scripts/build_resource_pack.py --version 2026.05.29
    python scripts/build_resource_pack.py --clean

输出：
    dist/question_bank.zip        ← 上传到 GitHub Release
    dist/question_bank.zip.sha256 ← 上传到 GitHub Release
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import shutil
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

# ── 日志 ──────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("build_resource_pack")

# ── 路径常量 ──────────────────────────────────────────────────────────

_SCRIPT_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _SCRIPT_DIR.parent
_SOURCE_DIR = _PROJECT_ROOT / "题库"
_DIST_DIR = _PROJECT_ROOT / "dist"
_ZIP_PATH = _DIST_DIR / "question_bank.zip"
_HASH_PATH = _DIST_DIR / "question_bank.zip.sha256"

# zip 内部的根目录名（可选，resource_pack_updater 兼容单层根目录）
_ZIP_ROOT = "question_bank"


# ── 工具函数 ──────────────────────────────────────────────────────────


def sha256_file(path: Path) -> str:
    """计算文件 SHA256，返回 hex digest."""
    h = hashlib.sha256()
    try:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception as e:
        logger.error("SHA256 计算失败: %s -> %s", path, e)
        raise


def scan_source(source_dir: Path) -> dict[str, dict]:
    """扫描题库目录，返回 manifest.files 字典.

    key 是相对于 source_dir 的路径（正斜杠），不含 题库/ 前缀。
    跳过：manifest.json、隐藏文件、__pycache__。
    """
    files_info: dict[str, dict] = {}

    for p in sorted(source_dir.rglob("*")):
        if not p.is_file():
            continue
        name = p.name
        # 跳过不需要打包的文件
        if name in ("manifest.json", ".DS_Store", "Thumbs.db"):
            continue
        if name.startswith(".") or name.startswith("__"):
            continue

        rel = p.relative_to(source_dir).as_posix()
        stat = p.stat()
        file_hash = sha256_file(p)

        files_info[rel] = {
            "sha256": file_hash,
            "size": stat.st_size,
            "modified": int(stat.st_mtime),
        }
        logger.debug("  扫描: %s (%d bytes)", rel, stat.st_size)

    return files_info


def build_manifest(
    version: str,
    files_info: dict[str, dict],
) -> dict:
    """构建 manifest 字典."""
    return {
        "version": version,
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "total_files": len(files_info),
        "files": files_info,
    }


def build_zip(
    source_dir: Path,
    manifest: dict,
    zip_path: Path,
) -> None:
    """构建 zip 包.

    结构：
        question_bank.zip
        ├── manifest.json
        ├── xxx.xlsx
        ├── 试题图片/
        └── 试题文件夹/

    不含 题库/ 前缀层。
    """
    zip_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        # 写入 manifest.json
        manifest_bytes = json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8")
        zf.writestr(f"{_ZIP_ROOT}/manifest.json", manifest_bytes)
        logger.info("  写入 manifest.json (%d bytes)", len(manifest_bytes))

        # 写入所有文件
        for rel_path in sorted(manifest["files"]):
            src = source_dir / rel_path
            if not src.exists():
                logger.warning("  文件不存在，跳过: %s", src)
                continue
            arcname = f"{_ZIP_ROOT}/{rel_path}"
            zf.write(str(src), arcname)
            logger.debug("  打包: %s", arcname)

    logger.info("zip 构建完成: %s", zip_path)


def write_sha256_file(zip_path: Path, hash_path: Path) -> str:
    """计算 zip 的 SHA256 并写入校验文件."""
    file_hash = sha256_file(zip_path)
    hash_path.write_text(
        f"{file_hash}  {zip_path.name}\n",
        encoding="utf-8",
    )
    logger.info("SHA256: %s", file_hash[:16] + "...")
    return file_hash


def clean_dist(dist_dir: Path) -> None:
    """清理 dist 目录."""
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
        logger.info("已清理: %s", dist_dir)


# ── 主流程 ────────────────────────────────────────────────────────────


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="构建题库资源包 (question_bank.zip)",
    )
    parser.add_argument(
        "--version",
        type=str,
        default=None,
        help="资源包版本号，默认使用当天日期 YYYY.MM.DD",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="构建前清理 dist/ 目录",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="显示详细日志（DEBUG 级别）",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # 检查源目录
    if not _SOURCE_DIR.is_dir():
        logger.error("题库目录不存在: %s", _SOURCE_DIR)
        return 1

    # 版本号
    version = args.version or datetime.now().strftime("%Y.%m.%d")
    logger.info("资源包版本: %s", version)

    # 清理
    if args.clean:
        clean_dist(_DIST_DIR)

    # 1. 扫描
    logger.info("扫描题库目录: %s", _SOURCE_DIR)
    files_info = scan_source(_SOURCE_DIR)
    if not files_info:
        logger.error("未找到任何文件")
        return 1
    logger.info("扫描完成: %d 个文件", len(files_info))

    # 2. 生成 manifest
    logger.info("生成 manifest.json...")
    manifest = build_manifest(version, files_info)

    # 保存一份 manifest 到源目录（供本地参考）
    source_manifest = _SOURCE_DIR / "manifest.json"
    with open(source_manifest, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    logger.info("  源目录 manifest: %s", source_manifest)

    # 3. 构建 zip
    logger.info("构建 zip: %s", _ZIP_PATH)
    build_zip(_SOURCE_DIR, manifest, _ZIP_PATH)

    # 4. SHA256
    logger.info("计算 SHA256...")
    zip_hash = write_sha256_file(_ZIP_PATH, _HASH_PATH)

    # 5. 汇总
    zip_size_mb = _ZIP_PATH.stat().st_size / (1024 * 1024)
    logger.info("=" * 50)
    logger.info("构建完成！")
    logger.info("  版本:     %s", version)
    logger.info("  文件数:   %d", len(files_info))
    logger.info("  zip 大小: %.2f MB", zip_size_mb)
    logger.info("  SHA256:   %s", zip_hash)
    logger.info("  zip 路径: %s", _ZIP_PATH)
    logger.info("  校验文件: %s", _HASH_PATH)
    logger.info("=" * 50)
    logger.info("")
    logger.info("下一步：上传到 GitHub Release")
    logger.info("  1. 创建 tag:  git tag %s", version)
    logger.info("  2. 推送 tag:  git push origin %s", version)
    logger.info("  3. 创建 Release 并上传以下文件：")
    logger.info("     - %s", _ZIP_PATH.name)
    logger.info("     - %s", _HASH_PATH.name)

    return 0


if __name__ == "__main__":
    sys.exit(main())
