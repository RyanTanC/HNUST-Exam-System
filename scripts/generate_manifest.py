#!/usr/bin/env python3
"""
生成题库 manifest.json 文件

用法：
    python generate_manifest.py [版本号]

不传版本号时自动递增（读取现有 manifest version + 1）。
"""

import hashlib
import json
import os
import sys
from pathlib import Path


def calculate_sha256(file_path: str) -> str:
    """计算文件的 SHA256 哈希值"""
    sha256_hash = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                sha256_hash.update(chunk)
        return sha256_hash.hexdigest()
    except Exception as e:
        print(f"计算哈希失败: {file_path} -> {e}")
        return ""


def scan_directory(base_dir: str, relative_to: str) -> dict:
    """扫描目录，返回文件信息字典。

    key 是相对于 relative_to 的路径，使用正斜杠。
    跳过 manifest.json 本身。
    """
    files_info = {}
    base_path = Path(base_dir)

    for file_path in sorted(base_path.rglob("*")):
        if not file_path.is_file():
            continue
        # 跳过 manifest.json 自身
        if file_path.name == "manifest.json":
            continue
        # 计算相对路径
        relative_path = str(file_path.relative_to(relative_to))
        relative_path = relative_path.replace("\\", "/")

        file_hash = calculate_sha256(str(file_path))
        if file_hash:
            files_info[relative_path] = {
                "hash": file_hash,
                "size": file_path.stat().st_size,
            }

    return files_info


def main():
    # 获取项目根目录
    script_dir = Path(__file__).parent
    project_root = script_dir.parent
    
    # 题库目录
    question_bank_dir = project_root / "题库"
    if not question_bank_dir.exists():
        print(f"错误：题库目录不存在: {question_bank_dir}")
        sys.exit(1)
    
    # 版本号：命令行指定 > 自动递增 > 默认1
    manifest_path = question_bank_dir / "manifest.json"
    existing_version = 0
    if manifest_path.exists():
        try:
            with open(manifest_path, "r", encoding="utf-8") as f:
                existing_version = json.load(f).get("version", 0)
        except Exception:
            pass

    if len(sys.argv) > 1:
        try:
            version = int(sys.argv[1])
        except ValueError:
            print(f"错误：版本号必须是整数: {sys.argv[1]}")
            sys.exit(1)
    else:
        version = existing_version + 1
    
    print(f"扫描题库目录: {question_bank_dir}")
    print(f"版本号: {version}")
    
    # 扫描题库目录（key 相对于题库目录本身，不含 "题库/" 前缀）
    files_info = scan_directory(str(question_bank_dir), str(question_bank_dir))
    
    # 创建 manifest 结构
    manifest = {
        "version": version,
        "files": files_info
    }
    
    # 保存 manifest.json
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    
    print(f"生成 manifest.json 完成: {manifest_path}")
    print(f"文件数量: {len(files_info)}")
    
    # 显示部分文件信息
    print("\n部分文件示例:")
    for i, (file_path, info) in enumerate(list(files_info.items())[:5]):
        print(f"  {file_path}: {info['hash'][:16]}... ({info['size']} bytes)")
    
    if len(files_info) > 5:
        print(f"  ... 还有 {len(files_info) - 5} 个文件")


if __name__ == "__main__":
    main()