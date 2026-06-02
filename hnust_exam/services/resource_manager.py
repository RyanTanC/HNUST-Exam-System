"""统一管理题库资源路径.

所有题库资源（xlsx、图片、程序文件）的读取都通过此模块。
运行时只从 ~/.hnust_exam/question_bank/files/ 读取。
首次启动时从 _MEIPASS 初始化默认资源（一次性复制）。
"""

from __future__ import annotations

import logging
import os
import shutil

from hnust_exam.utils.constants import QUESTION_BANK_FILES_DIR

logger = logging.getLogger(__name__)


def get_question_bank_root() -> str:
    """返回题库运行时根目录：~/.hnust_exam/question_bank/files/"""
    return QUESTION_BANK_FILES_DIR


def list_exam_files() -> list[str]:
    """列出所有可用的试卷文件名（只从 AppData 读取）."""
    if not os.path.isdir(QUESTION_BANK_FILES_DIR):
        return []
    return [f for f in os.listdir(QUESTION_BANK_FILES_DIR) if f.endswith(".xlsx")]


def find_exam_file(exam_name: str) -> str | None:
    """查找试卷文件完整路径（只从 AppData 读取）."""
    path = os.path.join(QUESTION_BANK_FILES_DIR, exam_name + ".xlsx")
    return path if os.path.exists(path) else None


def get_image_dir() -> str:
    """返回试题图片目录路径."""
    return os.path.join(QUESTION_BANK_FILES_DIR, "试题图片")


def get_program_dir() -> str:
    """返回试题文件夹（程序文件）目录路径."""
    return os.path.join(QUESTION_BANK_FILES_DIR, "试题文件夹")


def ensure_initialized() -> None:
    """首次启动时，如果 AppData 中没有题库，从 _MEIPASS 初始化默认资源.

    打包环境：只在 AppData/QUESTION_BANK_FILES_DIR 为空时执行一次。
    开发环境（PyCharm）：每次启动都同步，自动拷贝新增的题库文件。
    初始化完成后，程序运行期间不再依赖 _MEIPASS。
    """
    import sys

    is_frozen = getattr(sys, "frozen", False)

    # 确定源目录：打包环境从 _MEIPASS，开发环境从项目源码
    if is_frozen:
        meipass = getattr(sys, "_MEIPASS", None)
        if not meipass:
            return
        source_dir = os.path.join(meipass, "题库")
        # 打包环境：缓存已有 xlsx 就跳过
        if os.path.isdir(QUESTION_BANK_FILES_DIR):
            existing = [f for f in os.listdir(QUESTION_BANK_FILES_DIR) if f.endswith(".xlsx")]
            if existing:
                logger.debug("题库已初始化，跳过（%d 个 xlsx 文件）", len(existing))
                return
    else:
        # 开发环境：从项目根目录的 题库/ 复制
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        source_dir = os.path.join(project_root, "题库")
        # 开发环境：先清空缓存，保证与源目录完全一致
        if os.path.isdir(QUESTION_BANK_FILES_DIR):
            shutil.rmtree(QUESTION_BANK_FILES_DIR)
            logger.info("开发环境：已清空题库缓存")

    if not os.path.isdir(source_dir):
        logger.warning("打包资源中未找到题库目录: %s", source_dir)
        return

    logger.info("同步题库: %s -> %s", source_dir, QUESTION_BANK_FILES_DIR)
    try:
        os.makedirs(QUESTION_BANK_FILES_DIR, exist_ok=True)
        copied = 0
        for fname in os.listdir(source_dir):
            src = os.path.join(source_dir, fname)
            if os.path.isfile(src) and fname.endswith(".xlsx"):
                dst = os.path.join(QUESTION_BANK_FILES_DIR, fname)
                shutil.copy2(src, dst)
                copied += 1
        # 复制子目录（试题图片、试题文件夹）
        for subdir in ("试题图片", "试题文件夹"):
            src_sub = os.path.join(source_dir, subdir)
            if os.path.isdir(src_sub):
                dst_sub = os.path.join(QUESTION_BANK_FILES_DIR, subdir)
                shutil.copytree(src_sub, dst_sub)
        logger.info("题库同步完成（%d 个 xlsx 文件）", copied)
    except Exception as e:
        logger.error("题库初始化失败: %s", e, exc_info=True)
