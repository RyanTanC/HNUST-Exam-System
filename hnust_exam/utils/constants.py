"""全局常量定义."""

import os

CURRENT_VERSION = "v1.1.0"
GITHUB_USERNAME = "RyanTanC"
GITHUB_REPO_NAME = "HNUST-Exam-System"

_CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".hnust_exam")
CONFIG_FILE = os.path.join(_CONFIG_DIR, "config.json")
PROGRESS_FILE = os.path.join(_CONFIG_DIR, "progress.json")
SKIP_VERSION_FILE = os.path.join(_CONFIG_DIR, "skip_ver")

QUESTION_TYPE_ORDER = ["单选", "填空", "判断", "程序填空", "程序改错", "程序设计"]
PROGRAM_TYPES = {"程序设计", "程序填空", "程序改错"}

EXAM_TIME_SECONDS = 60 * 60

REQUIRED_COLUMNS = {"题号", "题型", "题目", "正确答案", "分值"}
