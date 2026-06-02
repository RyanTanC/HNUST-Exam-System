"""全局常量定义."""

import os

CURRENT_VERSION = "v1.1.5"
GITHUB_USERNAME = "RyanTanC"
GITHUB_REPO_NAME = "HNUST-Exam-System"

# 题库热更新（独立 Gitee 仓库）
GITEE_USERNAME = "ryan-tanc"
GITEE_REPO_NAME = "hnust-computer-exam-tiku"

_CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".hnust_exam")
CONFIG_DIR = _CONFIG_DIR  # backward-compatible alias
CONFIG_FILE = os.path.join(_CONFIG_DIR, "config.json")
PROGRESS_FILE = os.path.join(_CONFIG_DIR, "progress.json")
SKIP_VERSION_FILE = os.path.join(_CONFIG_DIR, "skip_ver")
TELEMETRY_QUEUE_FILE = os.path.join(_CONFIG_DIR, "telemetry_queue.json")

# 用户本地题库缓存目录（持久化，断网可用）
QUESTION_BANK_DIR = os.path.join(_CONFIG_DIR, "question_bank")
QUESTION_BANK_FILES_DIR = os.path.join(QUESTION_BANK_DIR, "files")
MANIFEST_FILE = os.path.join(QUESTION_BANK_DIR, "manifest.json")

LOG_DIR = os.path.join(_CONFIG_DIR, "logs")

QUESTION_TYPE_ORDER = ["单选", "多选", "填空", "判断", "程序填空", "程序改错", "程序设计"]
PROGRAM_TYPES = {"程序设计", "程序填空", "程序改错"}

EXAM_TIME_SECONDS = 60 * 60

# 图片目录（相对于项目根目录）
IMAGES_DIR = "题库/试题图片"

REQUIRED_COLUMNS = {"题号", "题型", "题目", "正确答案", "分值"}

# 可选列（向后兼容：不存在时使用默认值）
OPTIONAL_COLUMNS = {
    "CompileCheck", "OutputCheck", "KeywordCheck", "Keywords", "ExpectedOutput", "PartialCredit",
    "编译检查", "输出检查", "关键字检查", "关键字", "预期输出", "部分得分",
}

# 匿名使用统计
TELEMETRY_BASE_URL = "https://hnust-exam-telemetry.hnust-exam-stats.workers.dev"