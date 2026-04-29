"""HNUST 仿真平台入口."""

import os
import sys

# PyInstaller --onefile 模式下，设置 Qt 平台插件路径
if getattr(sys, "frozen", False):
    os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = os.path.join(
        sys._MEIPASS, "PySide6", "plugins"
    )

from hnust_exam.app import run

if __name__ == "__main__":
    run()
