"""应用入口：初始化 QApplication、主题、主窗口."""

from __future__ import annotations

import os
import sys
import traceback
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import QApplication

from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.helpers import get_resource_path
from hnust_exam.utils.theme import Theme
from hnust_exam.views.main_window import MainWindow


def run() -> None:
    """启动应用."""
    # Windows 高 DPI 感知（必须在 QApplication 之前设置）
    if sys.platform == "win32":
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-Monitor DPI Aware
        except Exception:
            pass

    # 高 DPI 支持
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("HNUST仿真平台")
    app.setOrganizationName("HNUST")

    # 设置应用全局图标（任务栏 + 窗口）
    _icon_path = get_resource_path("icon.ico")
    if os.path.exists(_icon_path):
        app.setWindowIcon(QIcon(_icon_path))

    # 加载配置
    config_mgr = ConfigManager()
    cfg = config_mgr.load_config()

    # 应用主题
    Theme._font_scale = cfg.get("font_scale", 1.0)
    Theme.update_fonts()
    Theme.set_dark_mode(cfg.get("dark_mode", False))

    # 设置全局字体
    font = QFont("Microsoft YaHei", max(8, int(11 * Theme._font_scale)))
    font.setHintingPreference(QFont.PreferNoHinting)
    app.setFont(font)

    # 设置全局样式表
    app.setStyleSheet(_generate_stylesheet())

    # 创建主窗口
    main_window = MainWindow(config_mgr)
    main_window.show()

    # 全局异常钩子
    def _excepthook(exc_type, exc_value, exc_tb):
        _log_crash(exc_type, exc_value, exc_tb)

    sys.excepthook = _excepthook

    sys.exit(app.exec())


def _generate_stylesheet() -> str:
    """根据当前主题生成全局 QSS 样式表."""
    c = Theme.get_current_colors()
    s = Theme._font_scale
    fs = max(8, int(11 * s))
    fs_small = max(7, int(9 * s))
    fs_title = max(9, int(12 * s))

    return f"""
        QMainWindow {{
            background-color: {c["BG"]};
        }}
        QWidget {{
            color: {c["TEXT"]};
        }}
        QLabel {{
            color: {c["TEXT"]};
            border: none;
        }}
        QFrame {{
            border: none;
        }}
        QPushButton {{
            background-color: {c["SURFACE"]};
            color: {c["TEXT"]};
            border: 1px solid {c["BORDER"]};
            padding: 6px 16px;
            border-radius: 4px;
            font-size: {fs}pt;
        }}
        QPushButton:hover {{
            background-color: {c["PRIMARY"]};
            color: white;
            border: none;
        }}
        QPushButton:pressed {{
            background-color: {c["PRIMARY_HOVER"]};
        }}
        QPushButton:disabled {{
            color: {c["MUTED"]};
            background-color: {c["PROGRESS_BG"]};
            border: 1px solid {c["BORDER"]};
        }}
        QListWidget {{
            background-color: {c["WHITE"]};
            color: {c["TEXT"]};
            border: 1px solid {c["BORDER"]};
            border-radius: 4px;
            font-size: {fs}pt;
        }}
        QListWidget::item:selected {{
            background-color: {c["PRIMARY"]};
            color: white;
        }}
        QListWidget::item:hover {{
            background-color: {c["NAV_ACTIVE"]};
        }}
        QLineEdit {{
            background-color: {c["INPUT_BG"]};
            color: {c["TEXT"]};
            border: 1px solid {c["BORDER"]};
            border-radius: 4px;
            padding: 4px 8px;
            font-size: {fs}pt;
        }}
        QTextEdit {{
            background-color: {c["INPUT_BG"]};
            color: {c["TEXT"]};
            border: 1px solid {c["BORDER"]};
            border-radius: 4px;
            padding: 4px 8px;
            font-family: Consolas, monospace;
            font-size: {fs}pt;
        }}
        QScrollArea {{
            border: none;
            background-color: transparent;
        }}
        QScrollBar:vertical {{
            background-color: {c["SURFACE"]};
            width: 10px;
            border-radius: 5px;
        }}
        QScrollBar::handle:vertical {{
            background-color: {c["BORDER"]};
            border-radius: 5px;
            min-height: 30px;
        }}
        QScrollBar::handle:vertical:hover {{
            background-color: {c["PRIMARY"]};
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0;
        }}
        QProgressBar {{
            background-color: {c["PROGRESS_BG"]};
            border: none;
            border-radius: 4px;
            text-align: center;
            color: {c["MUTED"]};
            font-size: {fs_small}pt;
        }}
        QProgressBar::chunk {{
            background-color: {c["SUCCESS"]};
            border-radius: 4px;
        }}
    """


def _log_crash(exc_type, exc_value, exc_tb) -> None:
    """记录崩溃日志."""
    try:
        from hnust_exam.utils.helpers import get_log_dir
        from hnust_exam.utils.constants import CURRENT_VERSION

        log_dir = get_log_dir()
        log_name = f"exam_crash_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path = f"{log_dir}/{log_name}"

        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"{'=' * 60}\n")
            f.write(f"错误: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"版本: {CURRENT_VERSION}\n")
            f.write(f"系统: {sys.platform} / {sys.version}\n")
            f.write(f"{'-' * 60}\n")
            f.write("".join(traceback.format_exception(exc_type, exc_value, exc_tb)))
            f.write("\n")
    except Exception:
        pass
