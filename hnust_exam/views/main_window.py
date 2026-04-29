"""主窗口：QMainWindow + QStackedWidget 管理多页面."""

from __future__ import annotations

from PySide6.QtCore import Qt, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QIcon, QShortcut, QKeySequence, QFont
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QStackedWidget,
    QLabel,
    QStatusBar,
    QMessageBox,
)

from hnust_exam.utils.ui_helpers import themed_question

from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.constants import CURRENT_VERSION
from hnust_exam.utils.helpers import get_resource_path
from hnust_exam.utils.theme import Theme


class MainWindow(QMainWindow):
    """应用主窗口，管理页面切换."""

    PAGE_WELCOME = 0
    PAGE_SELECT = 1
    PAGE_EXAM = 2
    PAGE_RESULT = 3

    def __init__(self, config_mgr: ConfigManager, parent=None) -> None:
        super().__init__(parent)
        self.config_mgr = config_mgr

        self.setWindowTitle("HNUST仿真平台 | 免费使用 禁止售卖")
        self.setMinimumSize(900, 600)
        self.resize(1200, 800)

        # 设置窗口图标
        icon_path = get_resource_path("icon.ico")
        import os
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # 中央堆叠窗口
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        # 导入并创建各页面（延迟导入避免循环依赖）
        from hnust_exam.views.welcome_page import WelcomePage
        from hnust_exam.views.select_page import SelectPage
        from hnust_exam.views.exam_page import ExamPage
        from hnust_exam.views.result_page import ResultPage

        self.welcome_page = WelcomePage(self)
        self.select_page = SelectPage(self)
        self.exam_page = ExamPage(self)
        self.result_page = ResultPage(self)

        self.stack.addWidget(self.welcome_page)   # index 0
        self.stack.addWidget(self.select_page)     # index 1
        self.stack.addWidget(self.exam_page)       # index 2
        self.stack.addWidget(self.result_page)     # index 3

        # 水印标签
        self._watermark_left = QLabel("HNUST Exam · Free · 禁止售卖")
        self._watermark_left.setStyleSheet(
            f"color: {Theme.MUTED}; font-size: 7pt; padding: 2px;"
        )
        self._watermark_left.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.statusBar().addWidget(self._watermark_left)

        self._watermark_right = QLabel("该程序免费 禁止商用售卖")
        self._watermark_right.setStyleSheet(
            f"color: {Theme.MUTED}; font-size: 7pt; padding: 2px;"
        )
        self._watermark_right.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.statusBar().addPermanentWidget(self._watermark_right)

    def switch_to_page(self, page_index: int) -> None:
        """切换到指定页面."""
        self.stack.setCurrentIndex(page_index)

    def show_welcome(self) -> None:
        self.switch_to_page(self.PAGE_WELCOME)

    def show_select(self) -> None:
        self.switch_to_page(self.PAGE_SELECT)

    def show_exam(self) -> None:
        self.switch_to_page(self.PAGE_EXAM)

    def show_result(self) -> None:
        self.switch_to_page(self.PAGE_RESULT)

    def closeEvent(self, event) -> None:
        """关闭窗口时确认（考试进行中）."""
        exam_page = self.exam_page
        if exam_page.timer_running and not exam_page.exam_submitted:
            reply = themed_question(
                self, "确认退出",
                "考试正在进行中，确定要退出吗？\n未交卷的答案将不会保存。",
            )
            if reply != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            exam_page.timer_running = False
            exam_page._timer.stop()
            exam_page.backup_mgr.cleanup()
        event.accept()

    def _refresh_theme(self) -> None:
        """刷新全局主题（设置更改后调用）."""
        app = QApplication.instance()
        if app:
            from hnust_exam.app import _generate_stylesheet
            app.setStyleSheet(_generate_stylesheet())
            font = QFont("Microsoft YaHei", max(8, int(11 * Theme._font_scale)))
            app.setFont(font)

        # 刷新水印颜色
        c = Theme.get_current_colors()
        self._watermark_left.setStyleSheet(f"color: {c['MUTED']}; font-size: 7pt; padding: 2px;")
        self._watermark_right.setStyleSheet(f"color: {c['MUTED']}; font-size: 7pt; padding: 2px;")

        # 刷新当前页面
        current = self.stack.currentIndex()
        if current == self.PAGE_WELCOME:
            self._rebuild_welcome()
        elif current == self.PAGE_SELECT:
            self._rebuild_select()
        elif current == self.PAGE_EXAM:
            self.exam_page.refresh_theme()
        elif current == self.PAGE_RESULT:
            self.result_page.refresh_theme()

    def _rebuild_welcome(self) -> None:
        """重建欢迎页."""
        from hnust_exam.views.welcome_page import WelcomePage
        old = self.welcome_page
        self.welcome_page = WelcomePage(self)
        self.stack.removeWidget(old)
        self.stack.insertWidget(self.PAGE_WELCOME, self.welcome_page)
        self.stack.setCurrentIndex(self.PAGE_WELCOME)
        old.deleteLater()

    def _rebuild_select(self) -> None:
        """重建选卷页."""
        from hnust_exam.views.select_page import SelectPage
        old = self.select_page
        self.select_page = SelectPage(self)
        self.stack.removeWidget(old)
        self.stack.insertWidget(self.PAGE_SELECT, self.select_page)
        self.stack.setCurrentIndex(self.PAGE_SELECT)
        old.deleteLater()
