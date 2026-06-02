"""性能测试."""

import time
import pytest
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer


def test_scroll_performance():
    """测试滚动性能."""
    app = QApplication.instance() or QApplication([])

    from hnust_exam.views.welcome_page import WelcomePage

    class MockMainWindow:
        def show_select(self):
            pass

    welcome_page = WelcomePage(MockMainWindow())
    welcome_page.show()

    start_time = time.time()
    for _ in range(100):
        app.processEvents()
        time.sleep(0.016)

    end_time = time.time()
    duration = end_time - start_time

    assert duration < 5.0, f"性能测试耗时过长: {duration}秒"


def test_animation_performance():
    """测试动画性能."""
    app = QApplication.instance() or QApplication([])

    from hnust_exam.views.main_window import MainWindow
    from hnust_exam.services.config_manager import ConfigManager

    config_mgr = ConfigManager()
    main_window = MainWindow(config_mgr)
    main_window.show()

    start_time = time.time()
    main_window.switch_to_page(main_window.PAGE_SELECT)

    for _ in range(30):
        app.processEvents()
        time.sleep(0.01)

    end_time = time.time()
    duration = end_time - start_time

    assert duration < 1.0, f"动画性能测试耗时过长: {duration}秒"
    assert main_window.stack.currentIndex() == main_window.PAGE_SELECT
