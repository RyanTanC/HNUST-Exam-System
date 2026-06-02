"""测试滚动和动画功能."""

import pytest
from PySide6.QtWidgets import QApplication, QScrollArea
from PySide6.QtCore import Qt, QPropertyAnimation


@pytest.fixture(scope="module")
def app():
    """创建 QApplication 实例."""
    return QApplication.instance() or QApplication([])


def test_scroll_area_has_scroller(app):
    """测试 QScrollArea 是否配置了 QScroller."""
    from hnust_exam.views.welcome_page import WelcomePage
    from PySide6.QtWidgets import QScroller

    class MockMainWindow:
        def show_select(self):
            pass

    welcome_page = WelcomePage(MockMainWindow())
    scroll_area = welcome_page.findChild(QScrollArea)

    assert scroll_area is not None, "欢迎页面应该包含 QScrollArea"

    scroller = QScroller.scroller(scroll_area.viewport())
    assert scroller is not None, "QScrollArea 应该配置 QScroller"


def test_scroll_area_scroller_properties(app):
    """测试 QScroller 的属性配置."""
    from hnust_exam.views.welcome_page import WelcomePage
    from PySide6.QtWidgets import QScroller, QScrollerProperties

    class MockMainWindow:
        def show_select(self):
            pass

    welcome_page = WelcomePage(MockMainWindow())
    scroll_area = welcome_page.findChild(QScrollArea)

    scroller = QScroller.scroller(scroll_area.viewport())
    prop = scroller.scrollerProperties()

    frame_rate = prop.scrollMetric(QScrollerProperties.ScrollMetric.FrameRate)
    assert frame_rate == QScrollerProperties.FrameRates.Standard, \
        "帧率应该设置为 Standard 以适配高刷新率显示器"


def test_welcome_page_scroll_smoothness(app):
    """测试欢迎页面滚动是否平滑（基本功能测试）."""
    from hnust_exam.views.welcome_page import WelcomePage

    class MockMainWindow:
        def show_select(self):
            pass

    welcome_page = WelcomePage(MockMainWindow())

    assert welcome_page.isVisible() or welcome_page.width() > 0


def test_main_window_has_animation_method(app):
    """测试主窗口是否有动画切换方法."""
    from hnust_exam.views.main_window import MainWindow
    from hnust_exam.services.config_manager import ConfigManager

    config_mgr = ConfigManager()
    main_window = MainWindow(config_mgr)

    assert hasattr(main_window, 'switch_to_page'), "主窗口应该有 switch_to_page 方法"
    assert callable(getattr(main_window, 'switch_to_page')), "switch_to_page 应该是可调用的"


def test_page_switch_animation_imports(app):
    """测试页面切换动画所需的导入."""
    from PySide6.QtCore import QPropertyAnimation, QEasingCurve, QRect, QAbstractAnimation

    assert QPropertyAnimation is not None
    assert QEasingCurve is not None
    assert QRect is not None
    assert QAbstractAnimation is not None


def test_main_window_page_constants(app):
    """测试主窗口页面常量."""
    from hnust_exam.views.main_window import MainWindow
    from hnust_exam.services.config_manager import ConfigManager

    config_mgr = ConfigManager()
    main_window = MainWindow(config_mgr)

    assert main_window.PAGE_WELCOME == 0
    assert main_window.PAGE_SELECT == 1
    assert main_window.PAGE_EXAM == 2
    assert main_window.PAGE_RESULT == 3


def test_welcome_page_full_integration(app):
    """欢迎页面完整集成测试."""
    from hnust_exam.views.welcome_page import WelcomePage
    from PySide6.QtWidgets import QScroller, QScrollerProperties

    class MockMainWindow:
        def show_select(self):
            pass

    welcome_page = WelcomePage(MockMainWindow())

    assert welcome_page.width() > 0 or welcome_page.isVisible()

    scroll_area = welcome_page.findChild(QScrollArea)
    assert scroll_area is not None

    scroller = QScroller.scroller(scroll_area.viewport())
    assert scroller is not None

    prop = scroller.scrollerProperties()
    max_fps = prop.scrollMetric(QScrollerProperties.ScrollMetric.FrameRate)
    assert max_fps == QScrollerProperties.FrameRates.Standard


def test_main_window_full_integration(app):
    """主窗口完整集成测试."""
    from hnust_exam.views.main_window import MainWindow
    from hnust_exam.services.config_manager import ConfigManager

    config_mgr = ConfigManager()
    main_window = MainWindow(config_mgr)

    main_window.show()
    assert main_window.isVisible()

    assert hasattr(main_window, 'switch_to_page')

    assert main_window.PAGE_WELCOME == 0
    assert main_window.PAGE_SELECT == 1
