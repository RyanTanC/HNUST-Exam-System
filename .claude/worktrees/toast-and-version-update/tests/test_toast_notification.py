"""测试 ToastWidget 顶部通知组件."""
import pytest
from PySide6.QtWidgets import QMainWindow, QLabel
from PySide6.QtCore import Qt
from hnust_exam.views.widgets.toast_notification import ToastWidget


def test_toast_widget_creation(qtbot):
    """验证 ToastWidget 能正常创建和显示."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "测试通知")
    qtbot.addWidget(toast)

    # 验证 label 文本（包含图标前缀）
    label = toast.findChild(QLabel)
    assert label is not None
    assert "测试通知" in label.text()

    # 验证窗口标志包含 FramelessWindowHint
    assert toast.windowFlags() & Qt.FramelessWindowHint


def test_toast_widget_auto_close(qtbot):
    """验证 ToastWidget 在指定时间后自动关闭并销毁."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "自动关闭测试", duration_ms=100)

    # 不将 toast 注册到 qtbot（window 是父对象，window 关闭时会清理子对象）
    # 避免在 destroyed 后 teardown 再次操作已销毁的对象
    qtbot.waitSignal(toast.destroyed, timeout=500)


def test_toast_widget_position_at_top_center(qtbot):
    """验证 ToastWidget 定位在父窗口顶部居中."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    # 显示父窗口以获得有效的全局坐标
    window.show()
    qtbot.wait(100)

    toast = ToastWidget(window, "位置测试")
    qtbot.addWidget(toast)
    toast.show()

    # 使用全局坐标比较
    parent_top_left = window.mapToGlobal(window.rect().topLeft())
    parent_center_x = parent_top_left.x() + window.width() // 2
    toast_center_x = toast.geometry().center().x()
    assert abs(toast_center_x - parent_center_x) < toast.width() // 2 + 1

    # toast 的顶部应该在父窗口顶部下方 20px
    expected_y = parent_top_left.y() + 20
    assert abs(toast.geometry().y() - expected_y) <= 1


def test_toast_widget_window_stays_on_top(qtbot):
    """验证 ToastWidget 设置了 WindowStaysOnTopHint 标志."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "置顶测试")
    qtbot.addWidget(toast)

    # WindowStaysOnTopHint 应被设置
    assert toast.windowFlags() & Qt.WindowStaysOnTopHint
