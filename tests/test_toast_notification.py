"""测试 ToastWidget 顶部通知组件."""
import pytest
from PySide6.QtWidgets import QMainWindow, QApplication, QLabel
from PySide6.QtCore import Qt, QTimer
from hnust_exam.views.widgets.toast_notification import ToastWidget


def test_toast_widget_creation(qtbot):
    """验证 ToastWidget 能正常创建和显示."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "测试通知")
    qtbot.addWidget(toast)

    # 验证 label 文本
    label = toast.findChild(QLabel)
    assert label is not None
    assert label.text() == "测试通知"

    # 验证窗口标志包含 FramelessWindowHint
    assert toast.windowFlags() & Qt.FramelessWindowHint


def test_toast_widget_auto_close(qtbot):
    """验证 ToastWidget 在指定时间后自动关闭."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "自动关闭测试", duration_ms=100)
    qtbot.addWidget(toast)
    toast.show()

    # 等待 200ms 确保定时器触发
    qtbot.wait(200)

    # 验证 toast 已关闭
    assert not toast.isVisible()


def test_toast_widget_position_at_top_center(qtbot):
    """验证 ToastWidget 定位在父窗口顶部居中."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "位置测试")
    qtbot.addWidget(toast)
    toast.show()

    # toast 的中心 x 应该在父窗口中心附近（允许偏移半个 toast 宽度）
    parent_center_x = window.rect().center().x()
    toast_center_x = toast.geometry().center().x()
    assert abs(toast_center_x - parent_center_x) < toast.width() // 2 + 1

    # toast 的顶部应该在父窗口顶部附近（距顶部 20px 左右）
    assert toast.y() == 20 or abs(toast.y() - 20) <= 1


def test_toast_widget_window_stays_on_top(qtbot):
    """验证 ToastWidget 设置了 WindowStaysOnTopHint 标志."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "置顶测试")
    qtbot.addWidget(toast)

    # WindowStaysOnTopHint 应被设置
    assert toast.windowFlags() & Qt.WindowStaysOnTopHint
