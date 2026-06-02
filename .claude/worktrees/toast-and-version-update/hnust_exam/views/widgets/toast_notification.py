"""非模态顶部浮层通知组件.

使用方式：
    ToastWidget(parent, "消息内容").show()
    ToastWidget(parent, "消息内容", duration_ms=3000, toast_type="error").show()
"""

from __future__ import annotations

from PySide6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout, QWidget


class ToastWidget(QFrame):
    """顶部浮层通知，非模态，自动消褪.

    Parameters
    ----------
    parent : QWidget
        父窗口（通常是主窗口）。
    text : str
        通知文本。
    duration_ms : int
        显示时长（毫秒），默认 2000。
    toast_type : str
        通知类型："info"（默认）、"success"、"warning"、"error"，
        影响背景色。
    """

    _TYPEDEF = {
        "info": {"bg": "#0078d7", "icon": "ℹ️"},
        "success": {"bg": "#28a745", "icon": "✅"},
        "warning": {"bg": "#ffc107", "icon": "⚠️"},
        "error": {"bg": "#dc3545", "icon": "❌"},
    }

    def __init__(
        self,
        parent: QWidget,
        text: str,
        duration_ms: int = 2000,
        toast_type: str = "info",
    ):
        super().__init__(parent)
        style = self._TYPEDEF.get(toast_type, self._TYPEDEF["info"])

        # 窗口标志：无边框 + 工具窗口 + 置顶（避免被浏览器遮挡）
        self.setWindowFlags(
            Qt.FramelessWindowHint
            | Qt.Tool
            | Qt.WindowStaysOnTopHint
        )
        # 不抢占焦点（让浏览器窗口保持焦点）
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)
        # 鼠标穿透（不阻止用户与下方 UI 交互）
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        # 关闭时释放内存
        self.setAttribute(Qt.WA_DeleteOnClose, True)

        # 圆角背景
        self.setStyleSheet(f"""
            ToastWidget {{
                background-color: {style['bg']};
                border-radius: 8px;
                padding: 12px 24px;
            }}
        """)

        # 布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 12, 24, 12)

        self.label = QLabel(f"{style['icon']} {text}")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("color: white; font-size: 11pt; background: transparent;")
        layout.addWidget(self.label)

        # 自适应大小
        self.adjustSize()
        # 设置最小宽度，避免太窄
        if self.width() < 160:
            self.resize(160, self.height())

        # 定位到父窗口顶部居中
        self._position_at_top_center()

        # 自动关闭定时器
        self._close_timer = QTimer(self)
        self._close_timer.setSingleShot(True)
        self._close_timer.timeout.connect(self._fade_out)
        self._close_timer.start(duration_ms)

        # 淡入动画
        self._fade_in()

    def _position_at_top_center(self) -> None:
        """将自身定位到父窗口顶部居中."""
        parent = self.parentWidget()
        if parent:
            # 使用父窗口的全局坐标
            parent_global = parent.mapToGlobal(parent.rect().topLeft())
            parent_width = parent.width()
            x = parent_global.x() + (parent_width - self.width()) // 2
            y = parent_global.y() + 20  # 距父窗口顶部 20px
            self.move(x, y)

    def _fade_in(self) -> None:
        """淡入效果：透明度 0 → 1."""
        self.setWindowOpacity(0.0)
        self.show()
        self._anim = QPropertyAnimation(self, b"windowOpacity")
        self._anim.setDuration(200)
        self._anim.setStartValue(0.0)
        self._anim.setEndValue(1.0)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)
        self._anim.start()

    def _fade_out(self) -> None:
        """淡出效果：透明度 1 → 0，完成后关闭."""
        if self._anim:
            self._anim.stop()
        self._anim = QPropertyAnimation(self, b"windowOpacity")
        self._anim.setDuration(300)
        self._anim.setStartValue(self.windowOpacity())
        self._anim.setEndValue(0.0)
        self._anim.setEasingCurve(QEasingCurve.InCubic)
        self._anim.finished.connect(self.close)
        self._anim.start()

    def showEvent(self, event) -> None:
        """每次 show() 时刷新位置，适应窗口大小变化."""
        super().showEvent(event)
        self._position_at_top_center()
