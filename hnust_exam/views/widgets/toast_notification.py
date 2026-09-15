"""顶部 Toast 通知组件."""

from __future__ import annotations

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QLabel, QFrame, QHBoxLayout

from hnust_exam.utils.theme import Theme


class ToastWidget(QFrame):
    """显示在父窗口顶部居中的短时通知."""

    def __init__(self, parent, message: str, duration_ms: int = 2500) -> None:
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool |
            Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        c = Theme.get_current_colors()
        self.setStyleSheet(
            f"background-color: {c['PRIMARY']}; color: white; "
            "border-radius: 6px; padding: 8px 14px;"
        )

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 8, 14, 8)
        label = QLabel(message, self)
        label.setStyleSheet("color: white; font-size: 10pt; font-weight: bold;")
        layout.addWidget(label)

        self.adjustSize()
        self._position_at_top_center()

        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.close)
        if duration_ms > 0:
            self._timer.start(duration_ms)

    def showEvent(self, event) -> None:
        self._position_at_top_center()
        super().showEvent(event)

    def _position_at_top_center(self) -> None:
        if self.parentWidget() is None:
            return
        parent_rect = self.parentWidget().rect()
        x = parent_rect.center().x() - self.width() // 2
        self.move(max(0, x), 20)
