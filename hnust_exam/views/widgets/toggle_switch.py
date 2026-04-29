"""自定义开关控件."""

from __future__ import annotations

from PySide6.QtCore import Qt, QPropertyAnimation, Property, QRectF, Signal
from PySide6.QtGui import QPainter, QColor, QPen
from PySide6.QtWidgets import QWidget


class ToggleSwitch(QWidget):
    """带动画的开关控件."""

    toggled = Signal(bool)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFixedSize(52, 26)
        self._checked = False
        self._circle_pos = 4.0
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        self._animation = QPropertyAnimation(self, b"circle_pos")
        self._animation.setDuration(200)

    def isChecked(self) -> bool:
        return self._checked

    def setChecked(self, checked: bool) -> None:
        self._checked = checked
        self._circle_pos = 26.0 if checked else 4.0
        self.update()

    def get_circle_pos(self) -> float:
        return self._circle_pos

    def set_circle_pos(self, pos: float) -> None:
        self._circle_pos = pos
        self.update()

    circle_pos = Property(float, get_circle_pos, set_circle_pos)

    def mousePressEvent(self, event) -> None:
        self._checked = not self._checked
        self._animation.setStartValue(self._circle_pos)
        self._animation.setEndValue(26.0 if self._checked else 4.0)
        self._animation.start()
        self.toggled.emit(self._checked)

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # 背景
        bg_color = QColor(76, 175, 80) if self._checked else QColor(189, 189, 189)
        painter.setBrush(bg_color)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(QRectF(0, 0, 52, 26), 13, 13)

        # 滑块阴影
        painter.setBrush(QColor(0, 0, 0, 22))
        painter.drawEllipse(QRectF(self._circle_pos, 3, 20, 20))

        # 滑块
        painter.setBrush(QColor(255, 255, 255))
        painter.drawEllipse(QRectF(self._circle_pos + 1, 3, 20, 20))

        painter.end()
