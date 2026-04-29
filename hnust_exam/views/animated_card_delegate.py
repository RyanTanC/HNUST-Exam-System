"""抗锯齿文字 + 选中卡片弹性弹出动画的自定义代理."""

import math
import time

from PySide6.QtCore import Qt, QSize, QRectF, QTimer
from PySide6.QtGui import QPainter, QColor, QPen, QFont, QFontMetrics
from PySide6.QtWidgets import QStyledItemDelegate, QStyle


class AnimatedCardDelegate(QStyledItemDelegate):
    """抗锯齿文字 + 选中卡片弹性弹出动画"""

    DURATION = 0.38
    NORMAL_PAD_V = 4
    NORMAL_PAD_H = 16
    SELECTED_PAD_V = 2
    SELECTED_PAD_H = 10
    EXPAND_EXTRA = 10
    SHADOW_SPACE = 4
    EXPAND_HEIGHT = 8           # 选中时额外增加的行高（推开相邻卡片）

    def __init__(self, colors: dict, parent=None):
        super().__init__(parent)
        self._c = colors
        self._progress: dict = {}       # row → 0.0~1.0
        self._t0: dict = {}             # row → 动画起始时间戳
        self._active: set = set()       # 正在动画的行

        self._timer = QTimer(self)
        self._timer.setInterval(16)     # ≈60 fps
        self._timer.timeout.connect(self._tick)

    # ──────────────── 公开接口 ────────────────

    def animateTo(self, new_row: int, old_row: int = -1):
        """选择变化时调用，触发动画"""
        now = time.time()
        if new_row >= 0:
            self._progress[new_row] = self._progress.get(new_row, 0.0)
            self._t0[new_row] = now
            self._active.add(new_row)
        if old_row >= 0 and old_row != new_row:
            self._progress[old_row] = self._progress.get(old_row, 1.0)
            self._t0[old_row] = now
            self._active.add(old_row)
        if self._active and not self._timer.isActive():
            self._timer.start()

    # ──────────────── 动画引擎 ────────────────

    def _tick(self):
        now = time.time()
        view = self.parent()
        done = set()

        for row in self._active:
            t = min((now - self._t0.get(row, now)) / self.DURATION, 1.0)
            eased = self._elastic_out(t)

            is_sel = False
            if view and view.model() and view.selectionModel():
                is_sel = view.selectionModel().isSelected(
                    view.model().index(row, 0)
                )
            self._progress[row] = eased if is_sel else 1.0 - eased

            if t >= 1.0:
                done.add(row)
                if not is_sel:
                    self._progress.pop(row, None)

        self._active -= done
        if view:
            # 触发重新布局，使 sizeHint 变化生效（推开相邻卡片）
            view.scheduleDelayedItemsLayout()
        if not self._active:
            self._timer.stop()

    @staticmethod
    def _elastic_out(t: float) -> float:
        """弹性缓出 — 先冲过头再弹回来"""
        if t <= 0.0:
            return 0.0
        if t >= 1.0:
            return 1.0
        p = 0.35
        s = p / 4.0
        return pow(2, -10 * t) * math.sin((t - s) * 2 * math.pi / p) + 1.0

    # ──────────────── 尺寸 & 绘制 ────────────────

    def sizeHint(self, option, index):
        """选中条目随动画增大行高，推开相邻卡片"""
        base = super().sizeHint(option, index)
        fm = QFontMetrics(option.font)
        h = base.height() if base.height() > 0 else fm.height()
        row = index.row()
        p = self._progress.get(row, 0.0)
        # 动态高度：选中时多出 EXPAND_HEIGHT * p
        extra = self.EXPAND_HEIGHT * p
        return QSize(
            max(base.width(), 100) + self.SHADOW_SPACE * 2,
            int(h + self.EXPAND_EXTRA + self.SHADOW_SPACE * 2 + 8 + extra),
        )

    def paint(self, painter: QPainter, option, index):
        painter.save()

        # ★ 抗锯齿
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)

        row = index.row()
        p = self._progress.get(row, 0.0)
        selected = bool(option.state & QStyle.StateFlag.State_Selected)
        hovered = bool(option.state & QStyle.StateFlag.State_MouseOver)

        if selected and p < 1.0 and row not in self._active:
            p = 1.0

        # ---- 计算卡片矩形 ----
        full = QRectF(option.rect).adjusted(
            self.SHADOW_SPACE, self.SHADOW_SPACE,
            -self.SHADOW_SPACE, -self.SHADOW_SPACE,
        )
        pv = self.NORMAL_PAD_V + (self.SELECTED_PAD_V - self.NORMAL_PAD_V) * p
        ph = self.NORMAL_PAD_H + (self.SELECTED_PAD_H - self.NORMAL_PAD_H) * p
        card = full.adjusted(ph, pv, -ph, -pv)
        grow = self.EXPAND_EXTRA * p * 0.5
        card = card.adjusted(0, -grow, 0, grow)

        # ---- 阴影 ----
        if p > 0.01:
            shadow = QColor(0, 0, 0, int(35 * p))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(shadow)
            painter.drawRoundedRect(card.adjusted(0, 2, 0, 4), 10, 10)

        # ---- 卡片背景 ----
        if p > 0.01:
            bg = QColor(self._c['PRIMARY'])
            bg.setAlpha(int(20 * p))
            border = QColor(self._c['PRIMARY'])
            border.setAlpha(int(100 * p))
            painter.setPen(QPen(border, 1.5))
            painter.setBrush(bg)
        elif hovered:
            painter.setPen(QPen(QColor(self._c['BORDER']), 1))
            painter.setBrush(QColor(self._c['NAV_ACTIVE']))
        else:
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(self._c['WHITE']))

        painter.drawRoundedRect(card, 10, 10)

        # ---- 左侧彩色竖条 ----
        if p > 0.01:
            bar = QColor(self._c['PRIMARY'])
            bar.setAlpha(int(255 * min(p * 1.5, 1.0)))
            painter.setBrush(bar)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(
                QRectF(card.left() + 2, card.top() + 6,
                       3, card.height() - 12),
                1.5, 1.5,
            )

        # ---- 文本（抗锯齿 + 自动省略） ----
        text = index.data(Qt.ItemDataRole.DisplayRole) or ""
        text_rect = card.adjusted(16 + 4 * p, 0, -14, 0)

        font = QFont("Microsoft YaHei", 13)
        if p > 0.5:
            font.setBold(True)
        painter.setFont(font)

        if p > 0.01:
            tc = QColor(self._c['PRIMARY'])
            tc.setAlpha(int(255 * max(p, 0.3)))
            painter.setPen(tc)
        else:
            painter.setPen(QColor(self._c['TEXT']))

        fm = QFontMetrics(font)
        elided = fm.elidedText(
            text, Qt.TextElideMode.ElideRight, int(text_rect.width())
        )
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignVCenter | Qt.TextFlag.TextSingleLine,
            elided,
        )

        painter.restore()
