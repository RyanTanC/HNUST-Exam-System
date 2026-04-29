"""右侧题目导航面板."""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, QPropertyAnimation, QEasingCurve, Property, QTimer
from PySide6.QtGui import QPainter, QPen, QColor
from PySide6.QtWidgets import (
    QScroller,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QFrame,
    QGraphicsOpacityEffect,
)

from hnust_exam.utils.theme import Theme

if TYPE_CHECKING:
    from hnust_exam.views.exam_page import ExamPage


class SmoothScrollArea(QScrollArea):
    """支持平滑滚动的 QScrollArea."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._anim = QPropertyAnimation(self.verticalScrollBar(), b"value")
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._anim.setDuration(120)

    def wheelEvent(self, event) -> None:
        sb = self.verticalScrollBar()
        step = 100
        delta = -step if event.angleDelta().y() > 0 else step
        self._anim.stop()
        self._anim.setStartValue(sb.value())
        self._anim.setEndValue(
            max(sb.minimum(), min(sb.maximum(), sb.value() + delta))
        )
        self._anim.start()


class ArrowIndicator(QWidget):
    """可旋转的折叠箭头指示器，支持平滑旋转动画.

    折叠状态：箭头朝右 (▶)，rotation = 0°
    展开状态：箭头朝下 (▼)，rotation = 90°
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._rotation = 0.0
        self._color = QColor("#555555")
        self.setFixedSize(16, 16)

        self._anim = QPropertyAnimation(self, b"rotation")
        self._anim.setDuration(180)
        self._anim.setEasingCurve(QEasingCurve.Type.OutBack)

    # ---- rotation 属性 ----
    def _get_rotation(self) -> float:
        return self._rotation

    def _set_rotation(self, val: float) -> None:
        self._rotation = val
        self.update()

    rotation = Property(float, fget=_get_rotation, fset=_set_rotation)

    def set_color(self, color: str) -> None:
        """设置箭头颜色."""
        self._color = QColor(color)
        self.update()

    def set_expanded(self, expanded: bool, animated: bool = True) -> None:
        """设置展开/折叠状态。expanded=True 时箭头旋转到朝下."""
        target = 90.0 if expanded else 0.0
        if animated:
            self._anim.stop()
            self._anim.setStartValue(self._rotation)
            self._anim.setEndValue(target)
            self._anim.start()
        else:
            self._rotation = target
            self.update()

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.translate(self.width() / 2, self.height() / 2)
        painter.rotate(self._rotation)

        pen = QPen(
            self._color, 2.0,
            Qt.PenStyle.SolidLine,
            Qt.PenCapStyle.RoundCap,
            Qt.PenJoinStyle.RoundJoin,
        )
        painter.setPen(pen)

        # 右向 V 形箭头，旋转 90° 后朝下
        painter.drawLine(-3, -4, 3, 0)
        painter.drawLine(3, 0, -3, 4)

        painter.end()


class NavPanel(QWidget):
    """右侧题目导航面板，支持折叠分组及动画效果.

    动画效果：
    - 分组展开/折叠时的平滑高度过渡（QPropertyAnimation on maximumHeight）
    - 箭头指示器的旋转动画（0° ↔ 90°，OutBack 弹性曲线）
    - 展开时题目按钮的交错渐显（QGraphicsOpacityEffect + stagger）
    - 面板首次构建时的分组入场渐显
    """

    def __init__(self, exam_page: ExamPage, parent=None) -> None:
        super().__init__(parent)
        self.exam_page = exam_page
        self._panels: dict[str, dict] = {}
        self._q_buttons: dict[int, QPushButton] = {}
        self._animating = False
        self._anim_gen = 0  # 用于使过期的交错回调失效
        self._running_anims: list[QPropertyAnimation] = []
        self.setFixedWidth(250)
        self._build_ui()

    # ── UI 构建 ──────────────────────────────────────────────

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 外框
        self._frame = QFrame()
        self._frame.setStyleSheet(
            f"background-color: {c['WHITE']}; "
            f"border: 1px solid {c['BORDER']}; border-radius: 4px;"
        )
        frame_layout = QVBoxLayout(self._frame)
        frame_layout.setContentsMargins(0, 0, 0, 0)
        frame_layout.setSpacing(0)

        # 标题
        self._header = QLabel("题目导航")
        self._header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._header.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 12pt; font-weight: bold; "
            f"padding: 12px 5px 6px 5px;"
        )
        frame_layout.addWidget(self._header)

        self._sep = QFrame()
        self._sep.setFixedHeight(1)
        self._sep.setStyleSheet(f"background-color: {c['BORDER']};")
        frame_layout.addWidget(self._sep)

        # 可滚动区域
        scroll_area = SmoothScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff,
        )
        QScroller.grabGesture(
            scroll_area.viewport(),
            QScroller.ScrollerGestureType.TouchGesture,
        )

        self._nav_content = QWidget()
        self._nav_content.setStyleSheet(f"background-color: {c['WHITE']};")
        self._nav_layout = QVBoxLayout(self._nav_content)
        self._nav_layout.setContentsMargins(4, 4, 4, 4)
        self._nav_layout.setSpacing(0)
        self._nav_layout.addStretch()

        scroll_area.setWidget(self._nav_content)
        frame_layout.addWidget(scroll_area, 1)

        # 底部状态
        self._status_label = QLabel("")
        self._status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._status_label.setStyleSheet(
            f"color: {c['MUTED']}; font-size: 9pt; padding: 6px 0 10px 0;"
        )
        frame_layout.addWidget(self._status_label)

        layout.addWidget(self._frame)

    def refresh_theme(self) -> None:
        """刷新主题颜色."""
        c = Theme.get_current_colors()
        self._frame.setStyleSheet(
            f"background-color: {c['WHITE']}; "
            f"border: 1px solid {c['BORDER']}; border-radius: 4px;"
        )
        self._header.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 12pt; font-weight: bold; "
            f"padding: 12px 5px 6px 5px;"
        )
        self._sep.setStyleSheet(f"background-color: {c['BORDER']};")
        self._nav_content.setStyleSheet(f"background-color: {c['WHITE']};")
        self._status_label.setStyleSheet(
            f"color: {c['MUTED']}; font-size: 9pt; padding: 6px 0 10px 0;"
        )
        self.refresh()

    # ── 重置 ─────────────────────────────────────────────────

    def reset(self) -> None:
        """重置面板状态，切换试卷时调用."""
        for anim in self._running_anims:
            anim.stop()
        self._running_anims.clear()
        self._animating = False
        self._anim_gen += 1

        self._panels.clear()
        self._q_buttons.clear()

        while self._nav_layout.count() > 1:
            item = self._nav_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    # ── 刷新 ─────────────────────────────────────────────────

    def refresh(self) -> None:
        """刷新导航面板状态."""
        exam = self.exam_page.exam
        if not exam:
            return

        c = Theme.get_current_colors()

        # 首次构建面板
        if not self._panels:
            self._build_panels(exam, c)

        # 更新按钮状态
        for q in exam.questions:
            idx = q.index
            if idx not in self._q_buttons:
                continue

            btn = self._q_buttons[idx]
            global_num = q.number

            # 计算类型内序号
            type_questions = exam.question_groups.get(q.q_type, [])
            type_idx = (
                type_questions.index(q) + 1 if q in type_questions else 0
            )

            if idx == exam.current_index:
                bg = c["NAV_CURRENT"]
                fg = "white"
                weight = "bold"
            elif global_num in exam.answer_map:
                bg = c["NAV_ANSWERED_BG"]
                fg = c["NAV_ANSWERED_FG"]
                weight = "normal"
            else:
                bg = c["WHITE"]
                fg = c["TEXT"]
                weight = "normal"

            btn_text = f"  第{type_idx}题（{global_num}）"
            if exam.is_marked(idx):
                btn_text = "🚩" + btn_text.strip()
                if idx != exam.current_index and global_num not in exam.answer_map:
                    bg = c["NAV_MARKED_BG"]
                    fg = c["NAV_MARKED_FG"]

            btn.setText(btn_text)
            btn.setStyleSheet(
                f"background-color: {bg}; color: {fg}; font-weight: {weight}; "
                f"font-size: 9pt; text-align: left; padding: 6px 20px; "
                f"border: none;"
            )

        # 更新标记按钮
        if exam.is_marked(exam.current_index):
            self.exam_page._buttons["标记试题"].setText("取消标记")
        else:
            self.exam_page._buttons["标记试题"].setText("标记试题")

        # 更新状态
        answered = exam.answered_count
        marked = exam.marked_count
        unanswered = exam.unanswered_count
        self._status_label.setText(
            f"未答 {unanswered}，已答 {answered}，标记 {marked}"
        )

        self.exam_page._update_progress()

    # ── 面板构建 ──────────────────────────────────────────────

    def _build_panels(self, exam, c: dict) -> None:
        """首次构建导航面板."""
        # 清除已有内容（保留 stretch）
        while self._nav_layout.count() > 1:
            item = self._nav_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        first_type = True
        for group_idx, q_type in enumerate(exam.active_type_order):
            questions = exam.question_groups.get(q_type, [])
            if not questions:
                continue

            if not first_type:
                sep = QFrame()
                sep.setFixedHeight(1)
                sep.setStyleSheet(
                    f"background-color: {c['BORDER']}; margin: 4px 0;"
                )
                self._nav_layout.insertWidget(
                    self._nav_layout.count() - 1, sep,
                )
            first_type = False

            # ── 分组标题（可点击折叠/展开）──
            header = QFrame()
            header.setCursor(Qt.CursorShape.PointingHandCursor)
            header.setStyleSheet(
                f"background-color: {c['NAV_HEADER_BG']}; padding: 4px; "
                f"border-radius: 2px; border: none;"
            )
            header_layout = QHBoxLayout(header)
            header_layout.setContentsMargins(4, 2, 4, 2)

            arrow = ArrowIndicator()
            arrow.set_color(c["TEXT"])
            header_layout.addWidget(arrow)

            title_label = QLabel(f"{q_type}（{len(questions)}题）")
            title_label.setStyleSheet(
                f"color: {c['TEXT']}; font-size: 10pt; font-weight: bold;"
                f"border: none;"
            )
            header_layout.addWidget(title_label, 1)

            self._nav_layout.insertWidget(
                self._nav_layout.count() - 1, header,
            )

            # ── 题目按钮区 ──
            body = QFrame()
            body.setStyleSheet(f"background-color: {c['WHITE']}; border: none;")
            body_layout = QVBoxLayout(body)
            body_layout.setContentsMargins(0, 0, 0, 0)
            body_layout.setSpacing(0)

            for type_idx, q in enumerate(questions):
                btn = QPushButton(f"  第{type_idx + 1}题（{q.number}）")
                btn.setCursor(Qt.CursorShape.PointingHandCursor)
                btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
                btn.setStyleSheet(
                    f"background-color: {c['WHITE']}; color: {c['TEXT']}; "
                    f"font-size: 9pt; text-align: left; padding: 5px 20px; "
                    f"border: none;"
                )
                btn.clicked.connect(
                    lambda checked, gi=q.index: self.exam_page.jump_to(gi),
                )
                body_layout.addWidget(btn)
                self._q_buttons[q.index] = btn

            self._nav_layout.insertWidget(
                self._nav_layout.count() - 1, body,
            )

            # ── 初始展开/折叠状态 ──
            is_open = group_idx == 0
            if is_open:
                arrow.set_expanded(True, animated=False)
            else:
                body.hide()
                body.setMaximumHeight(0)
                arrow.set_expanded(False, animated=False)

            # ── 点击标题切换 ──
            header.mousePressEvent = (
                lambda event, bd=body, ar=arrow: self._toggle_group(bd, ar)
            )

            self._panels[q_type] = {
                "header": header,
                "arrow": arrow,
                "body": body,
            }

        # 入场动画延迟到布局计算完成后
        QTimer.singleShot(0, self._animate_entry)

    # ── 入场动画 ─────────────────────────────────────────────

    def _animate_entry(self) -> None:
        """所有分组的交错入场渐显动画."""
        if not self._panels:
            return

        delay = 0
        for q_type in self._panels:
            panel = self._panels[q_type]
            header = panel["header"]
            body = panel["body"]

            self._fade_in_widget(header, delay)
            delay += 60

            if body.isVisible():
                self._fade_in_widget(body, delay)
                delay += 60

    def _fade_in_widget(self, widget: QWidget, delay_ms: int) -> None:
        """对单个控件执行延迟渐显动画."""
        effect = QGraphicsOpacityEffect(widget)
        effect.setOpacity(0.0)
        widget.setGraphicsEffect(effect)

        anim = QPropertyAnimation(effect, b"opacity")
        anim.setDuration(200)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.finished.connect(lambda w=widget, a=anim: (
            w.setGraphicsEffect(None), self._safe_remove_anim(a)
        ))

        QTimer.singleShot(delay_ms, anim.start)
        self._running_anims.append(anim)

    # ── 展开/折叠动画 ────────────────────────────────────────

    def _toggle_group(self, body: QFrame, arrow: ArrowIndicator) -> None:
        """带动画的展开/折叠切换."""
        if self._animating:
            return

        # 使过期的交错回调失效
        self._anim_gen += 1

        # 停止残留动画并清理图形效果
        for anim in self._running_anims:
            anim.stop()
        self._running_anims.clear()
        self._clear_all_graphics_effects()

        if body.isVisible():
            self._collapse_group(body, arrow)
        else:
            self._expand_group(body, arrow)

    def _expand_group(self, body: QFrame, arrow: ArrowIndicator) -> None:
        """展开动画：箭头旋转 + 高度从 0 增长到目标值 + 按钮交错渐显."""
        self._animating = True
        arrow.set_expanded(True)

        # 测量目标高度
        body.show()
        body.setMaximumHeight(16777215)
        target_h = body.sizeHint().height()
        if target_h <= 0:
            target_h = self._estimate_body_height(body)

        # 从 0 开始动画
        body.setMaximumHeight(0)

        anim = QPropertyAnimation(body, b"maximumHeight")
        anim.setDuration(240)
        anim.setStartValue(0)
        anim.setEndValue(target_h)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)

        def _on_done():
            body.setMaximumHeight(16777215)
            self._animating = False
            self._safe_remove_anim(anim)

        anim.finished.connect(_on_done)
        anim.start()
        self._running_anims.append(anim)

        # 按钮交错渐显
        self._stagger_fade_in_buttons(body)

    def _collapse_group(self, body: QFrame, arrow: ArrowIndicator) -> None:
        """折叠动画：高度从当前值缩减到 0 + 箭头旋转."""
        self._animating = True
        arrow.set_expanded(False)

        current_h = body.height()

        anim = QPropertyAnimation(body, b"maximumHeight")
        anim.setDuration(200)
        anim.setStartValue(current_h)
        anim.setEndValue(0)
        anim.setEasingCurve(QEasingCurve.Type.InOutCubic)

        def _on_done():
            body.hide()
            body.setMaximumHeight(16777215)
            self._animating = False
            self._safe_remove_anim(anim)

        anim.finished.connect(_on_done)
        anim.start()
        self._running_anims.append(anim)

    # ── 按钮交错渐显 ─────────────────────────────────────────

    _MAX_STAGGER = 10  # 超过此数量的按钮跳过交错动画

    def _stagger_fade_in_buttons(self, body: QFrame) -> None:
        """展开时对题目按钮做交错渐显动画."""
        buttons = self._nav_buttons_in(body)
        if not buttons:
            return

        if len(buttons) > self._MAX_STAGGER:
            return

        gen = self._anim_gen
        for i, btn in enumerate(buttons):
            effect = QGraphicsOpacityEffect(btn)
            effect.setOpacity(0.0)
            btn.setGraphicsEffect(effect)

            anim = QPropertyAnimation(effect, b"opacity")
            anim.setDuration(160)
            anim.setStartValue(0.0)
            anim.setEndValue(1.0)
            anim.setEasingCurve(QEasingCurve.Type.OutCubic)
            anim.finished.connect(lambda w=btn, a=anim: (
                w.setGraphicsEffect(None), self._safe_remove_anim(a)
            ))

            delay = 60 + i * 25

            def _start(a=anim, g=gen):
                if g == self._anim_gen:
                    a.start()
                    self._running_anims.append(a)

            QTimer.singleShot(delay, _start)

    # ── 工具方法 ─────────────────────────────────────────────

    def _clear_all_graphics_effects(self) -> None:
        """清除所有导航控件上残留的 QGraphicsEffect."""
        for btn in self._q_buttons.values():
            btn.setGraphicsEffect(None)
        for panel in self._panels.values():
            panel["header"].setGraphicsEffect(None)
            panel["body"].setGraphicsEffect(None)

    @staticmethod
    def _nav_buttons_in(body: QFrame) -> list[QPushButton]:
        """获取 body 容器内所有 QPushButton."""
        buttons: list[QPushButton] = []
        lay = body.layout()
        if lay:
            for i in range(lay.count()):
                w = lay.itemAt(i).widget()
                if isinstance(w, QPushButton):
                    buttons.append(w)
        return buttons

    @staticmethod
    def _estimate_body_height(body: QFrame) -> int:
        """手动估算 body 的期望高度（sizeHint 失效时的兜底）."""
        lay = body.layout()
        if not lay:
            return 200
        h = lay.contentsMargins().top() + lay.contentsMargins().bottom()
        for i in range(lay.count()):
            item = lay.itemAt(i)
            if item and item.widget():
                h += item.widget().sizeHint().height()
        h += lay.spacing() * max(0, lay.count() - 1)
        return max(h, 100)

    def _safe_remove_anim(self, anim: QPropertyAnimation) -> None:
        """安全地从运行列表中移除已完成的动画引用."""
        try:
            self._running_anims.remove(anim)
        except ValueError:
            pass
