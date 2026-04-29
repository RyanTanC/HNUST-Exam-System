"""交卷确认对话框."""

from __future__ import annotations

from PySide6.QtCore import QEasingCurve, Qt, QTimer
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QFrame,
    QWidget,
    QApplication,
    QSizePolicy,
)

from hnust_exam.models.exam import Exam
from hnust_exam.utils.theme import Theme
from hnust_exam.utils.ui_helpers import themed_info


class SmoothScrollArea(QScrollArea):
    """鼠标滚轮触发平滑滚动，替代默认顿挫滚动。"""

    _STEPS = 15
    _INTERVAL = 16

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._target = 0
        self._start = 0
        self._progress = 0.0
        self._curve = QEasingCurve(QEasingCurve.Type.OutCubic)
        self._timer = QTimer(self)
        self._timer.setInterval(self._INTERVAL)
        self._timer.timeout.connect(self._step)

    def wheelEvent(self, event) -> None:
        self._timer.stop()
        bar = self.verticalScrollBar()
        self._start = bar.value()
        delta = -event.angleDelta().y() // 2
        self._target = max(bar.minimum(), min(bar.maximum(), bar.value() + delta))
        if self._start == self._target:
            return
        self._progress = 0.0
        self._timer.start()

    def _step(self) -> None:
        self._progress += 1.0 / self._STEPS
        if self._progress >= 1.0:
            self.verticalScrollBar().setValue(self._target)
            self._timer.stop()
            return
        t = self._curve.valueForProgress(self._progress)
        pos = self._start + round((self._target - self._start) * t)
        self.verticalScrollBar().setValue(pos)


class SubmitDialog(QDialog):
    def __init__(self, exam: Exam, parent=None) -> None:
        super().__init__(parent)
        self.exam = exam
        self.check_marked_index: int | None = None
        self.setWindowTitle("交卷确认")
        self.setMinimumSize(480, 420)
        self._build_ui()

    # ─────────────────────── 构建入口 ───────────────────────

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── 顶部标题栏 ──
        root.addWidget(self._make_header(c))

        # ── 中部内容区 ──
        # 【修复 1】先把 body 加入布局，再填充内容；
        #           即使后续局部异常也不会导致整个 body 丢失
        body = QWidget()
        body.setStyleSheet(f"background: {c['WHITE']};")
        body.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        body_layout = QVBoxLayout(body)
        # 【修复 2】用 layout margins 代替 stylesheet padding，更可靠
        body_layout.setContentsMargins(20, 16, 20, 16)
        body_layout.setSpacing(12)
        root.addWidget(body, 1)

        # 获取统计数据
        total = self.exam.total_count
        answered = self.exam.answered_count
        unanswered = self.exam.unanswered_count
        marked = self.exam.marked_count

        # 统计卡片
        body_layout.addWidget(self._make_stats_card(c, total, answered, unanswered, marked))

        # 标记提醒
        warn = self._make_marked_warning(c, total, marked)
        if warn is not None:
            body_layout.addWidget(warn)

        # 未答列表
        if unanswered > 0:
            body_layout.addWidget(self._make_unanswered_section(c, unanswered))

        # 底部提示
        hint = QLabel("注意：请勿修改程序中的其它任何内容。")
        hint.setStyleSheet(f"color: {c['MUTED']}; font-size: 8pt; border: none;")
        body_layout.addWidget(hint)

        # ── 底部按钮栏 ──
        root.addWidget(self._make_footer(c))

    # ─────────────────────── 各区块工厂方法 ───────────────────────

    def _make_header(self, c: dict) -> QFrame:
        header = QFrame()
        header.setStyleSheet(f"background: {c['PRIMARY']};")
        lay = QHBoxLayout(header)
        lay.setContentsMargins(20, 14, 20, 14)

        title = QLabel("交卷前检查")
        title.setStyleSheet("color: #fff; font-size: 16pt; font-weight: bold; border: none;")
        lay.addWidget(title)
        return header

    def _make_stats_card(
        self, c: dict, total: int, answered: int, unanswered: int, marked: int
    ) -> QFrame:
        """统计信息卡片——用 addLayout 而非嵌套 QFrame，避免 sizeHint 归零。"""
        card = QFrame()
        card.setStyleSheet(
            f"background: {c['SURFACE']}; "
            f"border: 1px solid {c['BORDER']}; border-radius: 6px;"
        )
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 12, 16, 12)
        card_layout.setSpacing(6)

        rows = [
            ("总题数", total,    c["TEXT"]),
            ("已作答", answered, c["SUCCESS"]),
            ("未作答", unanswered,
             c["DANGER"] if unanswered > 0 else c["TEXT"]),
            ("标记数", marked,
             c["NAV_MARKED_FG"] if marked > 0 else c["TEXT"]),
        ]

        for label, value, color in rows:
            # 【修复 3】直接 addLayout，不包 QFrame，彻底规避
            #           "transparent + border:none 导致 QFrame 高度为 0" 的问题
            h = QHBoxLayout()
            h.setContentsMargins(0, 2, 0, 2)

            lbl = QLabel(label)
            lbl.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; border: none;")
            h.addWidget(lbl)

            val = QLabel(str(value))
            val.setStyleSheet(
                f"color: {color}; font-size: 13pt; font-weight: bold; border: none;"
            )
            val.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            h.addWidget(val, 1)

            card_layout.addLayout(h)

        return card

    def _make_marked_warning(self, c: dict, total: int, marked: int) -> QWidget | None:
        if marked <= 0:
            return None

        marked_unanswered = [
            idx for idx in self.exam.marked_indices
            if idx < total and self.exam.questions[idx].number not in self.exam.answer_map
        ]
        marked_answered = [
            idx for idx in self.exam.marked_indices
            if idx < total and self.exam.questions[idx].number in self.exam.answer_map
        ]

        parts: list[str] = []
        if marked_unanswered:
            parts.append(f"{len(marked_unanswered)} 道标记题尚未作答")
        if marked_answered:
            parts.append(f"{len(marked_answered)} 道标记题已作答待确认")
        if not parts:
            return None

        frame = QFrame()
        frame.setStyleSheet(
            f"background: {c['WARN_BG']}; "
            f"border: 1px solid {c['WARN_BORDER']}; border-radius: 4px;"
        )
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(12, 8, 12, 8)

        msg = QLabel(f"你有 {'，'.join(parts)}，是否仔细检查？")
        msg.setWordWrap(True)
        msg.setStyleSheet(f"color: {c['WARN_TEXT']}; font-size: 11pt; border: none;")
        lay.addWidget(msg)

        if marked_unanswered:
            btn = QPushButton("检查标记题")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet(
                f"background: {c['ACCENT']}; color: #fff; font-size: 10pt; "
                f"padding: 4px 12px; border: none; border-radius: 4px;"
            )
            btn.clicked.connect(lambda: self._check_marked(marked_unanswered[0]))
            lay.addWidget(btn)

        return frame

    def _make_unanswered_section(self, c: dict, unanswered: int) -> QWidget:
        container = QWidget()
        lay = QVBoxLayout(container)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        title = QLabel(f"以下 {unanswered} 题尚未作答：")
        title.setStyleSheet(
            f"color: {c['DANGER']}; font-size: 12pt; font-weight: bold; border: none;"
        )
        lay.addWidget(title)

        scroll = SmoothScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMaximumHeight(150)
        scroll.setStyleSheet(f"border: 1px solid {c['BORDER']}; border-radius: 4px;")

        inner = QWidget()
        inner.setStyleSheet(f"background: {c['WHITE']};")
        inner_lay = QVBoxLayout(inner)
        inner_lay.setContentsMargins(8, 4, 8, 4)
        inner_lay.setSpacing(2)

        for q in self.exam.questions:
            if q.number not in self.exam.answer_map:
                lbl = QLabel(f"  ✗ 第{q.number}题 · {q.q_type} · {q.score}分")
                lbl.setStyleSheet(f"color: {c['DANGER']}; font-size: 10pt; border: none;")
                inner_lay.addWidget(lbl)

        inner_lay.addStretch()
        scroll.setWidget(inner)
        lay.addWidget(scroll)

        btn = QPushButton("复制未答题号到剪贴板")
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet(
            f"background: {c['SURFACE']}; color: {c['TEXT']}; font-size: 9pt; "
            f"padding: 4px 10px; border: none; border-radius: 4px;"
        )
        btn.clicked.connect(self._copy_unanswered)
        lay.addWidget(btn)

        return container

    def _make_footer(self, c: dict) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet(f"background: {c['WHITE']};")
        lay = QHBoxLayout(frame)
        lay.setContentsMargins(20, 0, 20, 16)

        back = QPushButton("返回继续答题")
        back.setCursor(Qt.CursorShape.PointingHandCursor)
        back.setStyleSheet(
            f"background: {c['SURFACE']}; color: {c['TEXT']}; font-size: 12pt; "
            f"font-weight: bold; padding: 8px 20px; "
            f"border: 1px solid {c['BORDER']}; border-radius: 4px;"
        )
        back.clicked.connect(self.reject)
        lay.addWidget(back)

        lay.addStretch()

        submit = QPushButton("确认交卷")
        submit.setCursor(Qt.CursorShape.PointingHandCursor)
        submit.setStyleSheet(
            f"background: {c['DANGER']}; color: #fff; font-size: 12pt; "
            f"font-weight: bold; padding: 8px 20px; border: none; border-radius: 4px;"
        )
        submit.clicked.connect(self.accept)
        lay.addWidget(submit)

        return frame

    # ─────────────────────── 槽函数 ───────────────────────

    def _copy_unanswered(self) -> None:
        nums = ", ".join(
            q.number for q in self.exam.questions
            if q.number not in self.exam.answer_map
        )
        QApplication.clipboard().setText(nums)
        themed_info(self, "已复制", f"未答题号已复制到剪贴板：\n{nums}")

    def _check_marked(self, index: int) -> None:
        self.check_marked_index = index
        self.reject()