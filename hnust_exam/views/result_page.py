"""成绩展示页：总分、正确率、每题详情."""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, QBasicTimer, QTimerEvent
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QFrame,
    QSizePolicy,
)

from hnust_exam.models.result import Result
from hnust_exam.utils.theme import Theme

if TYPE_CHECKING:
    from hnust_exam.models.exam import Exam
    from hnust_exam.views.main_window import MainWindow


# ── 固定列宽（px），保证所有行严格对齐 ──
_C_SEQ = 36
_C_ICON = 28
_C_NUM = 72
_C_TYPE = 64
_C_SCORE = 56
_C_ANS = 190
_C_ACT = 56

# ── 样式片段 ──
_NB = "border:none;"
_LB = "border:none;background:transparent;"


class SmoothScrollArea(QScrollArea):
    """鼠标滚轮平滑滚动区域."""

    _DURATION_MS = 180
    _INTERVAL_MS = 6  # ~166 fps

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._timer = QBasicTimer()
        self._anim_start = 0.0
        self._anim_target = 0.0
        self._anim_progress = 0.0

    def wheelEvent(self, event) -> None:
        sb = self.verticalScrollBar()
        current = float(sb.value())

        # 动画进行中：从当前位置继续
        if self._timer.isActive():
            frac = min(1.0, self._anim_progress)
            eased = 1.0 - (1.0 - frac) ** 3  # outCubic
            current = self._anim_start + (self._anim_target - self._anim_start) * eased

        delta = -event.angleDelta().y()
        target = current + delta
        target = max(float(sb.minimum()), min(target, float(sb.maximum())))

        self._anim_start = current
        self._anim_target = target
        self._anim_progress = 0.0

        if self._timer.isActive():
            self._timer.stop()
        self._timer.start(self._INTERVAL_MS, self)

    def timerEvent(self, event: QTimerEvent) -> None:
        if event.timerId() != self._timer.timerId():
            return

        self._anim_progress += self._INTERVAL_MS / self._DURATION_MS

        if self._anim_progress >= 1.0:
            self._timer.stop()
            self.verticalScrollBar().setValue(int(self._anim_target))
            return

        frac = self._anim_progress
        eased = 1.0 - (1.0 - frac) ** 3  # outCubic
        value = self._anim_start + (self._anim_target - self._anim_start) * eased
        self.verticalScrollBar().setValue(int(value))


class ResultPage(QWidget):
    """成绩展示页面."""

    def __init__(self, main_window: MainWindow, parent=None) -> None:
        super().__init__(parent)
        self.main_window = main_window
        self._results: list[Result] = []
        self._exam: Exam | None = None
        self._build_ui()

    # ═══════════════════════ 构建 ═══════════════════════

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._header(c))

        # ── 中部可滚动区域 ──
        scroll = SmoothScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setStyleSheet(_LB)

        body = QWidget()
        # 【修复】_NB 在 background 前面，不会覆盖背景色
        body.setStyleSheet(f"{_NB}background:{c['BG']};")
        self._body = body  # 保存引用供主题刷新
        body_lay = QVBoxLayout(body)
        body_lay.setContentsMargins(40, 24, 40, 24)
        body_lay.setSpacing(20)

        # 分数卡片
        body_lay.addWidget(self._score_card(c))

        # 摘要统计条
        self._summary = QWidget()
        self._summary.setStyleSheet(_LB)
        self._summary_lay = QHBoxLayout(self._summary)
        self._summary_lay.setContentsMargins(0, 0, 0, 0)
        self._summary_lay.setSpacing(12)
        body_lay.addWidget(self._summary)

        # 详情表格
        body_lay.addWidget(self._table_frame(c), 1)

        body_lay.addStretch()
        scroll.setWidget(body)
        root.addWidget(scroll, 1)

        root.addWidget(self._footer(c))

    # ─────────────── 顶部标题栏 ───────────────

    def _header(self, c: dict) -> QFrame:
        self._header_frame = QFrame()
        self._header_frame.setStyleSheet(f"{_NB}background:{c['PRIMARY']};")
        lay = QHBoxLayout(self._header_frame)
        lay.setContentsMargins(24, 14, 24, 14)

        lbl = QLabel("考试结果")
        lbl.setStyleSheet(f"color:#fff;font-size:16pt;font-weight:bold;{_LB}")
        lay.addWidget(lbl)
        return self._header_frame

    # ─────────────── 分数卡片 ───────────────

    def _score_card(self, c: dict) -> QFrame:
        self._score_card_frame = QFrame()
        # 这个容器需要边框，不用 _NB
        self._score_card_frame.setStyleSheet(
            f"background:{c['SURFACE']};"
            f"border:1px solid {c['BORDER']};border-radius:10px;"
        )
        self._score_card_frame.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)

        lay = QVBoxLayout(self._score_card_frame)
        lay.setContentsMargins(32, 28, 32, 24)
        lay.setSpacing(8)

        self._score_label = QLabel("—")
        self._score_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._score_label.setStyleSheet(
            f"font-size:44pt;font-weight:bold;color:{c['TEXT']};{_LB}"
        )
        lay.addWidget(self._score_label)

        self._grade_label = QLabel("")
        self._grade_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._grade_label.setStyleSheet(
            f"font-size:13pt;font-weight:600;color:{c['MUTED']};{_LB}"
        )
        lay.addWidget(self._grade_label)

        # 进度条轨道
        self._prog_track = QFrame()
        self._prog_track.setStyleSheet(f"{_NB}background:{c['BG']};border-radius:4px;")
        self._prog_track.setFixedHeight(8)

        self._prog_lay = QHBoxLayout(self._prog_track)
        self._prog_lay.setContentsMargins(0, 0, 0, 0)
        self._prog_lay.setSpacing(0)
        # 初始占位，setup_results 时重建
        self._prog_lay.addStretch(100)

        lay.addWidget(self._prog_track)
        return self._score_card_frame

    # ─────────────── 详情表格 ───────────────

    def _table_frame(self, c: dict) -> QFrame:
        self._detail_frame = QFrame()
        self._detail_frame.setStyleSheet(
            f"background:{c['SURFACE']};"
            f"border:1px solid {c['BORDER']};border-radius:10px;"
        )
        self._detail_frame.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)

        lay = QVBoxLayout(self._detail_frame)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # 表头
        lay.addWidget(self._table_header(c))

        # 分隔线
        self._table_sep = QFrame()
        self._table_sep.setFixedHeight(1)
        self._table_sep.setStyleSheet(f"{_NB}background:{c['BORDER']};")
        lay.addWidget(self._table_sep)

        # 表体容器
        self._tbody = QWidget()
        self._tbody.setStyleSheet(_LB)
        self._tbody_lay = QVBoxLayout(self._tbody)
        self._tbody_lay.setContentsMargins(0, 0, 0, 0)
        self._tbody_lay.setSpacing(0)
        lay.addWidget(self._tbody, 1)

        return self._detail_frame

    def _table_header(self, c: dict) -> QFrame:
        f = QFrame()
        f.setFixedHeight(38)
        f.setStyleSheet(f"{_NB}background:{c['BG']};")
        lay = QHBoxLayout(f)
        lay.setContentsMargins(14, 0, 14, 0)
        lay.setSpacing(0)

        columns = [
            ("#",         _C_SEQ,   Qt.AlignmentFlag.AlignCenter),
            ("",          _C_ICON,  Qt.AlignmentFlag.AlignCenter),
            ("题号",      _C_NUM,   Qt.AlignmentFlag.AlignLeft),
            ("题型",      _C_TYPE,  Qt.AlignmentFlag.AlignCenter),
            ("分值",      _C_SCORE, Qt.AlignmentFlag.AlignCenter),
            ("你的答案",  _C_ANS,   Qt.AlignmentFlag.AlignLeft),
            ("正确答案",  _C_ANS,   Qt.AlignmentFlag.AlignLeft),
            ("",          _C_ACT,   Qt.AlignmentFlag.AlignCenter),
        ]
        for text, w, align in columns:
            lbl = QLabel(text)
            lbl.setFixedWidth(w)
            lbl.setAlignment(align)
            lbl.setStyleSheet(
                f"color:{c['MUTED']};font-size:9pt;font-weight:bold;{_LB}"
            )
            lay.addWidget(lbl)

        return f

    # ─────────────── 底部按钮栏 ───────────────

    def _footer(self, c: dict) -> QFrame:
        self._footer_frame = QFrame()
        self._footer_frame.setStyleSheet(f"{_NB}background:{c['BG']};")
        lay = QHBoxLayout(self._footer_frame)
        lay.setContentsMargins(40, 12, 40, 16)
        lay.addStretch()

        self._back_btn = QPushButton("返回选卷")
        self._back_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._back_btn.setStyleSheet(
            f"background:{c['PRIMARY']};color:#fff;font-size:11pt;"
            f"font-weight:bold;padding:10px 32px;border:none;border-radius:6px;"
        )
        self._back_btn.clicked.connect(self.main_window.show_select)
        lay.addWidget(self._back_btn)
        return self._footer_frame

    # ═══════════════════════ 主题刷新 ═══════════════════════

    def refresh_theme(self) -> None:
        """刷新主题颜色."""
        c = Theme.get_current_colors()
        self._header_frame.setStyleSheet(f"{_NB}background:{c['PRIMARY']};")
        self._score_card_frame.setStyleSheet(
            f"background:{c['SURFACE']};"
            f"border:1px solid {c['BORDER']};border-radius:10px;"
        )
        self._body.setStyleSheet(f"{_NB}background:{c['BG']};")
        self._detail_frame.setStyleSheet(
            f"background:{c['SURFACE']};"
            f"border:1px solid {c['BORDER']};border-radius:10px;"
        )
        self._table_sep.setStyleSheet(f"{_NB}background:{c['BORDER']};")
        self._footer_frame.setStyleSheet(f"{_NB}background:{c['BG']};")
        self._back_btn.setStyleSheet(
            f"background:{c['PRIMARY']};color:#fff;font-size:11pt;"
            f"font-weight:bold;padding:10px 32px;border:none;border-radius:6px;"
        )
        # 重新填充数据（重建分数卡、摘要、表格行）
        if self._results and self._exam:
            self.setup_results(self._results, self._exam)

    # ═══════════════════════ 数据填充 ═══════════════════════

    def setup_results(self, results: list[Result], exam: Exam) -> None:
        self._results = results
        self._exam = exam
        c = Theme.get_current_colors()

        total = sum(r.score for r in results) or 1
        earned = sum(r.score for r in results if r.is_correct)
        correct_n = sum(1 for r in results if r.is_correct)
        wrong_n = len(results) - correct_n
        pct = earned / total * 100

        # 等级判定
        if pct >= 90:
            gc, gt = c["SUCCESS"], "优秀"
        elif pct >= 60:
            gc, gt = c["PRIMARY"], "及格"
        else:
            gc, gt = c["DANGER"], "不及格"

        # 分数 + 等级
        self._score_label.setText(f"{earned} / {total}")
        self._score_label.setStyleSheet(
            f"font-size:44pt;font-weight:bold;color:{gc};{_LB}"
        )
        self._grade_label.setText(f"正确率 {pct:.1f}%  ·  {gt}")
        self._grade_label.setStyleSheet(
            f"font-size:13pt;font-weight:600;color:{gc};{_LB}"
        )

        # 进度条
        self._rebuild_progress(gc, pct)

        # 摘要统计
        self._rebuild_summary(c, correct_n, wrong_n)

        # 表格行
        self._clear_layout(self._tbody_lay)
        for i, r in enumerate(results, 1):
            self._tbody_lay.addWidget(self._make_row(c, i, r))
        self._tbody_lay.addStretch()

    # ─────────────── 进度条重建 ───────────────

    def _rebuild_progress(self, color: str, pct: float) -> None:
        # 清除旧内容
        while self._prog_lay.count():
            item = self._prog_lay.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        bar = QFrame()
        bar.setFixedHeight(8)
        bar.setStyleSheet(f"{_NB}background:{color};border-radius:4px;")

        fill = max(1, int(pct))
        empty = max(1, 100 - fill)
        self._prog_lay.addWidget(bar, fill)
        self._prog_lay.addStretch(empty)

    # ─────────────── 摘要统计条 ───────────────

    def _rebuild_summary(self, c: dict, correct: int, wrong: int) -> None:
        self._clear_layout(self._summary_lay)

        items = [
            ("✓ 正确", str(correct), c["SUCCESS"]),
            ("✗ 错误", str(wrong),   c["DANGER"]),
        ]
        for label, value, color in items:
            chip = QFrame()
            chip.setStyleSheet(
                f"background:{c['SURFACE']};"
                f"border:1px solid {c['BORDER']};border-radius:6px;"
            )
            chip.setFixedHeight(32)
            chip_lay = QHBoxLayout(chip)
            chip_lay.setContentsMargins(14, 0, 14, 0)
            chip_lay.setSpacing(8)

            l1 = QLabel(label)
            l1.setStyleSheet(f"color:{color};font-size:10pt;font-weight:600;{_LB}")
            chip_lay.addWidget(l1)

            l2 = QLabel(value)
            l2.setStyleSheet(f"color:{c['TEXT']};font-size:10pt;font-weight:bold;{_LB}")
            chip_lay.addWidget(l2)

            self._summary_lay.addWidget(chip)

        self._summary_lay.addStretch()

    # ─────────────── 单行数据 ───────────────

    def _make_row(self, c: dict, seq: int, r: Result) -> QFrame:
        row = QFrame()
        bg = c['WHITE'] if seq % 2 else c['SURFACE']
        row.setStyleSheet(f"{_NB}background:{bg};")
        row.setFixedHeight(44)

        lay = QHBoxLayout(row)
        lay.setContentsMargins(14, 0, 14, 0)
        lay.setSpacing(0)

        ok = r.is_correct
        icon = "✓" if ok else "✗"
        clr = c["SUCCESS"] if ok else c["DANGER"]

        # 【关键】每列 setFixedWidth，所有行列严格对齐
        data = [
            (str(seq),               _C_SEQ,   Qt.AlignmentFlag.AlignCenter, c['MUTED']),
            (icon,                   _C_ICON,  Qt.AlignmentFlag.AlignCenter, clr),
            (r.question_number,      _C_NUM,   Qt.AlignmentFlag.AlignLeft,   c['TEXT']),
            (r.q_type,               _C_TYPE,  Qt.AlignmentFlag.AlignCenter, c['TEXT']),
            (f"{r.score}分",         _C_SCORE, Qt.AlignmentFlag.AlignCenter, c['MUTED']),
            (r.user_answer or "—",   _C_ANS,   Qt.AlignmentFlag.AlignLeft,   c['MUTED']),
            (r.correct_answer,       _C_ANS,   Qt.AlignmentFlag.AlignLeft,   clr),
        ]

        for text, w, align, tc in data:
            lbl = QLabel(text)
            lbl.setFixedWidth(w)
            lbl.setAlignment(align)
            lbl.setStyleSheet(f"color:{tc};font-size:9.5pt;{_LB}")
            # 答案列加 tooltip，鼠标悬停可看完整内容
            if w >= _C_ANS:
                lbl.setToolTip(text)
            lay.addWidget(lbl)

        # 查看按钮
        btn = QPushButton("查看")
        btn.setFixedWidth(_C_ACT)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet(
            f"QPushButton {{background:{c['BG']};color:{c['PRIMARY']};"
            f"font-size:9pt;font-weight:600;padding:3px 0;"
            f"border:1px solid {c['BORDER']};border-radius:3px;}}"
            f"QPushButton:hover {{background:{c['PRIMARY']};color:#fff;}}"
        )
        btn.clicked.connect(lambda checked=False, ri=r: self._review_question(ri))
        lay.addWidget(btn)

        return row

    # ═══════════════════════ 工具方法 ═══════════════════════

    def _review_question(self, result: Result) -> None:
        try:
            from hnust_exam.views.dialogs.review_dialog import ReviewDialog
            dlg = ReviewDialog(result, self._exam, self)
            dlg.exec()
        except Exception as e:
            print(f"[ReviewDialog] 打开失败: {e}")
            import traceback
            traceback.print_exc()

    @staticmethod
    def _clear_layout(layout) -> None:
        while layout.count():
            item = layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
