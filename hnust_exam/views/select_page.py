"""选卷页：试卷列表、进度标记、设置按钮（列表版，效果对齐原版）"""

from __future__ import annotations

import os
import time
from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QMouseEvent
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFrame,
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
)

from hnust_exam.models.exam import Exam
from hnust_exam.utils.theme import Theme
from hnust_exam.utils.ui_helpers import themed_warning, themed_critical
from hnust_exam.views.animated_card_delegate import AnimatedCardDelegate

if TYPE_CHECKING:
    from hnust_exam.views.main_window import MainWindow


class _ToggleListWidget(QListWidget):
    """支持"再次点击已选中项取消选择" + 惯性滚动的列表控件."""

    def __init__(self, parent=None):
        super().__init__(parent)

        self._deselect_timer = QTimer(self)
        self._deselect_timer.setSingleShot(True)
        self._deselect_timer.setInterval(250)
        self._deselect_timer.timeout.connect(self._do_deselect)

        # 惯性滚动相关
        self._target_value = None
        self._inertia_start = None    # 起始滚动位置
        self._inertia_t0 = None       # 起始时间戳
        self._inertia_duration = 0.5  # 滚动动画时长（秒）
        self._inertia_timer = QTimer(self)
        self._inertia_timer.setInterval(4)  # ~250fps，实际由显示器刷新率限制
        self._inertia_timer.timeout.connect(self._inertia_tick)

        # QScroller 配置仅用于参数参考，不 grabGesture 以免拦截双击事件

    # ─── 滚轮惯性 ───
    @staticmethod
    def _ease_out_expo(t: float) -> float:
        """指数缓出：1 - 2^(-10t)，前快后慢，手感自然"""
        if t >= 1.0:
            return 1.0
        return 1.0 - pow(2, -10 * t)

    def wheelEvent(self, e):
        delta = e.angleDelta().y()
        if delta == 0:
            return

        sb = self.verticalScrollBar()
        pixel_delta = delta  # 1:1 像素映射

        current = sb.value()
        new_target = current - pixel_delta
        new_target = max(sb.minimum(), min(sb.maximum(), new_target))

        # 叠加：连续快速滚轮时，目标累加而不是重置
        if self._inertia_timer.isActive() and self._target_value is not None:
            self._target_value = max(sb.minimum(), min(sb.maximum(),
                                     self._target_value - pixel_delta))
        else:
            self._target_value = new_target

        # 每次滚轮重置动画起点，实现滚动"跟手"
        self._inertia_start = current
        self._inertia_t0 = time.time()

        if not self._inertia_timer.isActive():
            self._inertia_timer.start()

        e.accept()

    def _inertia_tick(self):
        sb = self.verticalScrollBar()
        if self._target_value is None or self._inertia_t0 is None:
            self._inertia_timer.stop()
            return

        elapsed = time.time() - self._inertia_t0
        t = min(elapsed / self._inertia_duration, 1.0)
        progress = self._ease_out_expo(t)

        start = self._inertia_start
        target = self._target_value
        current = int(start + (target - start) * progress)
        sb.setValue(current)

        if t >= 1.0:
            sb.setValue(self._target_value)
            self._target_value = None
            self._inertia_start = None
            self._inertia_t0 = None
            self._inertia_timer.stop()

    # ─── 点击取消选择 ───
    def mousePressEvent(self, e: QMouseEvent):
        item = self.itemAt(e.pos())
        if item and item.isSelected():
            self._deselect_timer.start()
        else:
            self._deselect_timer.stop()
        super().mousePressEvent(e)

    def _do_deselect(self):
        self.clearSelection()


class SelectPage(QWidget):
    """选择试卷页面."""

    def __init__(self, main_window: MainWindow, parent=None) -> None:
        super().__init__(parent)
        self.main_window = main_window
        self.exam_files: list[str] = []
        self._selected_file: str = ""
        self._list_widget: _ToggleListWidget | None = None
        self._build_ui()

    def showEvent(self, event) -> None:
        """每次显示时刷新试卷列表和进度."""
        super().showEvent(event)
        self._refresh_exam_list()

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ---- 顶部标题栏 ----
        header = QFrame()
        header.setStyleSheet(
            f"background-color: {c['PRIMARY']}; padding: 10px 20px;"
        )
        header_layout = QHBoxLayout(header)
        title = QLabel("HNUST仿真平台")
        title.setStyleSheet("color: white; font-size: 16pt; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()

        settings_btn = QPushButton("⚙ 设置")
        settings_btn.setStyleSheet(
            f"background-color: transparent; color: white; border: none; "
            f"font-size: 11pt; padding: 5px 10px;"
        )
        settings_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        settings_btn.clicked.connect(self._on_settings)
        header_layout.addWidget(settings_btn)

        root_layout.addWidget(header)

        # ---- 主内容区域 ----
        main_frame = QFrame()
        main_frame.setStyleSheet(f"background-color: {c['BG']};")
        main_layout = QVBoxLayout(main_frame)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        select_label = QLabel("请选择试卷")
        select_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        select_label.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 24pt; font-weight: bold; "
            f"padding: 40px 0 20px 0;"
        )
        main_layout.addWidget(select_label)

        # ========== 试卷列表 ==========
        self._list_widget = _ToggleListWidget()
        self._list_widget.itemDoubleClicked.connect(lambda _: self._on_start())
        self._list_widget.setSpacing(0)
        self._list_widget.setMinimumWidth(500)
        self._list_widget.setMaximumWidth(640)
        self._list_widget.setMinimumHeight(600)
        self._list_widget.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self._list_widget.setVerticalScrollMode(
            QAbstractItemView.ScrollMode.ScrollPerPixel
        )
        self._list_widget.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self._list_widget.setContentsMargins(8, 8, 8, 8)

        # 挂载自定义代理
        self._card_delegate = AnimatedCardDelegate(c, self._list_widget)
        self._list_widget.setItemDelegate(self._card_delegate)

        # 信号连接
        self._list_widget.selectionModel().currentChanged.connect(
            self._on_paper_selection_changed
        )
        self._list_widget.currentItemChanged.connect(self._on_current_changed)

        # 样式表只保留容器 + 滚动条
        self._list_widget.setStyleSheet(
            f"QListWidget {{"
            f"  background-color: {c['SURFACE']};"
            f"  border: 1px solid {c['BORDER']};"
            f"  border-radius: 8px;"
            f"  outline: none;"
            f"  padding: 3px;"
            f"}}"
            f"QScrollBar:vertical {{"
            f"  width: 6px; background: transparent;"
            f"  margin: 8px 2px 8px 0; border: none;"
            f"}}"
            f"QScrollBar::handle:vertical {{"
            f"  background: {c['BORDER']}; border-radius: 3px; min-height: 30px;"
            f"}}"
            f"QScrollBar::handle:vertical:hover {{"
            f"  background: {c['TEXT']}66;"
            f"}}"
            f"QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{"
            f"  height: 0; background: none;"
            f"}}"
            f"QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{"
            f"  background: none;"
            f"}}"
        )

        self._list_widget.installEventFilter(self)

        main_layout.addWidget(self._list_widget, 1, Qt.AlignmentFlag.AlignCenter)

        # ---- 开始按钮 ----
        self.start_btn = QPushButton("开始考试")
        self.start_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.start_btn.setStyleSheet(
            f"background-color: {c['PRIMARY']}; color: white; font-size: 16pt; "
            f"font-weight: bold; padding: 12px 40px; border: none; border-radius: 4px; "
            f"margin-top: 20px;"
        )
        self.start_btn.clicked.connect(self._on_start)
        main_layout.addWidget(self.start_btn, 0, Qt.AlignmentFlag.AlignCenter)

        # ---- 底部提示 ----
        hint = QLabel("该程序免费提供给HNUST学生使用，禁止任何形式的商用售卖")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setStyleSheet(
            f"color: {c['MUTED']}; font-size: 8pt; padding-top: 10px;"
        )
        main_layout.addWidget(hint)

        root_layout.addWidget(main_frame, 1)

    # ================================================================
    #  键盘事件过滤
    # ================================================================
    def eventFilter(self, obj, event):
        if obj is self._list_widget:
            if event.type() == event.Type.KeyPress:
                if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                    self._on_start()
                    return True
        return super().eventFilter(obj, event)

    # ================================================================
    #  列表项文本构造（显示进度标记）
    # ================================================================
    def _make_item_text(self, fname: str, progress: dict) -> str:
        """返回列表项显示的文字，如：PY程序填空  ✓ 85%"""
        display = os.path.splitext(fname)[0]
        entry = progress.get(fname, {})
        status = entry.get("status", "")
        if status == "completed":
            score = entry.get("best_score", "")
            display += f"    ✓ {score}%" if score else "    ✓"
        elif status == "started":
            display += "    ○"
        return display

    def _make_tooltip(self, fname: str, progress: dict) -> str:
        """构造多行 Tooltip"""
        entry = progress.get(fname, {})
        lines = []
        if entry.get("last_completed"):
            lines.append(f"上次完成：{entry['last_completed']}")
        if entry.get("best_score") is not None:
            lines.append(f"最高得分：{entry['best_score']}%")
        if entry.get("last_started"):
            lines.append(f"上次开始：{entry['last_started']}")
        return "\n".join(lines)

    # ================================================================
    #  刷新试卷列表
    # ================================================================
    def _refresh_exam_list(self) -> None:
        """从题库文件夹获取试卷文件并显示在列表中"""
        self._list_widget.blockSignals(True)
        self._list_widget.clear()
        self._selected_file = ""

        self.exam_files = Exam.list_exam_files()
        progress = self.main_window.config_mgr.load_progress()

        if not self.exam_files:
            self.start_btn.setEnabled(False)
            placeholder = QListWidgetItem("题库文件夹中没有找到试卷文件")
            placeholder.setFlags(Qt.ItemFlag.NoItemFlags)
            self._list_widget.addItem(placeholder)
            self._list_widget.blockSignals(False)
            return

        self.start_btn.setEnabled(True)

        for fname in self.exam_files:
            text = self._make_item_text(fname, progress)
            tooltip = self._make_tooltip(fname, progress)

            item = QListWidgetItem(text)
            if tooltip:
                item.setToolTip(tooltip)
            item.setData(Qt.ItemDataRole.UserRole, fname)
            item.setFont(QFont("Microsoft YaHei", 13))
            self._list_widget.addItem(item)

        self._list_widget.blockSignals(False)

    def _on_current_changed(self, current, previous):
        """更新当前选中的试卷文件名"""
        if current:
            self._selected_file = current.data(Qt.ItemDataRole.UserRole)
        else:
            self._selected_file = ""

    def _on_paper_selection_changed(self, current, previous):
        """选择变化时触发弹性动画"""
        cur_row = current.row()
        if cur_row < 0:
            return
        prev_row = previous.row() if previous.isValid() else -1
        self._card_delegate.animateTo(cur_row, prev_row)

    # ================================================================
    #  开始考试
    # ================================================================
    def _on_start(self) -> None:
        if not self._selected_file:
            themed_warning(self, "警告", "请先选择一份试卷")
            return

        exam_file = self._selected_file
        exam_name = os.path.splitext(exam_file)[0]

        exam = Exam()
        file_path = exam.find_exam_file(exam_name)
        if not file_path:
            themed_critical(self, "错误", "找不到试卷文件")
            return

        error = exam.load_from_excel(file_path)
        if error:
            themed_critical(self, "错误", error)
            return

        self.main_window.exam_page.setup_exam(exam, file_path)
        self.main_window.show_exam()

    def _on_settings(self) -> None:
        """打开设置对话框."""
        from hnust_exam.views.dialogs.settings_dialog import SettingsDialog
        dlg = SettingsDialog(self.main_window.config_mgr, self)
        dlg.exec()
        self.main_window._refresh_theme()