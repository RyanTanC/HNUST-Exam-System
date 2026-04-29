"""设置对话框：即时反馈、深色模式、字体缩放."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSlider,
    QFrame,
    QWidget,
    QMessageBox,
)

from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.theme import Theme
from hnust_exam.utils.ui_helpers import themed_question, themed_info
from hnust_exam.views.widgets.toggle_switch import ToggleSwitch


class SettingsDialog(QDialog):
    """个性化设置对话框."""

    def __init__(self, config_mgr: ConfigManager, parent=None) -> None:
        super().__init__(parent)
        self.config_mgr = config_mgr
        self._orig_dark = Theme._is_dark
        self._orig_scale = Theme._font_scale
        self._marks_cleared = False

        cfg = config_mgr.load_config()
        self._show_immediately = cfg.get("show_answer_immediately", False)
        self._student_name = cfg.get("student_name", "")
        self._student_id = cfg.get("student_id", "")

        self.setWindowTitle("个性化设置")
        self.setMinimumWidth(500)
        self._build_ui()

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 顶部
        header = QFrame()
        header.setStyleSheet(f"background-color: {c['PRIMARY']}; padding: 6px 20px;")
        h_layout = QHBoxLayout(header)
        title = QLabel("⚙ 个性化设置")
        title.setStyleSheet("color: white; font-size: 16pt; font-weight: bold;")
        h_layout.addWidget(title)
        layout.addWidget(header)

        # 卡片容器
        body = QFrame()
        body.setStyleSheet(f"background-color: {c['BG']};")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(20, 16, 20, 12)
        body_layout.setSpacing(10)

        card_style = (
            f"background-color: {c['WHITE']}; "
            f"border: none; border-radius: 6px;"
        )

        # ── 卡片0：个人信息 ──
        info_card = QFrame()
        info_card.setStyleSheet(card_style)
        ic = QVBoxLayout(info_card)
        ic.setContentsMargins(16, 12, 16, 12)
        ic.setSpacing(8)
        ic_title = QLabel("个人信息")
        ic_title.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold;")
        ic.addWidget(ic_title)
        ic_desc = QLabel("设置姓名和学号，将在考试页面顶部显示")
        ic_desc.setStyleSheet(f"color: {c['MUTED']}; font-size: 9pt;")
        ic.addWidget(ic_desc)

        input_style = (
            f"QLineEdit {{ background-color: {c['SURFACE']}; color: {c['TEXT']}; "
            f"border: 1px solid {c['BORDER']}; border-radius: 4px; "
            f"padding: 6px 10px; font-size: 10pt; }}"
            f"QLineEdit:focus {{ border: 1px solid {c['PRIMARY']}; }}"
        )
        name_row = QHBoxLayout()
        name_label = QLabel("姓名：")
        name_label.setStyleSheet(f"color: {c['TEXT']}; font-size: 10pt; min-width: 50px;")
        self._name_input = QLineEdit(self._student_name)
        self._name_input.setPlaceholderText("请输入姓名")
        self._name_input.setStyleSheet(input_style)
        name_row.addWidget(name_label)
        name_row.addWidget(self._name_input)
        ic.addLayout(name_row)

        id_row = QHBoxLayout()
        id_label = QLabel("学号：")
        id_label.setStyleSheet(f"color: {c['TEXT']}; font-size: 10pt; min-width: 50px;")
        self._id_input = QLineEdit(self._student_id)
        self._id_input.setPlaceholderText("请输入学号")
        self._id_input.setStyleSheet(input_style)
        id_row.addWidget(id_label)
        id_row.addWidget(self._id_input)
        ic.addLayout(id_row)

        body_layout.addWidget(info_card)

        # ── 卡片1：即时反馈 ──
        feedback_card = QFrame()
        feedback_card.setStyleSheet(card_style)
        fc = QHBoxLayout(feedback_card)
        fc.setContentsMargins(16, 12, 16, 12)
        left = QVBoxLayout()
        fl = QLabel("答题后立即显示对错")
        fl.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold;")
        left.addWidget(fl)
        fd = QLabel("开启后选择答案即时显示对错，关闭后交卷统一评判")
        fd.setStyleSheet(f"color: {c['MUTED']}; font-size: 9pt;")
        left.addWidget(fd)
        fc.addLayout(left)
        fc.addStretch()
        self.feedback_check = ToggleSwitch()
        self.feedback_check.setChecked(self._show_immediately)
        fc.addWidget(self.feedback_check)
        self.feedback_check.toggled.connect(lambda: self._update_hint() if hasattr(self, '_status_hint') else None)
        body_layout.addWidget(feedback_card)

        # ── 卡片2：深色模式 ──
        dark_card = QFrame()
        dark_card.setStyleSheet(card_style)
        dc = QHBoxLayout(dark_card)
        dc.setContentsMargins(16, 12, 16, 12)
        left2 = QVBoxLayout()
        dl = QLabel("深色模式")
        dl.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold;")
        left2.addWidget(dl)
        dd = QLabel("切换深色/浅色主题，减少视觉疲劳")
        dd.setStyleSheet(f"color: {c['MUTED']}; font-size: 9pt;")
        left2.addWidget(dd)


        tip_label = QLabel("建议关闭应用重新启动，有恶性bug，主包不会修太难了")
        tip_label.setStyleSheet(f"color: {c['MUTED']}; font-size: 9pt;")
        left2.addWidget(tip_label)

        dc.addLayout(left2)
        dc.addStretch()
        self.dark_check = ToggleSwitch()
        self.dark_check.setChecked(Theme._is_dark)
        dc.addWidget(self.dark_check)
        self.dark_check.toggled.connect(lambda: self._update_hint() if hasattr(self, '_status_hint') else None)
        body_layout.addWidget(dark_card)

        # ── 卡片3：字体大小 ──
        font_card = QFrame()
        font_card.setStyleSheet(card_style)
        font_card_l = QVBoxLayout(font_card)
        font_card_l.setContentsMargins(16, 12, 16, 12)
        font_card_l.setSpacing(6)

        # 标题行
        fl_title = QLabel("字体大小")
        fl_title.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold;")
        font_card_l.addWidget(fl_title)
        fl_desc = QLabel("调整界面文字大小（80% ~ 150%）")
        fl_desc.setStyleSheet(f"color: {c['MUTED']}; font-size: 9pt;")
        font_card_l.addWidget(fl_desc)

        # 滑块 + 百分比
        slider_row = QHBoxLayout()
        slider_row.setSpacing(10)
        self._scale_slider = QSlider(Qt.Orientation.Horizontal)
        self._scale_slider.setRange(80, 150)
        self._scale_slider.setValue(int(Theme._font_scale * 100))
        self._scale_slider.setStyleSheet(
            f"QSlider::groove:horizontal {{ background: {c['BORDER']}; height: 6px; border-radius: 3px; }}"
            f"QSlider::handle:horizontal {{ background: {c['PRIMARY']}; width: 16px; height: 16px; "
            f"margin: -5px 0; border-radius: 8px; }}"
        )
        self._scale_label = QLabel(f"{int(Theme._font_scale * 100)}%")
        self._scale_label.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold; min-width: 40px;"
        )
        self._scale_slider.valueChanged.connect(self._on_scale_changed)
        slider_row.addWidget(self._scale_slider)
        slider_row.addWidget(self._scale_label)
        font_card_l.addLayout(slider_row)

        # 预览区
        preview_frame = QFrame()
        preview_frame.setStyleSheet(
            f"background-color: {c['SURFACE']}; border-radius: 4px;"
        )
        preview_layout = QVBoxLayout(preview_frame)
        preview_layout.setContentsMargins(10, 8, 10, 8)
        preview_layout.setSpacing(2)
        self._preview_title = QLabel("这是标题文字的预览效果")
        self._preview_title.setStyleSheet(
            f"color: {c['TEXT']}; font-size: {max(9, int(12 * Theme._font_scale))}pt; "
            f"font-weight: bold;"
        )
        preview_layout.addWidget(self._preview_title)
        self._preview_body = QLabel("这是正文内容的预览效果，用于确认字体大小是否合适")
        self._preview_body.setStyleSheet(
            f"color: {c['TEXT']}; font-size: {max(8, int(11 * Theme._font_scale))}pt;"
        )
        preview_layout.addWidget(self._preview_body)
        self._preview_small = QLabel("这是小字标注的预览效果")
        self._preview_small.setStyleSheet(
            f"color: {c['MUTED']}; font-size: {max(7, int(9 * Theme._font_scale))}pt;"
        )
        preview_layout.addWidget(self._preview_small)
        font_card_l.addWidget(preview_frame)
        body_layout.addWidget(font_card)

        # ── 卡片4：试卷标记管理 ──
        mark_card = QFrame()
        mark_card.setStyleSheet(card_style)
        mark_layout = QHBoxLayout(mark_card)
        mark_layout.setContentsMargins(16, 12, 16, 12)
        mark_text = QVBoxLayout()
        mark_title = QLabel("试卷标记管理")
        mark_title.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold;")
        mark_text.addWidget(mark_title)
        mark_desc = QLabel("查看标记含义说明，或清除所有试卷的完成标记")
        mark_desc.setStyleSheet(f"color: {c['MUTED']}; font-size: 9pt;")
        mark_text.addWidget(mark_desc)
        mark_layout.addLayout(mark_text)
        mark_layout.addStretch()

        mark_btn_layout = QVBoxLayout()
        mark_btn_layout.setSpacing(6)
        legend_btn = QPushButton("标记说明")
        legend_btn.setStyleSheet(
            f"background-color: {c['SURFACE']}; color: {c['TEXT']}; font-size: 9pt; "
            f"padding: 5px 14px; border: none; border-radius: 4px;"
        )
        legend_btn.clicked.connect(self._show_mark_legend)
        mark_btn_layout.addWidget(legend_btn)

        clear_btn = QPushButton("清除所有标记")
        clear_btn.setStyleSheet(
            f"background-color: {c['SURFACE']}; color: {c['DANGER']}; font-size: 9pt; "
            f"padding: 5px 14px; border: none; border-radius: 4px;"
        )
        clear_btn.clicked.connect(self._clear_all_marks)
        mark_btn_layout.addWidget(clear_btn)
        mark_layout.addLayout(mark_btn_layout)

        body_layout.addWidget(mark_card)

        # 状态提示
        self._status_hint = QLabel("")
        self._status_hint.setStyleSheet(
            f"color: {c['PRIMARY']}; font-size: 9pt; padding: 4px 0;"
        )
        body_layout.addWidget(self._status_hint)
        self._update_hint()

        layout.addWidget(body, 1)

        # 底部按钮栏
        btn_bar = QFrame()
        btn_bar.setStyleSheet(f"background-color: {c['WHITE']}; border-top: 1px solid {c['BORDER']};")
        btn_layout = QHBoxLayout(btn_bar)
        btn_layout.setContentsMargins(20, 10, 20, 14)
        btn_layout.addStretch()

        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet(
            f"background-color: {c['SURFACE']}; color: {c['MUTED']}; "
            f"padding: 7px 20px; border: none; border-radius: 4px;"
        )
        cancel_btn.clicked.connect(self._on_cancel)
        btn_layout.addWidget(cancel_btn)
        btn_layout.addSpacing(8)

        done_btn = QPushButton("完 成")
        done_btn.setStyleSheet(
            f"background-color: {c['PRIMARY']}; color: white; font-weight: bold; "
            f"padding: 7px 36px; border: none; border-radius: 4px;"
        )
        done_btn.clicked.connect(self._on_done)
        btn_layout.addWidget(done_btn)

        layout.addWidget(btn_bar)

    def _on_scale_changed(self, value: int) -> None:
        self._scale_label.setText(f"{value}%")
        Theme._font_scale = value / 100.0
        Theme.update_fonts()
        c = Theme.get_current_colors()
        self._preview_title.setStyleSheet(
            f"color: {c['TEXT']}; font-size: {max(9, int(12 * Theme._font_scale))}pt; "
            f"font-weight: bold; padding: 2px 0;"
        )
        self._preview_body.setStyleSheet(
            f"color: {c['TEXT']}; font-size: {max(8, int(11 * Theme._font_scale))}pt; padding: 2px 0;"
        )
        self._preview_small.setStyleSheet(
            f"color: {c['MUTED']}; font-size: {max(7, int(9 * Theme._font_scale))}pt; padding: 2px 0;"
        )
        self._update_hint()

    def _on_done(self) -> None:
        self._show_immediately = self.feedback_check.isChecked()
        Theme.set_dark_mode(self.dark_check.isChecked())

        self.config_mgr.save_config({
            "font_scale": Theme._font_scale,
            "dark_mode": Theme._is_dark,
            "show_answer_immediately": self._show_immediately,
            "user_python_path": self.config_mgr.load_config().get("user_python_path", ""),
            "student_name": self._name_input.text().strip(),
            "student_id": self._id_input.text().strip(),
        })
        self.accept()

    def _on_cancel(self) -> None:
        # 恢复原始设置
        Theme._font_scale = self._orig_scale
        Theme.update_fonts()
        Theme.set_dark_mode(self._orig_dark)
        self.reject()

    def _update_hint(self) -> None:
        """更新状态提示."""
        mode = "即时反馈" if self.feedback_check.isChecked() else "考试模式"
        theme = "深色" if self.dark_check.isChecked() else "浅色"
        scale = int(Theme._font_scale * 100)
        is_active = self.feedback_check.isChecked()
        # 使用预览状态的颜色（深色模式切换尚未生效）
        preview_dark = self.dark_check.isChecked()
        colors = Theme._DARK if preview_dark else Theme._LIGHT
        color = colors["SUCCESS"] if is_active else colors["PRIMARY"]
        self._status_hint.setText(f"答题：{mode}  |  主题：{theme}  |  字体：{scale}%")
        self._status_hint.setStyleSheet(f"color: {color}; font-size: 9pt; padding: 5px 0;")

    def _show_mark_legend(self) -> None:
        """显示标记说明."""
        c = Theme.get_current_colors()
        themed_info(
            self, "标记说明",
            "✓ 已完成 — 已交卷的试卷，后面显示最高得分\n\n"
            "○ 进行中 — 已开始但尚未交卷的试卷\n\n"
            "（无标记）— 尚未打开过的试卷"
        )

    def _clear_all_marks(self) -> None:
        """清除所有标记."""
        progress = self.config_mgr.load_progress()
        if not progress:
            themed_info(self, "提示", "当前没有任何标记记录")
            return
        reply = themed_question(
            self, "确认清除",
            "确定要清除所有试卷的完成标记吗？\n\n"
            "清除后，所有试卷的完成状态和得分记录将被重置。\n"
            "此操作不可撤销。",
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.config_mgr.save_progress({})
            self._marks_cleared = True
            themed_info(self, "完成", "所有标记已清除")
