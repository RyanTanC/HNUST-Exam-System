"""设置对话框：基于 qfluentwidgets 的 Fluent 风格设置页.

界面组件全部改用 qfluentwidgets（SettingCardGroup / SwitchSettingCard /
PushSettingCard / SmoothScrollArea 等），业务逻辑与原实现保持一致。
"""

from __future__ import annotations

import sys
from threading import Thread

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QMessageBox,
    QVBoxLayout,
    QWidget,
)

from qfluentwidgets import (
    BodyLabel,
    CaptionLabel,
    CardWidget,
    ComboBox,
    FluentIcon as FIF,
    IconWidget,
    LineEdit,
    PrimaryPushButton,
    ProgressBar,
    PushButton,
    PushSettingCard,
    SettingCard,
    SettingCardGroup,
    Slider,
    SmoothScrollArea,
    StrongBodyLabel,
    SubtitleLabel,
    SwitchSettingCard,
    Theme as FT,
    setTheme,
    setThemeColor,
)

from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.services.update_checker import fetch_update_info
from hnust_exam.utils import constants
from hnust_exam.utils.theme import Theme
from hnust_exam.utils.ui_helpers import (
    themed_critical,
    themed_info,
    themed_question,
    themed_warning,
)
from hnust_exam.views.dialogs.update_dialog import UpdateDialog


# app.py 那份全局 QSS 会样式化 QPushButton/QLineEdit 等基础类，覆盖掉
# qfluentwidgets 的自绘效果。这里在对话框范围内把这些规则清空，让 Fluent
# 组件按自己的方式绘制。
_QSS_NEUTRALIZE = """
QPushButton, QLineEdit, QComboBox, QProgressBar, QSlider {
    background: transparent;
    border: none;
    padding: 0;
}
"""

_STRICTNESS_VALUES = ["strict", "normal", "lenient"]


class _SectionCard(CardWidget):
    """自适应高度的自定义卡片。

    qfluentwidgets 的 ``ExpandLayout`` 用 ``widget.height()``（当前高度）而不是
    ``sizeHint()`` 累加分组高度，而 ``CardWidget`` 本身没有固定高度：首帧布局时
    它的高度还是默认值，整张卡片会被压扁成一条空白。

    标准 ``SettingCard`` 用 ``setFixedHeight`` 绕开了这个问题，这里照做——内容
    构建完（或可见性/字体变化后）调用 :meth:`sync_height` 同步一次固定高度。
    """

    def sync_height(self) -> None:
        layout = self.layout()
        if layout is not None:
            layout.activate()
        height = self.sizeHint().height()
        self.setFixedHeight(height)
        self.resize(self.width(), height)


class SettingsDialog(QDialog):
    """个性化设置对话框."""

    def __init__(self, config_mgr: ConfigManager, parent=None) -> None:
        super().__init__(parent)
        self.config_mgr = config_mgr
        self._orig_dark = Theme._is_dark
        self._orig_scale = Theme._font_scale
        self._marks_cleared = False
        self._checking_update = False
        self._checking_qb = False

        self._cfg = config_mgr.load_config()
        self._show_immediately = self._cfg.get("show_answer_immediately", False)
        self._auto_advance = self._cfg.get("auto_advance_on_choice", False)
        self._student_name = self._cfg.get("student_name", "")
        self._student_id = self._cfg.get("student_id", "")

        # qfluentwidgets 有独立主题，需与应用当前主题对齐
        self._sync_fluent_theme()

        self.setWindowTitle("个性化设置")
        self.setMinimumSize(520, 520)
        self.resize(580, 720)
        self.setSizeGripEnabled(True)
        self._build_ui()
        self._apply_local_style()

    @staticmethod
    def _sync_fluent_theme(dark: bool | None = None) -> None:
        """把 qfluentwidgets 的主题和主色对齐到应用主题.

        qfluentwidgets 自带一套主题，默认主色是青色 ``#009faa``，和 ``Theme`` 里的
        ``PRIMARY``（浅色 ``#0078d7`` / 深色 ``#4da6ff``）对不上。不对齐的话设置页
        的开关、滑块、进度条、主按钮会和软件其它部分两个色。

        注意：浅色下渲染出来和 ``Theme.PRIMARY`` 完全一致；深色下
        ``ThemeColor.PRIMARY.color()`` 会对主题色做一次 ``s *= 0.84, v = 1``
        的变换（库为了让深色背景上的强调色不那么刺眼，属于刻意设计），
        所以实际画出来会略淡一点（``#4da6ff`` → ``#69b4ff``）。色相一致，
        不要去反向补偿——那会把库的内部实现细节焊死在代码里。

        Parameters
        ----------
        dark : bool | None
            预览中的深色状态。传 ``None`` 表示按 ``Theme`` 的当前状态。
        """
        if dark is None:
            dark = Theme._is_dark
        colors = Theme._DARK if dark else Theme._LIGHT
        setTheme(FT.DARK if dark else FT.LIGHT)
        # save=False：不要往 qfluentwidgets 自己的配置文件里写东西
        # （它默认落在相对路径 config/config.json，也就是进程当前目录）
        setThemeColor(colors["PRIMARY"])

    def _apply_local_style(self) -> None:
        """下发对话框自己的样式表（底色 + 文字色 + 中性化基础控件）。

        app.py 的全局 QSS 只给 ``QMainWindow`` 上了底色，``QDialog`` 拿不到，
        深色模式下会变成浅底白字；而且那份 QSS 在深色预览期间是过期的。
        所以这里按「预览中的主题」自己算一遍颜色，让对话框在实时预览时也正确。
        """
        dark = self.dark_check.switchButton.isChecked()
        c = Theme._DARK if dark else Theme._LIGHT
        self.setStyleSheet(
            f'QDialog {{ background-color: {c["BG"]}; }}\n'
            f'QLabel {{ color: {c["TEXT"]}; background: transparent; border: none; padding: 0; }}\n'
            + _QSS_NEUTRALIZE
        )

    # ───────── 整体骨架 ─────────

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        layout.addWidget(self._make_header())

        scroll = SmoothScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.enableTransparentBackground()
        # 视口默认会铺一层系统调色板底色（浅色），深色模式下会在对话框里露白
        scroll.viewport().setAutoFillBackground(False)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(16, 16, 16, 16)
        body_layout.setSpacing(18)

        body_layout.addWidget(self._make_profile_group(body))
        body_layout.addWidget(self._make_answer_group(body))
        body_layout.addWidget(self._make_appearance_group(body))
        body_layout.addWidget(self._make_update_group(body))
        body_layout.addWidget(self._make_mark_group(body))

        self._status_hint = CaptionLabel("")
        body_layout.addWidget(self._status_hint)
        body_layout.addStretch(1)

        scroll.setWidget(body)
        # QScrollArea.setWidget() 会把内容控件设成自填背景，于是它铺的是系统浅色
        # 调色板，深色模式下会在对话框里露白。关掉，让对话框底色透上来。
        body.setAutoFillBackground(False)
        layout.addWidget(scroll, 1)
        layout.addWidget(self._make_footer())

        self._update_hint()

    def _make_header(self) -> QWidget:
        header = QWidget()
        h = QHBoxLayout(header)
        h.setContentsMargins(24, 18, 24, 12)
        h.setSpacing(10)

        col = QVBoxLayout()
        col.setSpacing(2)
        col.addWidget(SubtitleLabel("个性化设置"))
        col.addWidget(CaptionLabel("自定义你的使用体验"))
        h.addLayout(col)
        h.addStretch(1)
        return header

    def _make_footer(self) -> QWidget:
        footer = QWidget()
        f = QHBoxLayout(footer)
        f.setContentsMargins(24, 12, 24, 16)
        f.setSpacing(10)
        f.addStretch(1)

        cancel_btn = PushButton("取消")
        cancel_btn.clicked.connect(self._on_cancel)
        f.addWidget(cancel_btn)

        done_btn = PrimaryPushButton("完成")
        done_btn.clicked.connect(self._on_done)
        f.addWidget(done_btn)
        return footer

    # ───────── 通用小工具 ─────────

    @staticmethod
    def _make_section(
        icon, title: str, desc: str, parent: QWidget
    ) -> tuple[_SectionCard, QVBoxLayout]:
        """构建一张带「图标 + 标题 + 说明」头部的自定义卡片，返回 (卡片, 内容布局).

        头部几何刻意和 ``SettingCard`` 保持一致（左内边距 16 + 图标 16 + 间距 16），
        这样标题会落在同一个 x 上，整页左对齐。

        注意：填充完内容后必须调用 ``card.sync_height()``，否则会被
        ``ExpandLayout`` 压扁成空白条。
        """
        card = _SectionCard(parent)
        outer = QVBoxLayout(card)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header = QWidget(card)
        h = QHBoxLayout(header)
        h.setContentsMargins(16, 12, 16, 0)
        h.setSpacing(16)

        icon_widget = IconWidget(icon, header)
        icon_widget.setFixedSize(16, 16)
        h.addWidget(icon_widget, 0, Qt.AlignmentFlag.AlignTop)

        col = QVBoxLayout()
        col.setContentsMargins(0, 0, 0, 0)
        col.setSpacing(2)
        col.addWidget(StrongBodyLabel(title))
        if desc:
            col.addWidget(CaptionLabel(desc))
        h.addLayout(col)
        h.addStretch(1)
        outer.addWidget(header)

        content = QVBoxLayout()
        content.setContentsMargins(16, 10, 16, 14)
        content.setSpacing(8)
        outer.addLayout(content)
        return card, content

    @staticmethod
    def _make_labeled_row(label: str, widget: QWidget, label_width: int = 40) -> QWidget:
        """一行「标签 + 控件」."""
        row = QWidget()
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(10)

        lbl = BodyLabel(label)
        lbl.setFixedWidth(label_width)
        h.addWidget(lbl)
        h.addWidget(widget, 1)
        return row

    # ───────── 个人信息 ─────────

    def _make_profile_group(self, parent: QWidget) -> SettingCardGroup:
        group = SettingCardGroup("个人信息", parent)
        card, content = self._make_section(
            FIF.PEOPLE, "姓名与学号", "将在考试页面顶部显示", group
        )
        self._profile_card = card

        self._name_input = LineEdit()
        self._name_input.setPlaceholderText("请输入姓名")
        self._name_input.setText(self._student_name)
        content.addWidget(self._make_labeled_row("姓名", self._name_input))

        self._id_input = LineEdit()
        self._id_input.setPlaceholderText("请输入学号")
        self._id_input.setText(self._student_id)
        content.addWidget(self._make_labeled_row("学号", self._id_input))

        card.sync_height()
        group.addSettingCard(card)
        return group

    # ───────── 答题 ─────────

    def _make_answer_group(self, parent: QWidget) -> SettingCardGroup:
        group = SettingCardGroup("答题", parent)

        self.feedback_check = SwitchSettingCard(
            FIF.CHECKBOX,
            "答题后立即显示对错",
            "开启后选择答案即时显示对错，关闭后交卷统一评判",
            None,
            group,
        )
        self.feedback_check.switchButton.setChecked(self._show_immediately)
        self.feedback_check.switchButton.checkedChanged.connect(self._update_hint)
        group.addSettingCard(self.feedback_check)

        self.auto_advance_check = SwitchSettingCard(
            FIF.EDIT,
            "选择后自动下一题",
            "选择题/判断题点击选项后自动跳转到下一题",
            None,
            group,
        )
        self.auto_advance_check.switchButton.setChecked(self._auto_advance)
        self.auto_advance_check.switchButton.checkedChanged.connect(self._update_hint)
        group.addSettingCard(self.auto_advance_check)

        strict_card = SettingCard(
            FIF.FONT, "判分严格度", "调整程序题和填空题的判分标准", group
        )
        self._strictness_combo = ComboBox()
        self._strictness_combo.addItems(["严格", "标准", "宽松"])
        self._strictness_combo.setFixedWidth(110)
        strictness_map = {"strict": 0, "normal": 1, "lenient": 2}
        self._strictness_combo.setCurrentIndex(
            strictness_map.get(self._cfg.get("grading_strictness", "normal"), 1)
        )
        strict_card.hBoxLayout.addWidget(self._strictness_combo, 0, Qt.AlignmentFlag.AlignRight)
        strict_card.hBoxLayout.addSpacing(16)
        group.addSettingCard(strict_card)

        return group

    # ───────── 外观 ─────────

    def _make_appearance_group(self, parent: QWidget) -> SettingCardGroup:
        group = SettingCardGroup("外观", parent)

        self.dark_check = SwitchSettingCard(
            FIF.BRUSH,
            "深色模式",
            "切换深色/浅色主题，减少视觉疲劳",
            None,
            group,
        )
        self.dark_check.switchButton.setChecked(Theme._is_dark)
        self.dark_check.switchButton.checkedChanged.connect(self._on_dark_preview)
        group.addSettingCard(self.dark_check)

        card, content = self._make_section(
            FIF.FONT, "字体大小", "调整界面文字大小（80% ~ 150%）", group
        )
        self._appearance_card = card

        slider_row = QWidget()
        sr = QHBoxLayout(slider_row)
        sr.setContentsMargins(0, 0, 0, 0)
        sr.setSpacing(12)

        self._scale_slider = Slider(Qt.Orientation.Horizontal)
        self._scale_slider.setRange(80, 150)
        self._scale_slider.setValue(int(Theme._font_scale * 100))
        self._scale_slider.valueChanged.connect(self._on_scale_changed)
        sr.addWidget(self._scale_slider, 1)

        self._scale_label = CaptionLabel(f"{int(Theme._font_scale * 100)}%")
        self._scale_label.setFixedWidth(40)
        self._scale_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        sr.addWidget(self._scale_label)
        content.addWidget(slider_row)

        preview = QWidget()
        pv = QVBoxLayout(preview)
        pv.setContentsMargins(12, 8, 12, 8)
        pv.setSpacing(2)

        self._preview_title = StrongBodyLabel("这是标题文字的预览效果")
        self._preview_body = BodyLabel("这是正文内容的预览效果，用于确认字体大小是否合适")
        self._preview_small = CaptionLabel("这是小字标注的预览效果")
        pv.addWidget(self._preview_title)
        pv.addWidget(self._preview_body)
        pv.addWidget(self._preview_small)
        content.addWidget(preview)

        self._apply_preview_fonts()
        card.sync_height()
        group.addSettingCard(card)
        return group

    def _apply_preview_fonts(self) -> None:
        scale = Theme._font_scale
        for label, base, weight in (
            (self._preview_title, 12, QFont.Weight.DemiBold),
            (self._preview_body, 11, QFont.Weight.Normal),
            (self._preview_small, 9, QFont.Weight.Normal),
        ):
            font = label.font()
            font.setPointSize(max(7, int(base * scale)))
            font.setWeight(weight)
            label.setFont(font)

    def _on_scale_changed(self, value: int) -> None:
        self._scale_label.setText(f"{value}%")
        Theme._font_scale = value / 100.0
        Theme.update_fonts()
        self._apply_preview_fonts()
        # 预览字号变了，卡片高度要跟着重新算，否则文字会被裁掉
        self._appearance_card.sync_height()
        self._update_hint()

    def _on_dark_preview(self) -> None:
        """深色开关切换时即时预览（点完成才写入配置，取消会还原）."""
        self._sync_fluent_theme(self.dark_check.switchButton.isChecked())
        self._apply_local_style()
        self._update_hint()

    # ───────── 更新 ─────────

    def _make_update_group(self, parent: QWidget) -> SettingCardGroup:
        group = SettingCardGroup("更新", parent)

        self._check_update_card = PushSettingCard(
            "检查更新",
            FIF.UPDATE,
            "软件更新",
            f"当前版本：{self._get_current_version()}",
            group,
        )
        self._check_update_card.clicked.connect(self._check_update_now)
        group.addSettingCard(self._check_update_card)

        card, content = self._make_section(
            FIF.CLOUD_DOWNLOAD, "题库更新", "从 Gitee 远程检查并同步最新题库文件", group
        )
        self._qb_card = card

        current_ver = self._get_question_bank_version()
        self._qb_version_label = CaptionLabel(
            f"当前版本: {current_ver}" if current_ver else "未初始化"
        )
        content.addWidget(self._qb_version_label)

        btn_row = QWidget()
        br = QHBoxLayout(btn_row)
        br.setContentsMargins(0, 0, 0, 0)
        br.setSpacing(8)

        self._check_qb_btn = PushButton("检查题库更新")
        self._check_qb_btn.clicked.connect(self._check_question_bank_now)
        br.addWidget(self._check_qb_btn)

        self._force_qb_btn = PushButton("强制更新")
        self._force_qb_btn.clicked.connect(self._force_update_now)
        br.addWidget(self._force_qb_btn)

        br.addStretch(1)
        content.addWidget(btn_row)

        self._qb_progress_bar = ProgressBar()
        self._qb_progress_bar.setVisible(False)
        self._qb_progress_bar.setRange(0, 100)
        content.addWidget(self._qb_progress_bar)

        self._qb_progress_label = CaptionLabel("")
        self._qb_progress_label.setVisible(False)
        content.addWidget(self._qb_progress_label)

        card.sync_height()
        group.addSettingCard(card)
        return group

    def _set_qb_buttons_enabled(self, enabled: bool) -> None:
        self._check_qb_btn.setEnabled(enabled)
        self._force_qb_btn.setEnabled(enabled)

    def _show_qb_progress(self, visible: bool, text: str = "") -> None:
        self._qb_progress_bar.setVisible(visible)
        self._qb_progress_label.setVisible(visible)
        if text:
            self._qb_progress_label.setText(text)
        # 进度条/文字显隐会改变卡片高度，重新同步一次
        self._qb_card.sync_height()

    def _do_question_bank_update(self, force: bool) -> None:
        if self._checking_qb:
            return
        self._checking_qb = True
        self._set_qb_buttons_enabled(False)
        self._check_qb_btn.setText("检查中..." if not force else "更新中...")
        self._show_qb_progress(True, "正在连接...")
        self._qb_progress_bar.setValue(0)

        from hnust_exam.services.resource_pack_updater import (
            PackUpdateResult,
            check_pack_update_async,
        )

        def _on_progress(downloaded: int, total: int) -> None:
            if total > 0:
                pct = int(downloaded * 100 / total)
                self._qb_progress_bar.setValue(pct)
                if downloaded >= total:
                    self._qb_progress_label.setText("正在完成更新...")
                else:
                    self._qb_progress_label.setText(
                        f"下载中: {pct}% ({downloaded // 1024}KB / {total // 1024}KB)"
                    )

        def _on_status(text: str) -> None:
            self._qb_progress_label.setText(text)

        def _on_result(result: PackUpdateResult) -> None:
            self._checking_qb = False
            self._set_qb_buttons_enabled(True)
            self._check_qb_btn.setText("检查题库更新")
            self._show_qb_progress(False)

            if not result.success:
                error_msg = result.message
                if result.error_type == "network":
                    error_msg = "网络连接失败，请检查网络设置或稍后重试"
                elif result.error_type == "duplicate":
                    error_msg = "另一个更新进程正在运行，请稍候"
                themed_critical(self, "题库更新失败", error_msg)
                return

            if result.new_version:
                self._qb_version_label.setText(f"当前版本: {result.new_version}")

            if "已是最新" in result.message:
                themed_info(
                    self, "题库更新",
                    f"当前题库已是最新版本\n版本号: {result.new_version}",
                )
            elif force:
                themed_info(
                    self, "题库覆盖完成",
                    f"题库已完整覆盖\n版本号: {result.new_version}",
                )
            else:
                themed_info(self, "题库更新成功", result.message)

        check_pack_update_async(
            callback=_on_result,
            progress_callback=_on_progress,
            status_callback=_on_status,
            force_update=force,
        )

    def _check_question_bank_now(self) -> None:
        self._do_question_bank_update(force=False)

    def _force_update_now(self) -> None:
        self._do_question_bank_update(force=True)

    @staticmethod
    def _get_question_bank_version() -> str:
        from hnust_exam.services.resource_pack_updater import _get_local_version
        return _get_local_version()

    # ───────── 试卷标记 ─────────

    def _make_mark_group(self, parent: QWidget) -> SettingCardGroup:
        group = SettingCardGroup("试卷标记", parent)

        legend_card = PushSettingCard(
            "标记说明", FIF.INFO, "查看标记含义",
            "✓ 已完成　○ 进行中　（无标记）尚未打开", group,
        )
        legend_card.clicked.connect(self._show_mark_legend)
        group.addSettingCard(legend_card)

        clear_card = PushSettingCard(
            "清除标记", FIF.DELETE, "清除所有试卷标记",
            "重置所有试卷的完成状态与得分记录", group,
        )
        clear_card.clicked.connect(self._clear_all_marks)
        group.addSettingCard(clear_card)

        return group

    # ───────── 业务方法 ─────────

    def _on_done(self) -> None:
        self._show_immediately = self.feedback_check.switchButton.isChecked()
        self._auto_advance = self.auto_advance_check.switchButton.isChecked()
        Theme.set_dark_mode(self.dark_check.switchButton.isChecked())
        self._sync_fluent_theme()

        strictness = _STRICTNESS_VALUES[self._strictness_combo.currentIndex()]

        new_cfg = {
            "font_scale": Theme._font_scale,
            "dark_mode": Theme._is_dark,
            "show_answer_immediately": self._show_immediately,
            "auto_advance_on_choice": self._auto_advance,
            "user_python_path": self.config_mgr.load_config().get("user_python_path", ""),
            "student_name": self._name_input.text().strip(),
            "student_id": self._id_input.text().strip(),
            "grading_strictness": strictness,
        }
        old_cfg = self.config_mgr.load_config()
        for k in old_cfg:
            if k not in new_cfg:
                new_cfg[k] = old_cfg[k]
        self.config_mgr.save_config(new_cfg)
        self.accept()

    def _on_cancel(self) -> None:
        self.reject()

    def reject(self) -> None:
        """取消/关闭时还原预览期间的字体缩放与主题.

        对话框里的字体滑块和深色开关都是即时预览，只有点「完成」才写配置。
        这里统一兜底，避免用右上角 X 关闭时把预览状态泄漏到主界面。
        """
        Theme._font_scale = self._orig_scale
        Theme.update_fonts()
        Theme.set_dark_mode(self._orig_dark)
        self._sync_fluent_theme()
        super().reject()

    def _update_hint(self) -> None:
        if not hasattr(self, "_status_hint"):
            return
        feedback_on = self.feedback_check.switchButton.isChecked()
        dark = self.dark_check.switchButton.isChecked()
        mode = "即时反馈" if feedback_on else "考试模式"
        theme = "深色" if dark else "浅色"
        scale = int(Theme._font_scale * 100)
        self._status_hint.setText(f"答题：{mode}  |  主题：{theme}  |  字体：{scale}%")
        colors = Theme._DARK if dark else Theme._LIGHT
        accent = colors["SUCCESS"] if feedback_on else colors["PRIMARY"]
        self._status_hint.setStyleSheet(f"color: {accent};")

    def _show_mark_legend(self) -> None:
        themed_info(
            self, "标记说明",
            "✓ 已完成 — 已交卷的试卷，后面显示最高得分\n\n"
            "○ 进行中 — 已开始但尚未交卷的试卷\n\n"
            "（无标记）— 尚未打开过的试卷"
        )

    def _clear_all_marks(self) -> None:
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

    def _check_update_now(self) -> None:
        if self._checking_update:
            return
        self._checking_update = True
        self._check_update_card.button.setEnabled(False)
        self._check_update_card.button.setText("检查中...")

        class _Sig(QObject):
            done = Signal(object)

        sig = _Sig()
        sig.done.connect(self._on_manual_update_result)

        def _worker():
            try:
                info = fetch_update_info(self.config_mgr)
                if info is None:
                    sig.done.emit(None)  # 网络错误
                elif info.get("no_update"):
                    info["update_available"] = False
                    sig.done.emit(info)  # 已是最新
                else:
                    sig.done.emit(info)  # 有更新
            except Exception:
                sig.done.emit(None)

        Thread(target=_worker, daemon=True).start()

    def _on_manual_update_result(self, info: dict | None) -> None:
        self._checking_update = False
        self._check_update_card.button.setEnabled(True)
        self._check_update_card.button.setText("检查更新")

        if not info:
            themed_warning(
                self,
                "检查失败",
                "暂时无法获取最新版本信息。\n请检查网络连接后再试。",
            )
            return

        current_version = self._get_current_version()
        latest_version = info["latest_ver"]

        if not info.get("update_available", True):
            themed_info(
                self,
                "检查更新",
                f"已经是最新版本啦！\n\n当前版本：{current_version}\n最新版本：{latest_version}",
            )
            return

        info = dict(info)
        info["current_ver"] = current_version
        dlg = UpdateDialog(info, self.config_mgr, self)
        dlg.exec()

    @staticmethod
    def _get_current_version() -> str:
        """读取当前版本号.

        开发态（非打包）直接扫一遍 ``constants.py`` 源文件，改完版本号不用重启
        就能看到。这里刻意不用 ``importlib.reload``——它会重新执行整个模块，
        把运行期改过的常量（配置目录、各种文件路径）全部打回文件里的初值，
        进而把配置写到用户真实目录去。
        """
        if not getattr(sys, "frozen", False):
            try:
                with open(constants.__file__, "r", encoding="utf-8") as f:
                    for line in f:
                        if line.startswith("CURRENT_VERSION"):
                            return line.split("=", 1)[1].strip().strip('"').strip("'")
            except Exception:
                pass
        return constants.CURRENT_VERSION
