"""设置对话框（qfluentwidgets 版）回归测试.

重点盯三类容易回退的地方：

1. **自定义卡片高度**——qfluentwidgets 的 ``ExpandLayout`` 用 ``widget.height()``
   （当前高度）而不是 ``sizeHint()`` 累加分组高度。标准 ``SettingCard`` 靠
   ``setFixedHeight`` 规避，自定义卡片必须自己同步，否则会被压扁成空白条。
2. **完成/取消的语义**——点「完成」才写配置，点「取消」或直接关窗要还原预览。
3. **不碰真实用户目录**——配置、进度、题库版本文件都重定向到临时目录。
"""

from __future__ import annotations

import pytest
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication

from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.theme import Theme
from hnust_exam.views.dialogs.settings_dialog import SettingsDialog


def _flush(app: QApplication, times: int = 4) -> None:
    """把布局事件跑完，让几何生效."""
    for _ in range(times):
        app.processEvents()


@pytest.fixture(scope="module")
def app():
    """创建 QApplication 实例."""
    return QApplication.instance() or QApplication([])


@pytest.fixture
def theme_guard():
    """保存并还原 Theme 的全局状态，避免用例之间互相污染."""
    dark, scale = Theme._is_dark, Theme._font_scale
    yield
    Theme.set_dark_mode(dark)
    Theme._font_scale = scale
    Theme.update_fonts()


@pytest.fixture
def cfg_mgr():
    """配置读写走临时目录——路径重定向在 tests/conftest.py 里统一做掉了."""
    return ConfigManager()


def _open(app, cfg_mgr) -> SettingsDialog:
    """建一个不真正弹到屏幕上的设置对话框（仍会走完整布局）."""
    dlg = SettingsDialog(cfg_mgr, None)
    dlg.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
    dlg.resize(580, 720)
    dlg.show()
    _flush(app)
    return dlg


# ───────── 卡片高度 ─────────


class TestCustomCardHeight:
    """自定义卡片不能被 ExpandLayout 压扁."""

    def test_appearance_card_keeps_its_height(self, app, cfg_mgr, theme_guard):
        dlg = _open(app, cfg_mgr)
        card = dlg._appearance_card
        assert card.height() == card.sizeHint().height()
        assert card.height() > 120, "字体大小卡片被压扁了"

    def test_question_bank_card_keeps_its_height(self, app, cfg_mgr, theme_guard):
        dlg = _open(app, cfg_mgr)
        card = dlg._qb_card
        assert card.height() == card.sizeHint().height()
        assert card.height() > 90, "题库更新卡片被压扁了"

    def test_profile_inputs_are_inside_the_card(self, app, cfg_mgr, theme_guard):
        """姓名/学号输入框必须完整落在卡片范围内（而不是被裁掉）."""
        dlg = _open(app, cfg_mgr)
        card = dlg._profile_card
        for widget in (dlg._name_input, dlg._id_input):
            top = widget.mapTo(card, widget.rect().topLeft()).y()
            assert top >= 0, "输入框跑到了卡片上边缘之外"
            assert top + widget.height() <= card.height(), "输入框被卡片下边缘裁掉了"

    def test_progress_toggle_resizes_card(self, app, cfg_mgr, theme_guard):
        """进度条显隐要跟着改变卡片高度."""
        dlg = _open(app, cfg_mgr)
        collapsed = dlg._qb_card.height()

        dlg._show_qb_progress(True, "下载中: 42%")
        _flush(app)
        assert dlg._qb_card.height() > collapsed

        dlg._show_qb_progress(False)
        _flush(app)
        assert dlg._qb_card.height() == collapsed


# ───────── 完成 / 取消 ─────────


class TestDoneAndCancel:
    """完成写配置，取消还原预览."""

    def test_done_persists_config(self, app, cfg_mgr, theme_guard):
        dlg = _open(app, cfg_mgr)
        dlg.feedback_check.switchButton.setChecked(True)
        dlg.auto_advance_check.switchButton.setChecked(True)
        dlg._strictness_combo.setCurrentIndex(2)  # lenient
        dlg._name_input.setText("张三")
        dlg._id_input.setText("240110011")
        dlg._on_done()

        cfg = cfg_mgr.load_config()
        assert cfg["show_answer_immediately"] is True
        assert cfg["auto_advance_on_choice"] is True
        assert cfg["grading_strictness"] == "lenient"
        assert cfg["student_name"] == "张三"
        assert cfg["student_id"] == "240110011"

    def test_done_keeps_unrelated_keys(self, app, cfg_mgr, theme_guard):
        """保存时不能把用户已有的其它配置项丢掉."""
        cfg_mgr.save_config({"user_python_path": "D:/Python/python.exe"})
        dlg = _open(app, cfg_mgr)
        dlg._on_done()
        assert cfg_mgr.load_config()["user_python_path"] == "D:/Python/python.exe"

    def test_reject_restores_font_scale(self, app, cfg_mgr, theme_guard):
        Theme.set_dark_mode(False)
        Theme._font_scale = 1.0
        Theme.update_fonts()

        dlg = _open(app, cfg_mgr)
        dlg._scale_slider.setValue(150)
        assert Theme._font_scale == pytest.approx(1.5)

        dlg.reject()
        assert Theme._font_scale == pytest.approx(1.0)

    def test_reject_restores_dark_mode(self, app, cfg_mgr, theme_guard):
        """深色开关是实时预览，取消（含右上角 X）必须还原."""
        Theme.set_dark_mode(False)
        dlg = _open(app, cfg_mgr)
        dlg.dark_check.switchButton.setChecked(True)

        dlg.reject()
        assert Theme._is_dark is False

    def test_reject_does_not_write_config(self, app, cfg_mgr, theme_guard):
        dlg = _open(app, cfg_mgr)
        dlg._name_input.setText("不该被保存")
        dlg.reject()
        assert cfg_mgr.load_config().get("student_name", "") != "不该被保存"


# ───────── 状态提示 ─────────


class TestStatusHint:
    """底部状态提示随开关变化."""

    def test_hint_reflects_current_state(self, app, cfg_mgr, theme_guard):
        dlg = _open(app, cfg_mgr)
        dlg.feedback_check.switchButton.setChecked(False)
        dlg.dark_check.switchButton.setChecked(False)
        dlg._update_hint()
        text = dlg._status_hint.text()
        assert "考试模式" in text
        assert "浅色" in text
        assert "字体：100%" in text

    def test_hint_tracks_font_scale(self, app, cfg_mgr, theme_guard):
        dlg = _open(app, cfg_mgr)
        dlg._scale_slider.setValue(120)
        assert "字体：120%" in dlg._status_hint.text()


# ───────── 主题色 ─────────


class TestThemeColor:
    """qfluentwidgets 的主色必须跟应用的 Theme.PRIMARY 对齐.

    qfluentwidgets 默认主色是青色 #009faa，不对齐的话设置页的开关、滑块、
    主按钮会和软件其它部分两个色。
    """

    def test_light_accent_matches_app_primary(self, app, cfg_mgr, theme_guard):
        from qfluentwidgets import themeColor

        Theme.set_dark_mode(False)
        _open(app, cfg_mgr)
        assert themeColor().name() == Theme._LIGHT["PRIMARY"]

    def test_dark_accent_keeps_app_hue(self, app, cfg_mgr, theme_guard):
        """深色下库会降低饱和度（s *= 0.84），但色相必须和应用一致."""
        from qfluentwidgets import themeColor

        Theme.set_dark_mode(True)
        _open(app, cfg_mgr)
        got = themeColor()
        want = QColor(Theme._DARK["PRIMARY"])
        assert got.hue() == want.hue()
        assert got.saturation() <= want.saturation()

    def test_accent_is_not_the_library_default(self, app, cfg_mgr, theme_guard):
        """别退回到 qfluentwidgets 自带的青色."""
        from qfluentwidgets import themeColor

        Theme.set_dark_mode(False)
        _open(app, cfg_mgr)
        assert themeColor().name() != "#009faa"

    def test_dark_preview_switches_accent(self, app, cfg_mgr, theme_guard):
        """深色开关是实时预览，主色要跟着切."""
        from qfluentwidgets import themeColor

        Theme.set_dark_mode(False)
        dlg = _open(app, cfg_mgr)
        assert themeColor().name() == Theme._LIGHT["PRIMARY"]

        dlg.dark_check.switchButton.setChecked(True)
        assert themeColor().hue() == QColor(Theme._DARK["PRIMARY"]).hue()

    def test_reject_restores_accent(self, app, cfg_mgr, theme_guard):
        from qfluentwidgets import themeColor

        Theme.set_dark_mode(False)
        dlg = _open(app, cfg_mgr)
        dlg.dark_check.switchButton.setChecked(True)
        dlg.reject()
        assert themeColor().name() == Theme._LIGHT["PRIMARY"]

    def test_does_not_write_fluent_config(
        self, app, cfg_mgr, theme_guard, tmp_path, monkeypatch
    ):
        """setThemeColor 必须带 save=False.

        qfluentwidgets 的配置文件默认落在**相对路径** config/config.json，
        也就是进程当前目录。带 save=True 的话会在用户的工作目录里拉一坨垃圾。
        """
        monkeypatch.chdir(tmp_path)
        _open(app, cfg_mgr)
        assert not (tmp_path / "config").exists(), "qfluentwidgets 往当前目录写了配置文件"


# ───────── 试卷标记 ─────────


class TestMarkClearing:
    def test_clear_all_marks_sets_flag(self, app, cfg_mgr, theme_guard, monkeypatch):
        cfg_mgr.save_progress({"卷子A": {"done": True}})
        dlg = _open(app, cfg_mgr)

        import hnust_exam.views.dialogs.settings_dialog as sd
        from PySide6.QtWidgets import QMessageBox

        monkeypatch.setattr(sd, "themed_info", lambda *a, **k: None)
        monkeypatch.setattr(
            sd, "themed_question", lambda *a, **k: QMessageBox.StandardButton.Yes
        )
        dlg._clear_all_marks()

        assert dlg._marks_cleared is True
        assert cfg_mgr.load_progress() == {}

    def test_clear_all_marks_noop_when_empty(self, app, cfg_mgr, theme_guard, monkeypatch):
        dlg = _open(app, cfg_mgr)
        import hnust_exam.views.dialogs.settings_dialog as sd

        monkeypatch.setattr(sd, "themed_info", lambda *a, **k: None)
        dlg._clear_all_marks()
        assert dlg._marks_cleared is False
