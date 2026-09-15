"""成绩页进度条回归测试.

背景（真实 bug）：``ResultPage._rebuild_progress()`` 原来写的是

    fill = max(1, int(pct))
    empty = max(1, 100 - fill)

``max(1, ...)`` 的用意是"别让某一侧变成 0 宽度"，但副作用是
**满分时 ``empty`` 被强行撑成 1**，于是 100% 的进度条右侧永远留一条
1px 的空白 —— 看起来像没填满。

正确做法是按 0 处理：0 份就不添加那一项，让实心条真正铺满。
"""

from __future__ import annotations

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from hnust_exam.app import _generate_stylesheet
from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.theme import Theme
from hnust_exam.views.main_window import MainWindow


# ───────── 夹具 ─────────


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


@pytest.fixture
def theme_guard():
    """还原 Theme 全局状态，避免用例互相污染."""
    dark, scale = Theme._is_dark, Theme._font_scale
    yield
    Theme.set_dark_mode(dark)
    Theme._font_scale = scale
    Theme.update_fonts()


@pytest.fixture
def style_guard(app):
    """还原 QApplication 的全局样式表，别泄漏给别的测试模块."""
    before = app.styleSheet()
    yield app
    app.setStyleSheet(before)


@pytest.fixture
def page(app, theme_guard, style_guard):
    """离屏构建一个 MainWindow，取它的成绩页."""
    Theme.set_dark_mode(False)
    app.setStyleSheet(_generate_stylesheet())
    mw = MainWindow(ConfigManager())
    mw.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
    yield mw.result_page
    mw.close()
    mw.deleteLater()
    app.processEvents()


def _layout_shape(page) -> list[tuple[str, int]]:
    """把进度条布局读成 [(类型, 权重)] 列表."""
    lay = page._prog_lay
    shape = []
    for i in range(lay.count()):
        item = lay.itemAt(i)
        w = item.widget()
        kind = type(w).__name__ if w is not None else "stretch"
        shape.append((kind, lay.stretch(i)))
    return shape


# ───────── 用例 ─────────


class TestProgressBarLayout:
    """进度条在边界值上的布局."""

    def test_full_score_has_no_trailing_gap(self, page):
        """满分时只能是「实心条独占」，不能留空白.

        这是本次修复的核心：``max(1, 100 - fill)`` 会在满分时留下 1px 空白。
        """
        page._rebuild_progress("#00c853", 100.0)
        shape = _layout_shape(page)

        assert shape == [("QFrame", 100)], f"满分应只有实心条，实际 {shape}"
        assert all(kind != "stretch" for kind, _ in shape), "满分不应有任何空白占位"

    def test_zero_score_has_no_zero_width_widget(self, page):
        """0 分时只留空白，不要塞一个权重为 0 的控件."""
        page._rebuild_progress("#ff5252", 0.0)
        shape = _layout_shape(page)

        assert shape == [("stretch", 100)], f"0 分应只有空白，实际 {shape}"

    def test_mid_score_splits_evenly(self, page):
        """50 分时实心条和空白各占一半."""
        page._rebuild_progress("#2196f3", 50.0)
        shape = _layout_shape(page)

        assert shape == [("QFrame", 50), ("stretch", 50)], f"实际 {shape}"

    @pytest.mark.parametrize("pct", [0.4, 1.0, 33.3, 66.6, 99.4])
    def test_weights_always_sum_to_100(self, page, pct):
        """任意分数下权重之和恒为 100，保证条长比例正确."""
        page._rebuild_progress("#2196f3", pct)
        shape = _layout_shape(page)

        assert sum(w for _, w in shape) == 100, f"pct={pct} 权重和不为 100：{shape}"

    @pytest.mark.parametrize("pct,expected", [(-10.0, 0), (150.0, 100), (1000.0, 100)])
    def test_out_of_range_is_clamped(self, page, pct, expected):
        """越界分数要被夹到 [0, 100]，不能撑爆布局."""
        page._rebuild_progress("#2196f3", pct)
        shape = _layout_shape(page)

        total_fill = sum(w for kind, w in shape if kind == "QFrame")
        assert total_fill == expected, f"pct={pct} 应夹到 {expected}，实际 {shape}"

    def test_repeated_rebuild_does_not_accumulate(self, page):
        """反复重建不能残留旧控件（旧实现靠 takeAt 清理，要确认真的清干净）."""
        for pct in (100.0, 0.0, 50.0, 100.0):
            page._rebuild_progress("#2196f3", pct)
        shape = _layout_shape(page)

        assert len(shape) <= 2, f"反复重建后布局残留过多项：{shape}"
        assert shape == [("QFrame", 100)], f"最后一次是满分，实际 {shape}"
