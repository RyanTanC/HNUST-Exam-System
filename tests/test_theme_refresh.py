"""主题切换后各页面是否真的跟着变——回归测试.

背景（真实 bug）：设置页只能从「选卷页」打开，所以用户切深色时
``MainWindow._refresh_theme()`` 的当前页恰好是选卷页，只有选卷页被重建；
``ExamPage`` / ``QuestionWidget`` / ``NavPanel`` / ``ResultPage`` 里那些
**在构造时就把当时的主题色写死进内联 QSS** 的控件全都没动。

结果就是用户看到的「半深色」答题页：全局样式表已经是深色（导航列表底色、
按钮底色、文字颜色都变深了），而页面自己的内联底色还是浅色（顶部蓝条、
进度条区、底部按钮栏、题目白底、选项提示条……）。把应用关掉重开才好，
因为重启时 ``Theme.set_dark_mode()`` 跑在 ``MainWindow()`` 之前，
构造时写进去的就是深色值了。

这里盯三件事：

1. 切主题后**不在前台的页面**也要刷新（切过去时刷，不能漏）。
2. ``NavPanel.refresh_theme()`` 必须连分组标题/分隔线一起刷
   （``refresh()`` 只重刷题目按钮，刷不到 ``_build_panels()`` 里写死的那几个）。
3. 浅色↔深色来回切都要对，不能只修单向。
"""

from __future__ import annotations

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QPushButton

from hnust_exam.app import _generate_stylesheet
from hnust_exam.models.question import Question
from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.theme import Theme
from hnust_exam.views.main_window import MainWindow
from hnust_exam.views.nav_panel import NavPanel

LIGHT = Theme._LIGHT
DARK = Theme._DARK


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


def _flush(app: QApplication, times: int = 4) -> None:
    for _ in range(times):
        app.processEvents()


def _build_window(app: QApplication, dark: bool = False) -> MainWindow:
    """按 app.run() 的真实顺序建主窗口：先定主题，再构造页面."""
    Theme.set_dark_mode(dark)
    app.setStyleSheet(_generate_stylesheet())
    mw = MainWindow(ConfigManager())
    mw.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
    mw.resize(1200, 800)
    mw.show()
    _flush(app)
    return mw


# 答题页里「构造时写死颜色」的静态控件：(名字, 取值函数, Theme 色键)
EXAM_STATIC_WIDGETS = [
    ("_top_bar", lambda p: p._top_bar, "PRIMARY"),
    ("_progress_container", lambda p: p._progress_container, "BG"),
    ("_bottom_bar", lambda p: p._bottom_bar, "BG"),
    ("_submit_btn", lambda p: p._submit_btn, "DANGER"),
    ("_watermark", lambda p: p._watermark, "MUTED"),
    ("question._content", lambda p: p.question_widget._content, "WHITE"),
    ("question._title_bar", lambda p: p.question_widget._title_bar, "PRIMARY"),
    ("question._instruction_label", lambda p: p.question_widget._instruction_label, "TEXT"),
    ("question._question_text", lambda p: p.question_widget._question_text, "TEXT"),
    ("question._hint_label", lambda p: p.question_widget._hint_label, "HINT_BG"),
    ("question._kb_hint", lambda p: p.question_widget._kb_hint, "KB_HINT_BG"),
    ("nav._frame", lambda p: p.nav_panel._frame, "WHITE"),
    ("nav._header", lambda p: p.nav_panel._header, "TEXT"),
    ("nav._sep", lambda p: p.nav_panel._sep, "BORDER"),
    ("nav._nav_content", lambda p: p.nav_panel._nav_content, "WHITE"),
    ("nav._status_label", lambda p: p.nav_panel._status_label, "MUTED"),
]


def _assert_page_matches(page, palette: dict, other: dict) -> None:
    """页面上所有静态控件的内联 QSS 都必须是 palette 的颜色，不能是 other 的."""
    for name, getter, key in EXAM_STATIC_WIDGETS:
        qss = getter(page).styleSheet()
        assert palette[key] in qss, f"{name} 没跟上主题，缺 {palette[key]}：{qss!r}"
        assert other[key] not in qss, f"{name} 还残留旧主题色 {other[key]}：{qss!r}"


def _px(widget, x: int | None = None, y: int | None = None) -> str:
    """渲染单个控件并取一个像素.

    用 ``widget.grab()`` 而不是窗口级截图：页面切换动画还在跑的时候，
    窗口截图会拍到未绘制的区域，结果不稳定。

    取色点默认取控件正中；正中可能压着子控件（按钮、文字）时显式传坐标。
    高 DPI 下图像是设备像素，得按 devicePixelRatio 换算。
    """
    img = widget.grab().toImage()
    dpr = img.devicePixelRatio() or 1.0
    lx = widget.width() // 2 if x is None else x
    ly = widget.height() // 2 if y is None else y
    return img.pixelColor(int(lx * dpr), int(ly * dpr)).name()


# ───────── 核心：切主题后每个页面都要刷 ─────────


class TestExamPageFollowsTheme:
    """答题页的半深色 bug 不许回来."""

    def test_offscreen_exam_page_is_refreshed_on_switch(
        self, app, theme_guard, style_guard
    ):
        """真实路径：浅色启动 → 在选卷页切深色 → 进答题页."""
        mw = _build_window(app, dark=False)
        mw.show_select()
        _flush(app)

        # 用户在设置里切到深色并点「完成」
        Theme.set_dark_mode(True)
        mw._refresh_theme()
        _flush(app)

        # 进入答题页（= 点「开始答题」）
        mw.show_exam()
        _flush(app)

        _assert_page_matches(mw.exam_page, DARK, LIGHT)

    def test_switch_back_to_light_also_works(self, app, theme_guard, style_guard):
        """深色启动 → 切回浅色 → 答题页必须整体变回浅色（反向也要对）."""
        mw = _build_window(app, dark=True)
        mw.show_select()
        _flush(app)

        Theme.set_dark_mode(False)
        mw._refresh_theme()
        _flush(app)
        mw.show_exam()
        _flush(app)

        _assert_page_matches(mw.exam_page, LIGHT, DARK)

    def test_result_page_follows_theme(self, app, theme_guard, style_guard):
        mw = _build_window(app, dark=False)
        mw.show_select()
        _flush(app)

        Theme.set_dark_mode(True)
        mw._refresh_theme()
        _flush(app)
        mw.show_result()
        _flush(app)

        rp = mw.result_page
        for name, widget, key in [
            ("_header_frame", rp._header_frame, "PRIMARY"),
            ("_body", rp._body, "BG"),
            ("_footer_frame", rp._footer_frame, "BG"),
        ]:
            assert DARK[key] in widget.styleSheet(), f"{name} 没跟上主题"
            assert LIGHT[key] not in widget.styleSheet(), f"{name} 残留浅色"

    def test_rendered_pixels_match_dark_theme(self, app, theme_guard, style_guard):
        """直接渲染答题页取像素——就是用户截图里出错的那几个位置."""
        mw = _build_window(app, dark=False)
        mw.show_select()
        _flush(app)
        Theme.set_dark_mode(True)
        mw._refresh_theme()
        _flush(app)
        mw.show_exam()
        _flush(app)

        page = mw.exam_page
        # 顶部蓝条正中是空白区，直接取中心
        assert _px(page._top_bar) == DARK["PRIMARY"], "顶部蓝条还是浅色"
        # 底部按钮栏：x=2 落在左内边距里，避开所有按钮
        assert _px(page._bottom_bar, 2, 30) == DARK["BG"], "底部按钮栏还是浅色"
        # 题目区：y=4 落在 20px 上边距里，避开题干文字
        assert _px(page.question_widget._content, None, 4) == DARK["WHITE"], "题目区还是白底"


class TestStalePageBookkeeping:
    """不在前台的页面要记账，等切过去时刷——不能漏也不能白重建."""

    def test_all_pages_marked_stale_after_refresh(self, app, theme_guard, style_guard):
        mw = _build_window(app, dark=False)
        mw.show_welcome()
        _flush(app)

        mw._refresh_theme()
        # 当前页（欢迎页）当场刷掉，其余三页记账
        assert MainWindow.PAGE_WELCOME not in mw._theme_stale
        assert mw._theme_stale == {
            MainWindow.PAGE_SELECT,
            MainWindow.PAGE_EXAM,
            MainWindow.PAGE_RESULT,
        }

    def test_switching_to_page_consumes_its_stale_mark(
        self, app, theme_guard, style_guard
    ):
        mw = _build_window(app, dark=False)
        mw.show_welcome()
        _flush(app)
        mw._refresh_theme()

        mw.show_exam()
        _flush(app)
        assert MainWindow.PAGE_EXAM not in mw._theme_stale

        mw.show_result()
        _flush(app)
        assert MainWindow.PAGE_RESULT not in mw._theme_stale

    def test_no_refresh_needed_without_theme_change(
        self, app, theme_guard, style_guard
    ):
        """没切过主题时，切页面不该被记账逻辑拖着重刷."""
        mw = _build_window(app, dark=False)
        exam_page_before = mw.exam_page

        mw.show_exam()
        _flush(app)
        assert mw.exam_page is exam_page_before, "答题页被无谓地重建了（考试状态会丢）"


# ───────── NavPanel 分组标题 ─────────


class _FakeExam:
    """只喂 NavPanel 需要的那几个字段，避免依赖真实 xlsx."""

    def __init__(self, questions: list[Question]) -> None:
        self.questions = questions
        self.question_groups: dict[str, list[Question]] = {}
        for q in questions:
            self.question_groups.setdefault(q.q_type, []).append(q)
        self.active_type_order = list(self.question_groups)
        self.current_index = 0
        self.answer_map: dict = {}
        self.marked_count = 0
        self.answered_count = 0

    @property
    def total_count(self) -> int:
        return len(self.questions)

    @property
    def unanswered_count(self) -> int:
        return self.total_count - self.answered_count

    def is_marked(self, index: int) -> bool:
        return False


class _FakeExamPage:
    """NavPanel 只会用到 exam / _buttons / _update_progress."""

    def __init__(self, exam) -> None:
        self.exam = exam
        self._buttons = {"标记试题": QPushButton("标记试题")}
        self.progress_calls = 0

    def _update_progress(self) -> None:
        self.progress_calls += 1


class TestNavPanelTheme:
    """refresh() 只重刷题目按钮，分组标题/分隔线得靠 refresh_theme() 重建."""

    @staticmethod
    def _build(app, dark: bool) -> NavPanel:
        Theme.set_dark_mode(dark)
        questions = [
            Question(index=0, number="1", q_type="单选", text="题一"),
            Question(index=1, number="2", q_type="单选", text="题二"),
            Question(index=2, number="3", q_type="判断", text="题三"),
        ]
        page = _FakeExamPage(_FakeExam(questions))
        panel = NavPanel(page)
        panel.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
        panel.resize(250, 600)
        panel.show()
        # 真实流程里 setup_exam() 会调 reset() + refresh() 把分组面板建出来
        panel.refresh()
        _flush(app)
        return panel

    def test_group_header_follows_theme(self, app, theme_guard):
        """深色下重新 refresh_theme()，分组标题必须换成深色底."""
        panel = self._build(app, dark=False)
        q_type = "单选"
        header = panel._panels[q_type]["header"]
        assert LIGHT["NAV_HEADER_BG"] in header.styleSheet()

        Theme.set_dark_mode(True)
        panel.refresh_theme()
        _flush(app)

        header = panel._panels[q_type]["header"]
        assert DARK["NAV_HEADER_BG"] in header.styleSheet(), "分组标题还是浅色"
        assert LIGHT["NAV_HEADER_BG"] not in header.styleSheet()

    def test_separator_and_body_follow_theme(self, app, theme_guard):
        panel = self._build(app, dark=False)
        Theme.set_dark_mode(True)
        panel.refresh_theme()
        _flush(app)

        body = panel._panels["单选"]["body"]
        assert DARK["WHITE"] in body.styleSheet()
        assert LIGHT["WHITE"] not in body.styleSheet()

        # 「判断」是第二个分组，前面有一条分隔线
        sep = panel._panels["判断"]["sep"]
        assert sep is not None
        assert DARK["BORDER"] in sep.styleSheet()
        assert LIGHT["BORDER"] not in sep.styleSheet()

    def test_refresh_theme_restyles_in_place(self, app, theme_guard):
        """必须是原地改样式，不能重建面板.

        重建会删掉还挂着入场淡入动画的控件，动画的 finished 回调再去碰
        已析构的 QWidget 就会抛 "Internal C++ object already deleted"。
        """
        panel = self._build(app, dark=False)
        before = {
            k: (v["header"], v["body"], v["title"])
            for k, v in panel._panels.items()
        }

        Theme.set_dark_mode(True)
        panel.refresh_theme()
        _flush(app)

        for k, (header, body, title) in before.items():
            assert panel._panels[k]["header"] is header, f"{k} 的分组标题被重建了"
            assert panel._panels[k]["body"] is body, f"{k} 的按钮容器被重建了"
            assert panel._panels[k]["title"] is title, f"{k} 的标题文字被重建了"

    def test_question_buttons_are_restyled(self, app, theme_guard):
        panel = self._build(app, dark=False)
        Theme.set_dark_mode(True)
        panel.refresh_theme()
        _flush(app)

        btn = panel._q_buttons[0]  # 当前题
        assert DARK["NAV_CURRENT"] in btn.styleSheet()
        assert LIGHT["NAV_CURRENT"] not in btn.styleSheet()

    def test_refresh_theme_without_exam_does_not_crash(self, app, theme_guard):
        Theme.set_dark_mode(False)
        page = _FakeExamPage(None)
        panel = NavPanel(page)
        panel.refresh_theme()
        assert panel._panels == {}
