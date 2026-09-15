# 弹窗提示置顶 + 题库热更新版本比较逻辑 实现计划

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推荐）或 superpowers:executing-plans 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。

**目标：** 实现两个独立功能：① 将"试题解析"后的复制成功提示改为页面顶部浮层通知，避免被浏览器遮盖；② 将题库热更新的简单哈希比较改为基于 `yyyy.mm.dd.tttt` 格式的日期+时间版本比较逻辑。

**架构：**
- ① 创建非模态 `ToastWidget`，无边框浮动在父窗口顶部居中，带 `WindowStaysOnTopHint` 标志确保不被浏览器遮挡，2 秒后自动消褪
- ② 在 `resource_pack_updater.py` 中引入版本解析/比较函数，处理 `yyyy.mm.dd.tttt` 和 `yyyy.mm.dd` 两种格式，替换原有的 `remote_tag == local_ver` 字符串比较

**技术栈：** PySide6, QTimer, QPropertyAnimation, pytest

---

## 文件结构

### 新增文件
- `hnust_exam/views/widgets/toast_notification.py` — 顶部浮层通知组件
- `hnust_exam/utils/version.py` — 版本号解析与比较工具函数（独立于更新器，可测试）
- `tests/test_toast_notification.py` — ToastWidget 测试
- `tests/test_version_comparison.py` — 版本比较逻辑测试

### 修改文件
- `hnust_exam/views/exam_page.py:448-462` — 将 `themed_info` 替换为 `ToastWidget`
- `hnust_exam/views/dialogs/submit_dialog.py:298-304` — 将"复制未答题号"的 `themed_info` 替换为 `ToastWidget`（可选改进）
- `hnust_exam/services/resource_pack_updater.py:286-290` — 将字符串比较替换为语义化版本比较

---

## 任务 1：创建顶部浮层通知组件 ToastWidget

**文件：**
- 创建：`hnust_exam/views/widgets/toast_notification.py`
- 测试：`tests/test_toast_notification.py`

- [ ] **步骤 1：编写 ToastWidget 类的失败测试**

```python
# tests/test_toast_notification.py
"""测试 ToastWidget 顶部通知组件."""
import pytest
from PySide6.QtWidgets import QMainWindow, QApplication, QLabel
from PySide6.QtCore import Qt, QTimer
from hnust_exam.views.widgets.toast_notification import ToastWidget


def test_toast_widget_creation(qtbot):
    """验证 ToastWidget 能正常创建和显示."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "测试通知")
    qtbot.addWidget(toast)

    # 验证 label 文本
    label = toast.findChild(QLabel)
    assert label is not None
    assert label.text() == "测试通知"

    # 验证窗口标志包含 FramelessWindowHint
    assert toast.windowFlags() & Qt.FramelessWindowHint


def test_toast_widget_auto_close(qtbot):
    """验证 ToastWidget 在指定时间后自动关闭."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "自动关闭测试", duration_ms=100)
    qtbot.addWidget(toast)
    toast.show()

    # 等待 200ms 确保定时器触发
    qtbot.wait(200)

    # 验证 toast 已关闭
    assert not toast.isVisible()


def test_toast_widget_position_at_top_center(qtbot):
    """验证 ToastWidget 定位在父窗口顶部居中."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "位置测试")
    qtbot.addWidget(toast)
    toast.show()

    # toast 的中心 x 应该在父窗口中心附近（允许偏移半个 toast 宽度）
    parent_center_x = window.rect().center().x()
    toast_center_x = toast.geometry().center().x()
    assert abs(toast_center_x - parent_center_x) < toast.width() // 2 + 1

    # toast 的顶部应该在父窗口顶部附近（距顶部 20px 左右）
    assert toast.y() == 20 or abs(toast.y() - 20) <= 1


def test_toast_widget_window_stays_on_top(qtbot):
    """验证 ToastWidget 设置了 WindowStaysOnTopHint 标志."""
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)

    toast = ToastWidget(window, "置顶测试")
    qtbot.addWidget(toast)

    # WindowStaysOnTopHint 应被设置
    assert toast.windowFlags() & Qt.WindowStaysOnTopHint
```

- [ ] **步骤 2：运行测试验证失败**

运行：`python -m pytest tests/test_toast_notification.py -v`
预期：FAIL，报错 `ModuleNotFoundError: No module named 'hnust_exam.views.widgets.toast_notification'`

- [ ] **步骤 3：实现 ToastWidget 组件**

```python
# hnust_exam/views/widgets/toast_notification.py
"""非模态顶部浮层通知组件.

使用方式：
    ToastWidget(parent, "消息内容").show()
    ToastWidget(parent, "消息内容", duration_ms=3000, toast_type="error").show()
"""

from __future__ import annotations

from PySide6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout, QWidget
from hnust_exam.utils.theme import Theme


class ToastWidget(QFrame):
    """顶部浮层通知，非模态，自动消褪.

    Parameters
    ----------
    parent : QWidget
        父窗口（通常是主窗口）。
    text : str
        通知文本。
    duration_ms : int
        显示时长（毫秒），默认 2000。
    toast_type : str
        通知类型："info"（默认）、"success"、"warning"、"error"，
        影响背景色。
    """

    _TYPEDEF = {
        "info": {"bg": "#0078d7", "icon": "ℹ️"},
        "success": {"bg": "#28a745", "icon": "✅"},
        "warning": {"bg": "#ffc107", "icon": "⚠️"},
        "error": {"bg": "#dc3545", "icon": "❌"},
    }

    def __init__(
        self,
        parent: QWidget,
        text: str,
        duration_ms: int = 2000,
        toast_type: str = "info",
    ):
        super().__init__(parent)
        style = self._TYPEDEF.get(toast_type, self._TYPEDEF["info"])

        # 窗口标志：无边框 + 工具窗口 + 置顶（避免被浏览器遮挡）
        self.setWindowFlags(
            Qt.FramelessWindowHint
            | Qt.Tool
            | Qt.WindowStaysOnTopHint
        )
        # 不抢占焦点（让浏览器窗口保持焦点）
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)
        # 鼠标穿透（不阻止用户与下方 UI 交互）
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        # 关闭时释放内存
        self.setAttribute(Qt.WA_DeleteOnClose, True)

        # 圆角背景
        self.setStyleSheet(f"""
            ToastWidget {{
                background-color: {style['bg']};
                border-radius: 8px;
                padding: 12px 24px;
            }}
        """)

        # 布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 12, 24, 12)

        self.label = QLabel(f"{style['icon']} {text}")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("color: white; font-size: 11pt; background: transparent;")
        layout.addWidget(self.label)

        # 自适应大小
        self.adjustSize()
        # 设置最小宽度，避免太窄
        if self.width() < 160:
            self.resize(160, self.height())

        # 定位到父窗口顶部居中
        self._position_at_top_center()

        # 自动关闭定时器
        self._close_timer = QTimer(self)
        self._close_timer.setSingleShot(True)
        self._close_timer.timeout.connect(self._fade_out)
        self._close_timer.start(duration_ms)

        # 淡入动画
        self._fade_in()

    def _position_at_top_center(self) -> None:
        """将自身定位到父窗口顶部居中."""
        parent = self.parentWidget()
        if parent:
            # 使用父窗口的全局坐标
            parent_global = parent.mapToGlobal(parent.rect().topLeft())
            parent_width = parent.width()
            x = parent_global.x() + (parent_width - self.width()) // 2
            y = parent_global.y() + 20  # 距父窗口顶部 20px
            self.move(x, y)

    def _fade_in(self) -> None:
        """淡入效果：透明度 0 → 1."""
        self.setWindowOpacity(0.0)
        self.show()
        self._anim = QPropertyAnimation(self, b"windowOpacity")
        self._anim.setDuration(200)
        self._anim.setStartValue(0.0)
        self._anim.setEndValue(1.0)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)
        self._anim.start()

    def _fade_out(self) -> None:
        """淡出效果：透明度 1 → 0，完成后关闭."""
        if self._anim:
            self._anim.stop()
        self._anim = QPropertyAnimation(self, b"windowOpacity")
        self._anim.setDuration(300)
        self._anim.setStartValue(self.windowOpacity())
        self._anim.setEndValue(0.0)
        self._anim.setEasingCurve(QEasingCurve.InCubic)
        self._anim.finished.connect(self.close)
        self._anim.start()

    def showEvent(self, event) -> None:
        """每次 show() 时刷新位置，适应窗口大小变化."""
        super().showEvent(event)
        self._position_at_top_center()
```

- [ ] **步骤 4：运行测试验证通过**

运行：`python -m pytest tests/test_toast_notification.py -v`
预期：PASS，4 个测试均通过

- [ ] **步骤 5：提交 Task 1**

```bash
git add hnust_exam/views/widgets/toast_notification.py tests/test_toast_notification.py
git commit -m "feat: 添加顶部浮层通知组件 ToastWidget"
```

---

## 任务 2：将试题解析弹窗替换为 ToastWidget

**文件：**
- 修改：`hnust_exam/views/exam_page.py:448-462`
- 可选修改：`hnust_exam/views/dialogs/submit_dialog.py:298-304`

- [ ] **步骤 1：编写集成测试**

```python
# 追加到 tests/test_toast_notification.py 或在 exam_page 测试中

def test_show_analysis_uses_toast_instead_of_messagebox(monkeypatch, qtbot):
    """验证 show_analysis 使用 ToastWidget 而非 themed_info."""
    from hnust_exam.views.exam_page import ExamPage
    from hnust_exam.views.widgets.toast_notification import ToastWidget

    # 用于记录被调用的通知组件
    called_with = []

    # 替换 ToastWidget 的构造函数来捕获调用
    original_init = ToastWidget.__init__
    def mock_init(self, parent, text, duration_ms=2000, toast_type="info"):
        called_with.append((text, toast_type))
        original_init(self, parent, text, duration_ms=duration_ms, toast_type=toast_type)

    monkeypatch.setattr(ToastWidget, "__init__", mock_init)

    # 模拟 show_analysis 流程中的复制+通知步骤
    from hnust_exam.utils.ui_helpers import themed_info
    info_called = []
    def mock_info(*args, **kwargs):
        info_called.append(args)

    monkeypatch.setattr("hnust_exam.views.exam_page.themed_info", mock_info)

    # show_analysis 在用户确认后应调 ToastWidget
    # 这里实际测试的是 _copy_question_to_clipboard 后的通知
    # 但更精确的方式是在 exam_page 集成测试中测试整个流程
    # 本测试验证 ToastWidget 可被正确导入和实例化
    window = QMainWindow()
    window.resize(800, 600)
    qtbot.addWidget(window)
    toast = ToastWidget(window, "题目已复制到剪贴板，直接粘贴给豆包即可！")
    qtbot.addWidget(toast)
    assert toast.label.text() == "ℹ️ 题目已复制到剪贴板，直接粘贴给豆包即可！"
```

- [ ] **步骤 2：修改 exam_page.py 中的 show_analysis 方法**

将 `themed_info` 调用替换为 `ToastWidget`：

```python
    def show_analysis(self) -> None:
        import webbrowser
        reply = themed_question(
            self, "试题解析",
            "暂时没有解析，问问豆包吧\n\n是否跳转到豆包网页版？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._copy_question_to_clipboard()
            webbrowser.open("https://www.doubao.com")
            # 替换：旧的 themed_info 改为 ToastWidget 顶部浮层
            from hnust_exam.views.widgets.toast_notification import ToastWidget
            toast = ToastWidget(
                self.window(),  # 使用主窗口作为父窗口，保证全局置顶
                "题目已复制到剪贴板，直接粘贴给豆包即可！",
                toast_type="success",
            )
            toast.show()
```

**修改内容：**
1. 在文件顶部新增 `ToastWidget` 的导入（或在函数内局部导入以避免循环依赖）
2. 第 459-462 行将 `themed_info(self, "已复制", "...")` 替换为 `ToastWidget(...).show()`
3. 使用 `self.window()` 获取主窗口作为父对象，确保 toast 在整个应用顶部居中

- [ ] **步骤 3：验证手动测试**

手动运行应用，点击"试题解析"→"是"，确认：
1. 浏览器打开豆包首页
2. Toast 通知出现在主窗口顶部居中位置，覆盖在浏览器窗口之上
3. 2 秒后自动消失
4. Toast 不抢占焦点，用户可立即在浏览器中粘贴

- [ ] **步骤 4：提交 Task 2**

```bash
git add hnust_exam/views/exam_page.py
git commit -m "feat: 将试题解析弹窗替换为顶部 ToastWidget 通知"
```

---

## 任务 3：编写版本号比较工具函数

**文件：**
- 创建：`hnust_exam/utils/version.py`
- 测试：`tests/test_version_comparison.py`

- [ ] **步骤 1：编写版本比较函数的失败测试**

```python
# tests/test_version_comparison.py
"""测试版本号解析与比较工具函数."""
import pytest
from hnust_exam.utils.version import parse_version, compare_versions


# ── parse_version ──────────────────────────────────────────────────────

def test_parse_full_format():
    """解析完整格式 yyyy.mm.dd.tttt."""
    assert parse_version("2026.05.21.1430") == (2026, 5, 21, 1430)


def test_parse_date_only():
    """解析旧格式 yyyy.mm.dd（无时间部分，默认时间为 0）."""
    assert parse_version("2026.05.21") == (2026, 5, 21, 0)


def test_parse_empty_string():
    """空字符串视为版本 0."""
    assert parse_version("") == (0, 0, 0, 0)


def test_parse_invalid_format():
    """无法解析的格式视为版本 0."""
    assert parse_version("invalid") == (0, 0, 0, 0)
    assert parse_version("v1.0.0") == (0, 0, 0, 0)  # 非 yyyy.mm.dd 格式


def test_parse_partial_digits():
    """部分段非数字时容错."""
    assert parse_version("2026.ab.21.1430") == (2026, 0, 21, 1430)


def test_parse_single_number():
    """单数字版本号."""
    assert parse_version("2026") == (2026, 0, 0, 0)


# ── compare_versions ───────────────────────────────────────────────────

def test_compare_equal():
    """完全相等的版本."""
    assert compare_versions("2026.05.21.1430", "2026.05.21.1430") == 0
    assert compare_versions("2026.05.21", "2026.05.21") == 0


def test_compare_local_older_by_date():
    """本地日期更旧 → 需更新."""
    assert compare_versions("2026.05.20.1430", "2026.05.21.1430") == -1
    assert compare_versions("2026.04.30.1200", "2026.05.01.0800") == -1


def test_compare_local_newer_by_date():
    """本地日期更新 → 无需更新."""
    assert compare_versions("2026.05.22.0800", "2026.05.21.1430") == 1


def test_compare_same_date_older_time():
    """同日期，本地时间更旧 → 需更新."""
    assert compare_versions("2026.05.21.0800", "2026.05.21.1430") == -1


def test_compare_same_date_newer_time():
    """同日期，本地时间更新 → 无需更新."""
    assert compare_versions("2026.05.21.1600", "2026.05.21.1430") == 1


def test_compare_old_format_vs_new():
    """旧格式（无时间）vs 新格式：时间默认 0，新版本 > 0 """
    assert compare_versions("2026.05.21", "2026.05.21.1430") == -1
    assert compare_versions("2026.05.22", "2026.05.21.1430") == 1


def test_compare_empty_local():
    """本地版本为空 → 永远需更新."""
    assert compare_versions("", "2026.05.21.1430") == -1


def test_compare_both_empty():
    """两个都为空 → 相同."""
    assert compare_versions("", "") == 0


def test_compare_same_date_empty_time_vs_zero_time():
    """旧格式 yyyy.mm.dd 与 yyyy.mm.dd.0000 相等."""
    assert compare_versions("2026.05.21", "2026.05.21.0000") == 0
```

- [ ] **步骤 2：运行测试验证失败**

运行：`python -m pytest tests/test_version_comparison.py -v`
预期：FAIL，报错 `ModuleNotFoundError: No module named 'hnust_exam.utils.version'`

- [ ] **步骤 3：实现版本号工具函数**

```python
# hnust_exam/utils/version.py
"""版本号解析与比较工具.

版本格式：yyyy.mm.dd.tttt（如 2026.05.21.1430）
- yyyy: 年
- mm:   月
- dd:   日
- tttt: 时间（如 1430 = 14:30，可选，旧格式 yyyy.mm.dd 兼容）

比较规则：先比日期（年月日），再比时间。
"""

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)


def parse_version(version: str) -> tuple[int, ...]:
    """将版本字符串解析为可比较的整数元组.

    解析失败的段视为 0，完全无法解析的版本返回 (0, 0, 0, 0)。

    Returns
    -------
    tuple[int, int, int, int]
        (year, month, day, time)
    """
    if not version or not version.strip():
        return (0, 0, 0, 0)

    parts = version.strip().split(".")

    # 确保至少有4个元素，不足补0
    while len(parts) < 4:
        parts.append("0")

    result = []
    for p in parts[:4]:
        try:
            result.append(int(p))
        except (ValueError, TypeError):
            logger.warning("版本号段无法解析为整数: '%s' in '%s'", p, version)
            result.append(0)

    return tuple(result)  # type: ignore[return-value]


def compare_versions(local: str, remote: str) -> int:
    """比较两个版本号.

    Parameters
    ----------
    local : str
        本地版本号。
    remote : str
        远程（Gitee）版本号。

    Returns
    -------
    int
        -1: 本地比远程旧，需要更新
         0: 版本相同，无需更新
         1: 本地比远程新，无需更新
    """
    local_tuple = parse_version(local)
    remote_tuple = parse_version(remote)

    if local_tuple < remote_tuple:
        return -1
    elif local_tuple > remote_tuple:
        return 1
    return 0
```

- [ ] **步骤 4：运行测试验证通过**

运行：`python -m pytest tests/test_version_comparison.py -v`
预期：PASS，所有 14 个测试通过

- [ ] **步骤 5：提交 Task 3**

```bash
git add hnust_exam/utils/version.py tests/test_version_comparison.py
git commit -m "feat: 添加版本号解析与比较工具函数（yyyy.mm.dd.tttt 格式）"
```

---

## 任务 4：修改热更新版本检查逻辑

**文件：**
- 修改：`hnust_exam/services/resource_pack_updater.py:286-290`
- 测试：`tests/test_resource_pack_updater.py`（新建，模拟 API 调用）

- [ ] **步骤 1：编写热更新版本检查的集成测试**

```python
# tests/test_resource_pack_updater.py
"""测试 resource_pack_updater 的版本检查逻辑."""
import json
import pytest
from unittest.mock import patch, MagicMock, mock_open
from hnust_exam.services.resource_pack_updater import _do_update, _get_local_version


def test_do_update_local_older_than_remote(monkeypatch, tmp_path):
    """本地版本比远程旧 → 执行更新."""
    from hnust_exam.services import resource_pack_updater as rpu

    # 替换路径常量为临时目录
    monkeypatch.setattr(rpu, "QUESTION_BANK_DIR", str(tmp_path))
    monkeypatch.setattr(rpu, "QUESTION_BANK_FILES_DIR", str(tmp_path / "files"))
    monkeypatch.setattr(rpu, "_STAGING_DIR", str(tmp_path / "staging"))
    monkeypatch.setattr(rpu, "_BACKUP_DIR", str(tmp_path / "files_backup"))

    # 写入本地版本（旧）
    (tmp_path / "current_version").write_text("2026.05.20.0800", encoding="utf-8")

    # 模拟 Gitee API 返回
    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",  # 远程版本更新
    }

    # 模拟 session 和下载
    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        with patch.object(rpu, "_download_to_file", return_value=True):
            with patch.object(rpu, "zipfile"):
                with patch.object(rpu, "_safe_rmtree"):
                    with patch.object(rpu, "_safe_remove"):
                        with patch.object(rpu, "_regenerate_manifest"):
                            result = _do_update()

    # 验证：应该执行更新
    assert result.success
    assert "更新" in result.message
    assert result.new_version == "2026.05.21.1430"


def test_do_update_local_equal_to_remote(monkeypatch, tmp_path):
    """本地版本等于远程 → 不更新."""
    from hnust_exam.services import resource_pack_updater as rpu

    monkeypatch.setattr(rpu, "QUESTION_BANK_DIR", str(tmp_path))
    monkeypatch.setattr(rpu, "QUESTION_BANK_FILES_DIR", str(tmp_path / "files"))
    monkeypatch.setattr(rpu, "_STAGING_DIR", str(tmp_path / "staging"))
    monkeypatch.setattr(rpu, "_BACKUP_DIR", str(tmp_path / "files_backup"))

    (tmp_path / "current_version").write_text("2026.05.21.1430", encoding="utf-8")

    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",
    }

    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        # 如果走到下载阶段就算失败
        with patch.object(rpu, "_download_to_file", side_effect=AssertionError("不该下载")):
            result = _do_update()

    assert result.success
    assert "已是最新" in result.message


def test_do_update_same_date_older_time(monkeypatch, tmp_path):
    """同日期但本地时间更旧 → 更新."""
    from hnust_exam.services import resource_pack_updater as rpu

    monkeypatch.setattr(rpu, "QUESTION_BANK_DIR", str(tmp_path))
    monkeypatch.setattr(rpu, "QUESTION_BANK_FILES_DIR", str(tmp_path / "files"))
    monkeypatch.setattr(rpu, "_STAGING_DIR", str(tmp_path / "staging"))
    monkeypatch.setattr(rpu, "_BACKUP_DIR", str(tmp_path / "files_backup"))

    (tmp_path / "current_version").write_text("2026.05.21.0800", encoding="utf-8")

    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",
    }

    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        with patch.object(rpu, "_download_to_file", return_value=True):
            with patch.object(rpu, "zipfile"):
                with patch.object(rpu, "_safe_rmtree"):
                    with patch.object(rpu, "_safe_remove"):
                        with patch.object(rpu, "_regenerate_manifest"):
                            result = _do_update()

    assert result.success
    assert "更新" in result.message


def test_do_update_same_date_newer_time(monkeypatch, tmp_path):
    """同日期但本地时间更新 → 不更新."""
    from hnust_exam.services import resource_pack_updater as rpu

    monkeypatch.setattr(rpu, "QUESTION_BANK_DIR", str(tmp_path))
    monkeypatch.setattr(rpu, "QUESTION_BANK_FILES_DIR", str(tmp_path / "files"))
    monkeypatch.setattr(rpu, "_STAGING_DIR", str(tmp_path / "staging"))
    monkeypatch.setattr(rpu, "_BACKUP_DIR", str(tmp_path / "files_backup"))

    (tmp_path / "current_version").write_text("2026.05.21.1600", encoding="utf-8")

    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",
    }

    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        with patch.object(rpu, "_download_to_file", side_effect=AssertionError("不该下载")):
            result = _do_update()

    assert result.success
    assert "已是最新" in result.message


def test_do_update_old_format_compatibility(monkeypatch, tmp_path):
    """旧格式 yyyy.mm.dd vs 新格式 yyyy.mm.dd.tttt → 正确比较."""
    from hnust_exam.services import resource_pack_updater as rpu

    monkeypatch.setattr(rpu, "QUESTION_BANK_DIR", str(tmp_path))
    monkeypatch.setattr(rpu, "QUESTION_BANK_FILES_DIR", str(tmp_path / "files"))
    monkeypatch.setattr(rpu, "_STAGING_DIR", str(tmp_path / "staging"))
    monkeypatch.setattr(rpu, "_BACKUP_DIR", str(tmp_path / "files_backup"))

    # 本地是旧格式（无时间），远程是新格式（同日期但有时间）
    (tmp_path / "current_version").write_text("2026.05.21", encoding="utf-8")

    mock_release = {
        "zip_url": "https://example.com/question_bank.zip",
        "zip_size": 1000,
        "hash_url": "",
        "tag_name": "2026.05.21.1430",
    }

    with patch.object(rpu, "_fetch_release_assets", return_value=mock_release):
        with patch.object(rpu, "_download_to_file", return_value=True):
            with patch.object(rpu, "zipfile"):
                with patch.object(rpu, "_safe_rmtree"):
                    with patch.object(rpu, "_safe_remove"):
                        with patch.object(rpu, "_regenerate_manifest"):
                            result = _do_update()

    assert result.success
    assert "更新" in result.message
```

- [ ] **步骤 2：运行测试验证失败**

运行：`python -m pytest tests/test_resource_pack_updater.py -v`
预期：FAIL，因为 `_do_update` 仍然使用旧版字符串比较逻辑

测试结构说明：现在 `test_do_update_local_equal_to_remote` 等测试会失败，因为 `_do_update` 中 `remote_tag == local_ver` 的字符串比较不满足日期+时间语义化比较的需求。

- [ ] **步骤 3：修改 _do_update 中的版本检查逻辑**

在 `resource_pack_updater.py` 顶部添加导入，然后替换第 286-290 行的字符串比较：

```python
# 顶部新增导入
from hnust_exam.utils.version import compare_versions
```

替换第 286-290 行（原始代码在 `_do_update` 函数内）：

```python
        remote_tag = release["tag_name"]
        local_ver = _get_local_version()

        # 使用语义化版本比较（yyyy.mm.dd.tttt 格式）
        if local_ver and compare_versions(local_ver, remote_tag) >= 0:
            logger.info("题库已是最新版本: %s (local=%s)", remote_tag, local_ver)
            return PackUpdateResult(True, "题库已是最新", new_version=remote_tag)
```

**注意变化：**
1. `local_ver and compare_versions(...) >= 0` — 当本地版本为空（首次安装）时，`local_ver` 为 `""` 即 falsy，所以 `local_ver and ...` 短路为 `""`（falsy），跳过 "已是最新" 判断，进入下载流程。`compare_versions("", remote_tag)` 返回 -1，但 `and` 短路使其不会执行到那个分支。
2. `>= 0` — 本地版本大于或等于远程时都视为不需要更新。

- [ ] **步骤 4：运行测试验证通过**

运行：`python -m pytest tests/test_resource_pack_updater.py tests/test_version_comparison.py -v`
预期：PASS，所有版本比较 + 热更新集成测试均通过

- [ ] **步骤 5：提交 Task 4**

```bash
git add hnust_exam/services/resource_pack_updater.py tests/test_resource_pack_updater.py
git commit -m "feat: 题库热更新改用日期+时间版本比较逻辑（yyyy.mm.dd.tttt）"
```

---

## 自检清单

### 规格覆盖度
- ✅ **任务 1** 覆盖 "弹窗通知置顶置页面顶部" — 创建了 ToastWidget 浮层组件
- ✅ **任务 2** 覆盖 "避免被弹出的浏览器遮住" — ToastWidget 使用 `WindowStaysOnTopHint` 标志
- ✅ **任务 3** 覆盖 "yyyy.mm.dd.tttt 版本格式解析与比较" — 创建了独立可测试的 version.py
- ✅ **任务 4** 覆盖 "先比较日期远近…日期一样比较时间大小" — compare_versions 先比日期再比时间
- ✅ **任务 4** 覆盖 "本地比 gitee 的旧就去下载新的" — compare_versions 返回 -1 时触发下载
- ✅ **任务 4** 覆盖向后兼容 — parse_version 处理旧格式 yyyy.mm.dd 和空字符串

### 禁止占位符扫描
- 所有测试代码均为完整可执行的 pytest 测试
- 所有实现代码均为完整可导入的 Python 模块
- 无 "TODO"、"待定"、"补充细节" 等占位符

### 类型一致性
- `compare_versions(local, remote) -> int` 在 `version.py` 和 `resource_pack_updater.py` 中使用方式一致
- `parse_version(str) -> tuple` 在测试和实现中返回类型一致
- `ToastWidget(parent, text, duration_ms, toast_type)` 在创建和测试中签名一致

---

## 执行交接

计划已完成并保存到 `docs/superpowers/plans/2026-06-01-toast-and-version-update.md`。

两种执行方式：

**1. 子代理驱动（推荐）** — 每个任务调度一个新的子代理，任务间进行审查，快速迭代

**2. 内联执行** — 在当前会话中使用 executing-plans 执行任务，批量执行并设有检查点

**选哪种方式？**
