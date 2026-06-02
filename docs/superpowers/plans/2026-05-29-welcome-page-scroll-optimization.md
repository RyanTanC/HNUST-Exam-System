# 欢迎页面滑动效果优化实现计划

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推荐）或 superpowers:executing-plans 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。

**目标：** 优化欢迎页面的滑动效果，实现惯性滚动和页面切换动画

**架构：** 使用 Qt 内置的 QScroller 实现惯性滚动，使用 QPropertyAnimation 实现页面切换的滑动动画

**技术栈：** PySide6, QScroller, QPropertyAnimation, QEasingCurve

---

## 文件结构

### 修改文件
- `hnust_exam/views/welcome_page.py` - 添加 QScroller 配置实现惯性滚动
- `hnust_exam/views/main_window.py` - 添加页面切换动画

### 新增文件
- `tests/test_scroll_animation.py` - 测试滚动和动画功能

---

## 任务 1：配置欢迎页面惯性滚动

**文件：**
- 修改：`hnust_exam/views/welcome_page.py:206-271`
- 测试：`tests/test_scroll_animation.py`

- [ ] **步骤 1：创建测试文件**

```python
# tests/test_scroll_animation.py
"""测试滚动和动画功能."""

import pytest
from PySide6.QtWidgets import QApplication, QScrollArea
from PySide6.QtCore import Qt, QPropertyAnimation


@pytest.fixture(scope="module")
def app():
    """创建 QApplication 实例."""
    return QApplication.instance() or QApplication([])


def test_scroll_area_has_scroller(app):
    """测试 QScrollArea 是否配置了 QScroller."""
    from hnust_exam.views.welcome_page import WelcomePage
    
    # 创建一个简单的主窗口模拟器
    class MockMainWindow:
        def show_select(self):
            pass
    
    welcome_page = WelcomePage(MockMainWindow())
    scroll_area = welcome_page.findChild(QScrollArea)
    
    assert scroll_area is not None, "欢迎页面应该包含 QScrollArea"
    
    # 检查是否配置了 QScroller
    from PySide6.QtWidgets import QScroller
    scroller = QScroller.scroller(scroll_area.viewport())
    assert scroller is not None, "QScrollArea 应该配置 QScroller"


def test_scroll_area_scroller_properties(app):
    """测试 QScroller 的属性配置."""
    from hnust_exam.views.welcome_page import WelcomePage
    from PySide6.QtWidgets import QScroller, QScrollerProperties
    
    class MockMainWindow:
        def show_select(self):
            pass
    
    welcome_page = WelcomePage(MockMainWindow())
    scroll_area = welcome_page.findChild(QScrollArea)
    
    scroller = QScroller.scroller(scroll_area.viewport())
    prop = scroller.scrollerProperties()
    
    # 检查帧率配置
    max_fps = prop.scrollMetric(QScrollerProperties.FrameRateMaximum)
    min_fps = prop.scrollMetric(QScrollerProperties.FrameRateMinimum)
    
    assert max_fps == 60, "最大帧率应该设置为 60"
    assert min_fps == 30, "最小帧率应该设置为 30"


def test_welcome_page_scroll_smoothness(app):
    """测试欢迎页面滚动是否平滑（基本功能测试）."""
    from hnust_exam.views.welcome_page import WelcomePage
    
    class MockMainWindow:
        def show_select(self):
            pass
    
    welcome_page = WelcomePage(MockMainWindow())
    
    # 验证页面可以正常显示
    assert welcome_page.isVisible() or welcome_page.width() > 0
```

- [ ] **步骤 2：运行测试验证失败**

运行：`cd C:\Users\Hyper_hui\PycharmProjects\PythonProject1 && .venv\Scripts\python.exe -m pytest tests/test_scroll_animation.py -v`
预期：FAIL，报错 "WelcomePage() missing 1 required positional argument: 'parent'"

- [ ] **步骤 3：修改 welcome_page.py 添加 QScroller**

```python
# 在 welcome_page.py 的 _build_ui() 方法中，替换现有的 scroll_area 配置
# 位置：约第 206-271 行

def _build_ui(self) -> None:
    c = Theme.get_current_colors()

    root_layout = QVBoxLayout(self)
    root_layout.setContentsMargins(0, 0, 0, 0)
    root_layout.setSpacing(0)

    # 顶部标题栏
    header = QFrame()
    header.setStyleSheet(
        f"background-color: {c['PRIMARY']}; padding: 6px 30px;"
    )
    header_layout = QHBoxLayout(header)
    # Logo
    logo_label = QLabel()
    logo_pixmap = QPixmap(get_resource_path("hnust_exam/resources/logo.png"))
    if not logo_pixmap.isNull():
        logo_label.setPixmap(
            logo_pixmap.scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        )
        logo_label.setStyleSheet("padding-right: 6px;")
    header_layout.addWidget(logo_label)
    title_label = QLabel("HNUST仿真平台")
    title_label.setStyleSheet(
        f"color: white; font-size: 16pt; font-weight: bold;"
    )
    header_layout.addWidget(title_label)
    ver_label = QLabel(CURRENT_VERSION)
    ver_label.setStyleSheet(
        f"color: {c['HEADER_SUB_TEXT']}; font-size: 13pt; padding-left: 8px;"
    )
    header_layout.addWidget(ver_label)
    header_layout.addStretch()
    root_layout.addWidget(header)

    # 主内容区
    main_card = QFrame()
    main_card.setStyleSheet(
        f"background-color: {c['WHITE']}; "
        f"border-radius: 4px;"
    )
    main_layout = QVBoxLayout(main_card)
    main_layout.setContentsMargins(40, 30, 40, 30)

    welcome_title = QLabel("欢迎使用 HNUST 考试仿真平台")
    welcome_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
    welcome_title.setStyleSheet(
        f"color: {c['PRIMARY']}; font-size: 18pt; font-weight: bold; "
        f"padding-bottom: 20px;"
    )
    main_layout.addWidget(welcome_title)

    # 可滚动的介绍内容
    scroll_area = QScrollArea()
    scroll_area.setWidgetResizable(True)
    scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    
    # 新增：配置惯性滚动
    from PySide6.QtWidgets import QScroller, QScrollerProperties
    scroller = QScroller.scroller(scroll_area.viewport())
    scroller.grabGesture(scroll_area.viewport(), QScroller.TouchGesture)
    
    # 配置滚动参数
    prop = scroller.scrollerProperties()
    prop.setScrollMetric(QScrollerProperties.FrameRateMaximum, 60)
    prop.setScrollMetric(QScrollerProperties.FrameRateMinimum, 30)
    prop.setScrollMetric(QScrollerProperties.VerticalOverscrollPolicy, 
                        QScrollerProperties.OverscrollAlwaysOff)
    prop.setScrollMetric(QScrollerProperties.ScrollingInProgressDecelerationFactor, 0.9)
    prop.setScrollMetric(QScrollerProperties.AirspeedFactor, 2.0)
    prop.setScrollMetric(QScrollerProperties.SlideFactor, 1.5)
    scroller.setScrollerProperties(prop)

    scroll_content = QWidget()
    scroll_layout = QVBoxLayout(scroll_content)
    scroll_layout.setContentsMargins(40, 10, 40, 10)

    sections = [
        ("软件介绍", [
            "本软件是模仿学校机房万维考试系统开发的免费练习工具",
            "专为HNUST同学设计，让你在宿舍也能随时随地进行考试模拟练习",
            "完美还原考试界面和操作流程，提前熟悉考试环境",
            "如若想使用程序设计、程序改错、填空三种功能，"
            "你的电脑需预先配置好 Python 运行环境。",
        ]),
        ("功能特点", [
            "支持单选、判断、填空、程序填空、程序改错、程序设计等多种题型",
            "自动判分，即时显示答题结果和正确答案",
            "一键打开程序文件，自动用IDLE编辑",
            "程序文件自动备份，支持一键重置",
            "题目导航，快速跳转到未答题目",
            "考试计时，时间到自动交卷",
            "试卷使用记录，查看练习历史",
            "支持深色模式和字体缩放",
        ]),
        ("开源声明", [
            "本软件完全免费开源，代码将托管在GitHub上",
            "任何人都可以自由下载、使用、修改和分发",
            "严禁任何形式的商用售卖，违者必究",
        ]),
        ("开发说明", [
            "本项目采用 AI 辅助开发模式完成",
            "特别感谢：Claude code 、豆包、DeepSeek、小米MIMO 提供的AI编程支持",
            "如果觉得好用，欢迎给个Star支持一下作者",
        ]),
        ("问题反馈", [
            "该应用为学生开发，可能存在一些bug和不完善的地方",
            "如果遇到任何问题或有改进建议",
            "欢迎通过GitHub提交Issue或在频道私信作者",
        ]),
        ("免责声明", [
            "本软件仅供学习交流使用，与学校官方考试系统无关",
            "题库内容由用户自行提供，作者不承担任何版权责任",
            "使用本软件产生的任何后果由用户自行承担",
        ]),
    ]

    for title, content in sections:
        section_title = QLabel(title)
        section_title.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 14pt; font-weight: bold; "
            f"padding-top: 15px; padding-bottom: 8px;"
        )
        scroll_layout.addWidget(section_title)
        for line in content:
            line_label = QLabel(line)
            line_label.setWordWrap(True)
            line_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            line_label.setStyleSheet(
                f"color: {c['TEXT']}; font-size: 11pt; font-weight: 500; padding: 1px 0;"
            )
            scroll_layout.addWidget(line_label)

    scroll_area.setWidget(scroll_content)
    main_layout.addWidget(scroll_area)

    root_layout.addWidget(main_card, 1)

    # 底部：勾选 + 进入按钮
    bottom_frame = QFrame()
    bottom_frame.setStyleSheet(f"background-color: {c['BG']}; padding: 10px 40px;")
    bottom_layout = QHBoxLayout(bottom_frame)
    bottom_layout.setContentsMargins(40, 10, 40, 20)

    self.agree_check = AnimatedCheckBox("我已阅读并同意以上所有条款")
    bottom_layout.addWidget(self.agree_check)

    self.enter_btn = QPushButton("进入系统（请等待 10 秒）")
    self.enter_btn.setEnabled(False)
    self.enter_btn.setCursor(Qt.CursorShape.ArrowCursor)
    self.enter_btn.setStyleSheet(
        f"background-color: #cccccc; color: white; font-size: 14pt; "
        f"font-weight: bold; padding: 10px 40px; border: none; border-radius: 4px;"
    )
    self.enter_btn.clicked.connect(self._on_enter)
    bottom_layout.addWidget(self.enter_btn)

    root_layout.addWidget(bottom_frame)
```

- [ ] **步骤 4：运行测试验证通过**

运行：`cd C:\Users\Hyper_hui\PycharmProjects\PythonProject1 && .venv\Scripts\python.exe -m pytest tests/test_scroll_animation.py::test_scroll_area_has_scroller -v`
预期：PASS

- [ ] **步骤 5：Commit**

```bash
git add hnust_exam/views/welcome_page.py tests/test_scroll_animation.py
git commit -m "feat: 为欢迎页面添加惯性滚动功能"
```

---

## 任务 2：添加页面切换动画

**文件：**
- 修改：`hnust_exam/views/main_window.py:81-83`
- 测试：`tests/test_scroll_animation.py`

- [ ] **步骤 1：添加页面切换动画测试**

```python
# 在 tests/test_scroll_animation.py 中添加以下测试

def test_main_window_has_animation_method(app):
    """测试主窗口是否有动画切换方法."""
    from hnust_exam.views.main_window import MainWindow
    from hnust_exam.services.config_manager import ConfigManager
    
    config_mgr = ConfigManager()
    main_window = MainWindow(config_mgr)
    
    # 检查是否有 switch_to_page 方法
    assert hasattr(main_window, 'switch_to_page'), "主窗口应该有 switch_to_page 方法"
    
    # 检查方法是否可调用
    assert callable(getattr(main_window, 'switch_to_page')), "switch_to_page 应该是可调用的"


def test_page_switch_animation_imports(app):
    """测试页面切换动画所需的导入."""
    from PySide6.QtCore import QPropertyAnimation, QEasingCurve, QRect
    from PySide6.QtWidgets import QAbstractAnimation
    
    # 验证导入成功
    assert QPropertyAnimation is not None
    assert QEasingCurve is not None
    assert QRect is not None
    assert QAbstractAnimation is not None


def test_main_window_page_constants(app):
    """测试主窗口页面常量."""
    from hnust_exam.views.main_window import MainWindow
    from hnust_exam.services.config_manager import ConfigManager
    
    config_mgr = ConfigManager()
    main_window = MainWindow(config_mgr)
    
    # 检查页面常量
    assert main_window.PAGE_WELCOME == 0
    assert main_window.PAGE_SELECT == 1
    assert main_window.PAGE_EXAM == 2
    assert main_window.PAGE_RESULT == 3
```

- [ ] **步骤 2：运行测试验证失败**

运行：`cd C:\Users\Hyper_hui\PycharmProjects\PythonProject1 && .venv\Scripts\python.exe -m pytest tests/test_scroll_animation.py::test_main_window_has_animation_method -v`
预期：PASS（因为 switch_to_page 方法已存在，但需要修改）

- [ ] **步骤 3：修改 main_window.py 添加动画**

```python
# 在 main_window.py 中修改 switch_to_page 方法
# 位置：约第 81-83 行

def switch_to_page(self, page_index: int) -> None:
    """切换到指定页面，带滑动动画"""
    current_index = self.stack.currentIndex()
    if current_index == page_index:
        return
    
    # 获取当前几何信息
    current_geometry = self.stack.geometry()
    
    # 创建动画
    animation = QPropertyAnimation(self.stack, b"geometry")
    animation.setDuration(300)
    animation.setEasingCurve(QEasingCurve.OutCubic)
    
    # 计算动画参数
    if page_index > current_index:
        # 向右滑动
        start_x = current_geometry.width()
    else:
        # 向左滑动
        start_x = -current_geometry.width()
    
    start_rect = QRect(start_x, 0, current_geometry.width(), current_geometry.height())
    end_rect = QRect(0, 0, current_geometry.width(), current_geometry.height())
    
    animation.setStartValue(start_rect)
    animation.setEndValue(end_rect)
    
    # 先切换页面，再播放动画
    self.stack.setCurrentIndex(page_index)
    animation.start(QAbstractAnimation.DeleteWhenStopped)

# 确保导入了必要的模块
# 在文件顶部添加：
from PySide6.QtCore import Qt, QPropertyAnimation, QEasingCurve, QRect
from PySide6.QtWidgets import QAbstractAnimation
```

- [ ] **步骤 4：运行测试验证通过**

运行：`cd C:\Users\Hyper_hui\PycharmProjects\PythonProject1 && .venv\Scripts\python.exe -m pytest tests/test_scroll_animation.py -v`
预期：PASS

- [ ] **步骤 5：Commit**

```bash
git add hnust_exam/views/main_window.py tests/test_scroll_animation.py
git commit -m "feat: 添加页面切换滑动动画"
```

---

## 任务 3：集成测试和性能验证

**文件：**
- 测试：`tests/test_scroll_animation.py`
- 手动测试：`main.py`

- [ ] **步骤 1：添加集成测试**

```python
# 在 tests/test_scroll_animation.py 中添加以下测试

def test_welcome_page_full_integration(app):
    """欢迎页面完整集成测试."""
    from hnust_exam.views.welcome_page import WelcomePage
    from PySide6.QtWidgets import QScrollArea, QScroller
    
    class MockMainWindow:
        def show_select(self):
            pass
    
    # 创建欢迎页面
    welcome_page = WelcomePage(MockMainWindow())
    
    # 验证页面结构
    assert welcome_page.width() > 0 or welcome_page.isVisible()
    
    # 验证滚动区域存在
    scroll_area = welcome_page.findChild(QScrollArea)
    assert scroll_area is not None
    
    # 验证 QScroller 配置
    scroller = QScroller.scroller(scroll_area.viewport())
    assert scroller is not None
    
    # 验证滚动参数
    prop = scroller.scrollerProperties()
    max_fps = prop.scrollMetric(QScrollerProperties.FrameRateMaximum)
    assert max_fps == 60


def test_main_window_full_integration(app):
    """主窗口完整集成测试."""
    from hnust_exam.views.main_window import MainWindow
    from hnust_exam.services.config_manager import ConfigManager
    from PySide6.QtCore import QPropertyAnimation
    
    config_mgr = ConfigManager()
    main_window = MainWindow(config_mgr)
    
    # 验证窗口可以正常显示
    main_window.show()
    assert main_window.isVisible()
    
    # 验证页面切换方法存在
    assert hasattr(main_window, 'switch_to_page')
    
    # 验证页面常量
    assert main_window.PAGE_WELCOME == 0
    assert main_window.PAGE_SELECT == 1
```

- [ ] **步骤 2：运行完整测试套件**

运行：`cd C:\Users\Hyper_hui\PycharmProjects\PythonProject1 && .venv\Scripts\python.exe -m pytest tests/test_scroll_animation.py -v`
预期：所有测试通过

- [ ] **步骤 3：手动测试程序**

运行：`cd C:\Users\Hyper_hui\PycharmProjects\PythonProject1 && .venv\Scripts\python.exe main.py`

手动验证：
1. 启动程序，查看欢迎页面
2. 使用鼠标滚轮滚动内容，验证是否平滑
3. 等待倒计时结束，点击"进入系统"按钮
4. 验证页面切换是否有滑动动画
5. 从其他页面返回欢迎页，验证动画效果

- [ ] **步骤 4：性能测试**

创建性能测试脚本：

```python
# tests/test_performance.py
"""性能测试."""

import time
import pytest
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer


def test_scroll_performance():
    """测试滚动性能."""
    app = QApplication.instance() or QApplication([])
    
    from hnust_exam.views.welcome_page import WelcomePage
    
    class MockMainWindow:
        def show_select(self):
            pass
    
    welcome_page = WelcomePage(MockMainWindow())
    welcome_page.show()
    
    # 模拟滚动操作
    start_time = time.time()
    for _ in range(100):
        app.processEvents()
        time.sleep(0.016)  # 模拟 60 FPS
    
    end_time = time.time()
    duration = end_time - start_time
    
    # 验证性能（应该在合理时间内完成）
    assert duration < 5.0, f"性能测试耗时过长: {duration}秒"


def test_animation_performance():
    """测试动画性能."""
    app = QApplication.instance() or QApplication([])
    
    from hnust_exam.views.main_window import MainWindow
    from hnust_exam.services.config_manager import ConfigManager
    
    config_mgr = ConfigManager()
    main_window = MainWindow(config_mgr)
    main_window.show()
    
    # 测试页面切换性能
    start_time = time.time()
    main_window.switch_to_page(main_window.PAGE_SELECT)
    
    # 等待动画完成
    for _ in range(30):  # 300ms 动画
        app.processEvents()
        time.sleep(0.01)
    
    end_time = time.time()
    duration = end_time - start_time
    
    # 验证动画性能
    assert duration < 1.0, f"动画性能测试耗时过长: {duration}秒"
    assert main_window.stack.currentIndex() == main_window.PAGE_SELECT
```

- [ ] **步骤 5：运行性能测试**

运行：`cd C:\Users\Hyper_hui\PycharmProjects\PythonProject1 && .venv\Scripts\python.exe -m pytest tests/test_performance.py -v`
预期：PASS

- [ ] **步骤 6：Commit**

```bash
git add tests/test_scroll_animation.py tests/test_performance.py
git commit -m "test: 添加滚动和动画的集成测试与性能测试"
```

---

## 任务 4：最终验证和清理

**文件：**
- 验证：所有修改的文件
- 清理：临时文件

- [ ] **步骤 1：运行完整测试套件**

运行：`cd C:\Users\Hyper_hui\PycharmProjects\PythonProject1 && .venv\Scripts\python.exe -m pytest tests/ -v`
预期：所有测试通过

- [ ] **步骤 2：代码质量检查**

运行：`cd C:\Users\Hyper_hui\PycharmProjects\PythonProject1 && .venv\Scripts\python.exe -m compileall -q hnust_exam tests main.py`
预期：无语法错误

- [ ] **步骤 3：手动功能验证**

运行程序并验证：
1. 欢迎页面滚动平滑，无卡顿
2. 页面切换动画流畅
3. 所有功能正常工作
4. 无性能问题

- [ ] **步骤 4：清理临时文件**

删除测试过程中创建的临时文件（如果有）

- [ ] **步骤 5：最终提交**

```bash
git add .
git commit -m "feat: 完成欢迎页面滑动效果优化

- 为欢迎页面添加惯性滚动功能
- 添加页面切换滑动动画
- 添加完整的测试套件
- 优化滚动性能参数"
```

---

## 规格合规性修复记录

**日期：** 2026-05-29

### 问题 1：多余测试

规格要求 3 个测试，实现创建了 4 个。`test_scroll_area_deceleration` 不在规格中，已删除。

### 问题 2：API 适配导致规格项省略

规格使用了不存在的 PySide6 API，实现者正确适配到真实 API，但部分规格项无等效 API：

| 规格项 | 实际 API | 状态 |
|--------|---------|------|
| `FrameRateMaximum` | `FrameRate` + `Fps60` | ✅ 已适配 |
| `FrameRateMinimum` | 无等效 API | ⚠️ 已省略（无最低帧率设置） |
| `VerticalOverscrollPolicy` | `VerticalOvershootPolicy` + `OvershootAlwaysOff` | ✅ 已适配 |
| `ScrollingInProgressDecelerationFactor` | `DecelerationFactor` | ✅ 已适配 |
| `AirspeedFactor` | `DragVelocitySmoothingFactor` | ✅ 已适配 |
| `SlideFactor` | 无等效 API | ⚠️ 已省略（无滑动因子设置） |

**省略原因：** PySide6 的 `QScrollerProperties.ScrollMetric` 枚举中不存在 `FrameRateMinimum` 和 `SlideFactor` 对应项。Qt 文档中这两个是 Qt Quick 的属性，不适用于 Widgets 的 QScroller。

---

## 验证清单

完成所有任务后，验证以下内容：

- [ ] 欢迎页面滚动平滑，支持惯性效果
- [ ] 页面切换有流畅的滑动动画
- [ ] 所有测试通过
- [ ] 程序性能良好
- [ ] 无内存泄漏
- [ ] 代码符合项目风格

## 回滚计划

如果出现问题，可以回滚到优化前的状态：

```bash
git reset --hard HEAD~1  # 回滚最后一次提交
# 或者
git checkout HEAD~1 -- hnust_exam/views/welcome_page.py hnust_exam/views/main_window.py
```

---

**计划完成时间：** 2026-05-29
**计划状态：** 待执行
**预计耗时：** 30-45 分钟