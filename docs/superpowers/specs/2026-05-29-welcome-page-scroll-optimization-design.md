# 欢迎页面滑动效果优化设计

## 问题描述

当前欢迎页面的滑动效果一顿一顿的，用户体验不佳。用户希望实现：
1. **惯性滚动**：像手机触摸屏一样，手指松开后内容继续滑动并逐渐停止
2. **页面切换动画**：从欢迎页到选卷页等页面切换时有流畅的过渡动画

## 解决方案

采用 Qt 内置功能实现，改动最小且效果明显：

### 1. 欢迎页面滚动优化

**使用 `QScroller` 实现惯性滚动：**

```python
# 在 welcome_page.py 中
scroll_area = QScrollArea()
scroller = QScroller.scroller(scroll_area.viewport())
scroller.grabGesture(scroll_area.viewport(), QScroller.TouchGesture)

# 配置滚动参数
prop = scroller.scrollerProperties()
prop.setScrollMetric(QScrollerProperties.FrameRateMaximum, 60)  # 60 FPS
prop.setScrollMetric(QScrollerProperties.FrameRateMinimum, 30)  # 最低 30 FPS
prop.setScrollMetric(QScrollerProperties.VerticalOverscrollPolicy, QScrollerProperties.OverscrollAlwaysOff)
scroller.setScrollerProperties(prop)
```

**优点：**
- 使用 Qt 内置功能，无需额外依赖
- 支持触摸和鼠标滚轮
- 自带惯性和平滑滚动效果

### 2. 页面切换动画

**使用 `QPropertyAnimation` 实现滑动动画：**

```python
# 在 main_window.py 中
def switch_to_page(self, page_index: int) -> None:
    """切换到指定页面，带滑动动画"""
    current_index = self.stack.currentIndex()
    if current_index == page_index:
        return
    
    # 创建动画
    animation = QPropertyAnimation(self.stack, b"geometry")
    animation.setDuration(300)  # 300ms 过渡时间
    animation.setEasingCurve(QEasingCurve.OutCubic)  # 缓出曲线
    
    # 计算起始和结束位置
    start_geometry = self.stack.geometry()
    end_geometry = start_geometry
    
    if page_index > current_index:
        # 向右滑动：新页面从右边进入
        start_geometry.moveLeft(start_geometry.width())
    else:
        # 向左滑动：新页面从左边进入
        start_geometry.moveLeft(-start_geometry.width())
    
    animation.setStartValue(start_geometry)
    animation.setEndValue(end_geometry)
    
    # 先切换页面，再播放动画
    self.stack.setCurrentIndex(page_index)
    animation.start(QAbstractAnimation.DeleteWhenStopped)
```

**优点：**
- 实现简单，代码改动小
- 支持多种缓动曲线
- 性能良好

## 技术实现细节

### 文件修改列表

1. **`hnust_exam/views/welcome_page.py`**
   - 在 `_build_ui()` 方法中配置 `QScroller`
   - 添加滚动参数配置

2. **`hnust_exam/views/main_window.py`**
   - 修改 `switch_to_page()` 方法，添加动画
   - 导入必要的 Qt 模块

### 具体代码改动

#### welcome_page.py 改动

```python
# 在 _build_ui() 方法中，替换现有的 scroll_area 配置
scroll_area = QScrollArea()
scroll_area.setWidgetResizable(True)
scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

# 新增：配置惯性滚动
from PySide6.QtWidgets import QScroller
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
```

#### main_window.py 改动

```python
# 导入添加
from PySide6.QtCore import QPropertyAnimation, QEasingCurve, QRect

# 修改 switch_to_page 方法
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
```

## 测试计划

### 功能测试

1. **滚动测试**
   - 验证欢迎页面滚动是否平滑
   - 测试鼠标滚轮滚动
   - 测试触摸板滚动（如果可用）
   - 验证惯性效果是否正常

2. **页面切换测试**
   - 从欢迎页切换到选卷页
   - 从选卷页切换到考试页
   - 从考试页切换到成绩页
   - 验证动画是否流畅

### 性能测试

1. **帧率测试**
   - 使用 Qt 性能工具监控动画帧率
   - 确保动画保持 60 FPS

2. **内存测试**
   - 监控动画过程中的内存使用
   - 确保没有内存泄漏

### 兼容性测试

1. **不同 Windows 版本**
   - Windows 10
   - Windows 11

2. **不同分辨率**
   - 1920x1080
   - 1366x768
   - 4K 分辨率

## 风险评估

### 低风险
- 使用 Qt 内置功能，稳定性高
- 改动范围小，不影响其他功能

### 中等风险
- 不同硬件配置可能影响动画性能
- 需要测试在不同配置下的表现

## 实施步骤

1. **第一步：备份当前代码**
   - 提交当前更改到 git

2. **第二步：修改欢迎页面**
   - 在 `welcome_page.py` 中添加 QScroller 配置

3. **第三步：修改主窗口**
   - 在 `main_window.py` 中添加页面切换动画

4. **第四步：测试验证**
   - 运行程序测试滚动效果
   - 测试页面切换动画

5. **第五步：性能优化**
   - 根据测试结果调整参数

## 成功标准

1. 欢迎页面滚动平滑，无卡顿
2. 页面切换动画流畅，无延迟
3. 动画帧率保持在 30 FPS 以上
4. 不影响程序其他功能
5. 在不同硬件配置下表现一致

## 后续优化

如果效果不理想，可以考虑：
1. 自定义弹簧物理引擎（方案二）
2. 添加更多动画效果（如淡入淡出）
3. 优化滚动性能（如减少重绘区域）

---

**设计完成时间：** 2026-05-29
**设计者：** AI 助手
**状态：** 待用户审查