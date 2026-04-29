"""欢迎页：软件介绍、条款勾选、倒计时进入按钮."""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QCheckBox,
    QScrollArea,
    QFrame,
)

from hnust_exam.utils.constants import CURRENT_VERSION
from hnust_exam.utils.theme import Theme
from hnust_exam.utils.ui_helpers import themed_warning

if TYPE_CHECKING:
    from hnust_exam.views.main_window import MainWindow


class WelcomePage(QWidget):
    """欢迎页面."""

    def __init__(self, main_window: MainWindow, parent=None) -> None:
        super().__init__(parent)
        self.main_window = main_window
        self._countdown = 10
        self._build_ui()
        self._start_countdown()

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

        self.agree_check = QCheckBox("我已阅读并同意以上所有条款")
        self.agree_check.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 13pt; font-weight: bold;"
        )
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

    def _start_countdown(self) -> None:
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(1000)

    def _tick(self) -> None:
        self._countdown -= 1
        if self._countdown > 0:
            self.enter_btn.setText(f"进入系统（请等待 {self._countdown} 秒）")
        else:
            self._timer.stop()
            self.enter_btn.setText("进入系统")
            self.enter_btn.setEnabled(True)
            self.enter_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            c = Theme.get_current_colors()
            self.enter_btn.setStyleSheet(
                f"background-color: {c['PRIMARY']}; color: white; font-size: 14pt; "
                f"font-weight: bold; padding: 10px 40px; border: none; border-radius: 4px;"
            )

    def _on_enter(self) -> None:
        if not self.agree_check.isChecked():
            themed_warning(self, "提示", "请先阅读并同意以上条款")
            return
        self.main_window.show_select()
