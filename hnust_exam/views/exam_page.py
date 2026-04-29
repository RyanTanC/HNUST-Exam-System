"""考试页：题目展示 + 导航 + 底部按钮 + 计时器."""

from __future__ import annotations

import os
from datetime import datetime
from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFrame,
    QProgressBar,
    QScrollArea,
    QMessageBox,
)

from hnust_exam.models.exam import Exam
from hnust_exam.models.result import Result
from hnust_exam.services.backup_manager import BackupManager
from hnust_exam.utils.constants import EXAM_TIME_SECONDS
from hnust_exam.utils.theme import Theme
from hnust_exam.utils.ui_helpers import themed_question, themed_info, themed_critical

if TYPE_CHECKING:
    from hnust_exam.views.main_window import MainWindow


class ExamPage(QWidget):
    """考试主页面."""

    def __init__(self, main_window: MainWindow, parent=None) -> None:
        super().__init__(parent)
        self.main_window = main_window
        self.exam: Exam | None = None
        self.exam_file_path: str = ""
        self.remaining_time: int = EXAM_TIME_SECONDS
        self.timer_running: bool = False
        self.exam_submitted: bool = False
        self.show_answer_immediately: bool = False
        self.backup_mgr = BackupManager()

        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)

        # ──────────────────────────────────────────────
        # ★ 修复：让 ExamPage 能接收键盘焦点
        # ──────────────────────────────────────────────
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        self._build_ui()

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # 顶部信息栏
        self._top_bar = QFrame()
        self._top_bar.setStyleSheet(
            f"background-color: {c['PRIMARY']}; padding: 8px 15px;"
        )
        top_layout = QHBoxLayout(self._top_bar)

        left_info = QLabel("HNUST仿真平台")
        left_info.setStyleSheet("color: white; font-size: 12pt; font-weight: bold;")
        top_layout.addWidget(left_info)
        top_layout.addStretch()

        self.exam_name_label = QLabel("")
        self.exam_name_label.setStyleSheet(f"color: white; font-size: 10pt;")
        self.exam_name_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        top_layout.addWidget(self.exam_name_label)

        self.student_info = QLabel("姓名：xxx  学号：xxxxxxxxxxx")
        self.student_info.setStyleSheet(f"color: white; font-size: 10pt;")
        self.student_info.setAlignment(Qt.AlignmentFlag.AlignRight)
        top_layout.addWidget(self.student_info)

        root_layout.addWidget(self._top_bar)

        # 进度条
        progress_bar = QProgressBar()
        progress_bar.setFixedHeight(6)
        progress_bar.setTextVisible(False)
        progress_bar.setValue(0)
        root_layout.addWidget(progress_bar)
        self._progress_bar = progress_bar

        self.progress_label = QLabel("已完成 0 / 0 题")
        self.progress_label.setStyleSheet(f"color: {c['MUTED']}; font-size: 9pt; padding: 2px 10px;")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        # 将进度标签和条放在一起
        self._progress_container = QFrame()
        self._progress_container.setStyleSheet(f"background-color: {c['BG']};")
        pc_layout = QHBoxLayout(self._progress_container)
        pc_layout.setContentsMargins(10, 2, 10, 0)
        pc_layout.addWidget(progress_bar, 1)
        pc_layout.addWidget(self.progress_label)
        root_layout.addWidget(self._progress_container)

        # 主体：左题 + 右导航
        body = QHBoxLayout()
        body.setContentsMargins(10, 10, 10, 10)

        # 左侧题目区
        from hnust_exam.views.question_widget import QuestionWidget
        self.question_widget = QuestionWidget(self)
        body.addWidget(self.question_widget, 1)

        # 右侧导航
        from hnust_exam.views.nav_panel import NavPanel
        self.nav_panel = NavPanel(self)
        body.addWidget(self.nav_panel, 0)

        root_layout.addLayout(body, 1)

        # 底部按钮栏
        self._bottom_bar = QFrame()
        self._bottom_bar.setFixedHeight(60)
        self._bottom_bar.setStyleSheet(f"background-color: {c['BG']}; padding: 5px 10px;")
        bottom_layout = QHBoxLayout(self._bottom_bar)

        btn_configs = [
            ("上题", self.prev_question),
            ("下题", self.next_question),
            ("下一未答", self.jump_next_unanswered),
            ("答题", self.open_program_file),
            ("试题文件夹", self.open_exam_folder),
            ("重做", self.redo_question),
            ("标记试题", self.toggle_mark),
            ("答案", self.show_answer),
            ("试题解析", self.show_analysis),
        ]

        self._buttons = {}
        for text, cmd in btn_configs:
            btn = QPushButton(text)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            btn.clicked.connect(cmd)
            bottom_layout.addWidget(btn)
            self._buttons[text] = btn

        # 隐藏"答题"按钮（仅程序题显示）
        self._buttons["答题"].setVisible(False)

        bottom_layout.addStretch()

        # 倒计时
        self.time_label = QLabel("01:00:00")
        self.time_label.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 14pt; font-weight: bold; padding: 0 20px;"
        )
        bottom_layout.addWidget(self.time_label)

        # 交卷按钮
        self._submit_btn = QPushButton("交卷")
        self._submit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._submit_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._submit_btn.setStyleSheet(
            f"background-color: {c['DANGER']}; color: white; font-size: 12pt; "
            f"font-weight: bold; padding: 8px 25px; border: none; border-radius: 4px;"
        )
        self._submit_btn.clicked.connect(self.submit_exam)
        bottom_layout.addWidget(self._submit_btn)

        # 底部提示
        self._watermark = QLabel("免费使用 禁止售卖")
        self._watermark.setStyleSheet(f"color: {c['MUTED']}; font-size: 7pt; padding-left: 10px;")
        bottom_layout.addWidget(self._watermark)

        root_layout.addWidget(self._bottom_bar)

    def refresh_theme(self) -> None:
        """刷新主题颜色."""
        c = Theme.get_current_colors()
        # 顶部栏
        self._top_bar.setStyleSheet(
            f"background-color: {c['PRIMARY']}; padding: 8px 15px;"
        )
        # 进度区域
        self._progress_container.setStyleSheet(f"background-color: {c['BG']};")
        self._update_progress_label_style(c)
        # 底部栏
        self._bottom_bar.setStyleSheet(f"background-color: {c['BG']}; padding: 5px 10px;")
        self._submit_btn.setStyleSheet(
            f"background-color: {c['DANGER']}; color: white; font-size: 12pt; "
            f"font-weight: bold; padding: 8px 25px; border: none; border-radius: 4px;"
        )
        self._watermark.setStyleSheet(f"color: {c['MUTED']}; font-size: 7pt; padding-left: 10px;")
        self._update_time_display()
        # 刷新子组件
        self.question_widget.refresh_theme()
        self.nav_panel.refresh_theme()

    def _update_progress_label_style(self, c: dict | None = None) -> None:
        """更新进度标签样式."""
        if c is None:
            c = Theme.get_current_colors()
        if self.exam and self.exam.answered_count == self.exam.total_count and self.exam.total_count > 0:
            self.progress_label.setStyleSheet(f"color: {c['SUCCESS']}; font-size: 9pt; padding: 2px 10px;")
        else:
            self.progress_label.setStyleSheet(f"color: {c['MUTED']}; font-size: 9pt; padding: 2px 10px;")

    def setup_exam(self, exam: Exam, file_path: str) -> None:
        """设置考试数据并初始化."""
        self.exam = exam
        self.exam_file_path = file_path
        self.remaining_time = EXAM_TIME_SECONDS
        self.exam_submitted = False

        # 加载设置
        cfg = self.main_window.config_mgr.load_config()
        self.show_answer_immediately = cfg.get("show_answer_immediately", False)

        # 检查 Python 环境（程序题需要）
        from hnust_exam.utils.constants import PROGRAM_TYPES
        has_program = any(q.q_type in PROGRAM_TYPES for q in exam.questions)
        if has_program:
            python_path = cfg.get("user_python_path", "")
            from hnust_exam.services.python_env import find_system_python
            if not python_path or not os.path.isfile(python_path):
                found = find_system_python()
                if not found:
                    from hnust_exam.views.dialogs.python_env_dialog import PythonEnvDialog
                    dlg = PythonEnvDialog(self)
                    dlg.exec()
                    if dlg.python_path:
                        cfg["user_python_path"] = dlg.python_path
                        self.main_window.config_mgr.save_config(cfg)

        # 初始化备份
        self.backup_mgr.init_backup(file_path, exam)

        # 记录进度
        self._save_exam_progress("started")

        # 更新UI
        exam_name = os.path.splitext(os.path.basename(file_path))[0]
        self.exam_name_label.setText(f"{exam_name} · 练习")

        student_name = cfg.get("student_name", "")
        student_id = cfg.get("student_id", "")
        if student_name or student_id:
            self.student_info.setText(f"姓名：{student_name or 'xxx'}  学号：{student_id or 'xxxxxxxxxxx'}")
        else:
            self.student_info.setText("姓名：xxx  学号：xxxxxxxxxxx")

        self.question_widget.show_question()
        self.nav_panel.reset()
        self.nav_panel.refresh()
        self._update_progress()
        self._update_time_display()

        # 启动计时器
        self.remaining_time = EXAM_TIME_SECONDS
        self.timer_running = True
        self._timer.start(1000)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        if self.exam and not self.exam_submitted:
            self.timer_running = True
            if not self._timer.isActive():
                self._timer.start(1000)

    def hideEvent(self, event) -> None:
        super().hideEvent(event)
        self.timer_running = False
        self._timer.stop()

    def keyPressEvent(self, event) -> None:
        """处理考试页快捷键（ExamPage 自身获得焦点时）."""
        if self.exam_submitted or not self.exam:
            super().keyPressEvent(event)
            return

        if event.isAutoRepeat():
            return

        key = event.key()
        modifiers = event.modifiers()

        # Ctrl+N 下一未答
        if modifiers == Qt.KeyboardModifier.ControlModifier and key == Qt.Key.Key_N:
            self.jump_next_unanswered()
            return

        # Ctrl+A 查看答案
        if modifiers == Qt.KeyboardModifier.ControlModifier and key == Qt.Key.Key_A:
            self.show_answer()
            return

        # Ctrl+Enter 交卷
        if modifiers == Qt.KeyboardModifier.ControlModifier and key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.submit_exam()
            return

        if key == Qt.Key.Key_Left:
            self.prev_question()
            return
        if key == Qt.Key.Key_Right:
            self.next_question()
            return

        q = self.exam.get_question(self.exam.current_index)
        if not q:
            super().keyPressEvent(event)
            return

        # A-F 选择选项（单选题）
        if q.q_type == "单选":
            text = event.text().upper()
            if text in q.options:
                self.question_widget._on_choice(q.number, text)
                return

        # 判断题快捷键
        if q.q_type == "判断":
            ch = event.text().lower()
            if ch in ("t", "y", "1"):
                self.question_widget._on_choice(q.number, "A")
                return
            if ch in ("f", "n", "0"):
                self.question_widget._on_choice(q.number, "B")
                return

        super().keyPressEvent(event)

    # ── 计时器 ────────────────────────────────────────────────

    def _tick(self) -> None:
        if not self.timer_running or self.exam_submitted:
            return
        self.remaining_time -= 1
        if self.remaining_time <= 0:
            self.remaining_time = 0
            self._force_submit()
            return
        self._update_time_display()

    def _update_time_display(self) -> None:
        h = self.remaining_time // 3600
        m = (self.remaining_time % 3600) // 60
        s = self.remaining_time % 60
        self.time_label.setText(f"{h:02d}:{m:02d}:{s:02d}")

        c = Theme.get_current_colors()
        if self.remaining_time <= 300:
            self.time_label.setStyleSheet(
                f"color: {c['DANGER']}; font-size: 14pt; font-weight: bold; padding: 0 20px;"
            )
        elif self.remaining_time <= 600:
            self.time_label.setStyleSheet(
                f"color: {c['ACCENT']}; font-size: 14pt; font-weight: bold; padding: 0 20px;"
            )
        else:
            self.time_label.setStyleSheet(
                f"color: {c['TEXT']}; font-size: 14pt; font-weight: bold; padding: 0 20px;"
            )

    def _force_submit(self) -> None:
        if self.exam_submitted:
            return
        self.timer_running = False
        self._timer.stop()
        self.exam_submitted = True
        self.time_label.setText("00:00:00")
        themed_info(self, "提示", "考试时间到！系统将自动交卷。")
        results = self.exam.grade()
        score_pct = self._calc_score_pct(results)
        self._save_exam_progress("completed", score_pct)
        self._show_result(results)

    # ── 导航操作 ──────────────────────────────────────────────

    def prev_question(self) -> None:
        if self.exam and self.exam.current_index > 0:
            self.question_widget.save_current_answer()
            self.exam.current_index -= 1
            self.question_widget.show_question()
            self.nav_panel.refresh()

    def next_question(self) -> None:
        if self.exam and self.exam.current_index < self.exam.total_count - 1:
            self.question_widget.save_current_answer()
            self.exam.current_index += 1
            self.question_widget.show_question()
            self.nav_panel.refresh()

    def jump_next_unanswered(self) -> None:
        if not self.exam:
            return
        self.question_widget.save_current_answer()
        idx = self.exam.next_unanswered_index()
        if idx is not None:
            self.exam.current_index = idx
            self.question_widget.show_question()
            self.nav_panel.refresh()
        else:
            themed_info(self, "提示", "所有题目都已作答！")

    def jump_to(self, index: int) -> None:
        """跳转到指定题目."""
        if not self.exam:
            return
        self.question_widget.save_current_answer()
        if 0 <= index < self.exam.total_count:
            self.exam.current_index = index
            self.question_widget.show_question()
            self.nav_panel.refresh()

    def toggle_mark(self) -> None:
        if not self.exam:
            return
        self.exam.toggle_mark(self.exam.current_index)
        self.nav_panel.refresh()

    def show_answer(self) -> None:
        self.question_widget.display_correct_answer()

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
            themed_info(
                self, "已复制",
                "题目已复制到剪贴板，直接粘贴给豆包即可！",
            )

    def _copy_question_to_clipboard(self) -> None:
        if not self.exam:
            return
        q = self.exam.get_question(self.exam.current_index)
        if not q:
            return
        lines = [f"题目：{q.text}"]
        if q.options:
            for letter, content in q.options.items():
                lines.append(f"{letter}. {content}")
        lines.append(f"题型：{q.q_type}")
        if q.correct_answer:
            lines.append(f"参考答案：{q.correct_answer}")
        text = "\n".join(lines) + "\n请帮我解答这道题"
        QApplication.clipboard().setText(text)

    def redo_question(self) -> None:
        if not self.exam:
            return
        q = self.exam.get_question(self.exam.current_index)
        if not q:
            return

        if q.program_file and q.q_type in ("程序设计", "程序填空", "程序改错"):
            reply = themed_question(
                self, "确认重做",
                "确定要重做此题吗？\n答案和程序文件都将恢复为初始状态。",
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
            self.backup_mgr.restore_file(q.program_file, self.exam_file_path)
        else:
            reply = themed_question(
                self, "确认重做",
                "确定要重做此题吗？当前答案将被清空。",
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        self.exam.set_answer(q.number, "")
        self.question_widget.show_question()
        self.nav_panel.refresh()

    def open_program_file(self) -> None:
        """打开程序文件（用IDLE）."""
        if not self.exam:
            return
        q = self.exam.get_question(self.exam.current_index)
        if not q or not q.program_file:
            themed_info(self, "提示", "该题目没有对应的程序文件")
            return

        program_file = q.program_file
        if ".." in program_file or program_file.startswith(("/", "\\")):
            themed_critical(self, "错误", "不允许的文件路径")
            return

        exam_dir = os.path.dirname(self.exam_file_path)
        base_dir = os.path.join(exam_dir, "试题文件夹")
        if not os.path.exists(base_dir):
            base_dir = exam_dir

        program_path = os.path.normpath(os.path.join(base_dir, program_file))
        if not os.path.exists(program_path):
            alt_path = os.path.normpath(os.path.join(exam_dir, program_file))
            if os.path.exists(alt_path):
                program_path = alt_path
            else:
                themed_critical(self, "错误", f"找不到程序文件：{program_file}")
                return

        from hnust_exam.services.python_env import open_with_idle
        from hnust_exam.utils.constants import CURRENT_VERSION
        cfg = self.main_window.config_mgr.load_config()
        python_path = cfg.get("user_python_path", "")

        success = open_with_idle(program_path, python_path or None)
        if success:
            themed_info(self, "提示", f"已用IDLE打开：{program_file}\n修改完成后按Ctrl+S保存，然后回到本系统输入答案")
        else:
            # 回退到系统默认程序
            import subprocess
            try:
                if os.name == "nt":
                    os.startfile(program_path)
                else:
                    subprocess.run(["xdg-open", program_path], check=True)
                themed_info(self, "提示", f"已用默认程序打开：{program_file}\n（未检测到Python IDLE）")
            except Exception as e:
                themed_critical(self, "错误", f"打开文件失败：{e}")

    def open_exam_folder(self) -> None:
        exam_dir = os.path.dirname(self.exam_file_path)
        folder = os.path.join(exam_dir, "试题文件夹")
        if not os.path.exists(folder):
            folder = exam_dir
        try:
            if os.name == "nt":
                os.startfile(folder)
            else:
                import subprocess
                subprocess.run(["xdg-open", folder], check=True)
        except Exception as e:
            themed_critical(self, "错误", f"打开文件夹失败：{e}")

    # ── 交卷 ──────────────────────────────────────────────────

    def submit_exam(self) -> None:
        if self.exam_submitted or not self.exam:
            return

        self.question_widget.save_current_answer()

        # 交卷确认对话框
        from hnust_exam.views.dialogs.submit_dialog import SubmitDialog
        dlg = SubmitDialog(self.exam, self)
        if dlg.exec():
            self.timer_running = False
            self._timer.stop()
            self.exam_submitted = True
            self.backup_mgr.cleanup()
            results = self.exam.grade()
            score_pct = self._calc_score_pct(results)
            self._save_exam_progress("completed", score_pct)
            self._show_result(results)
        elif dlg.check_marked_index is not None:
            self.jump_to(dlg.check_marked_index)

    @staticmethod
    def _calc_score_pct(results: list[Result]) -> float:
        """根据判分结果计算得分百分比."""
        total = sum(r.score for r in results)
        if total == 0:
            return 0.0
        earned = sum(r.score for r in results if r.is_correct)
        return earned / total * 100

    def _show_result(self, results: list[Result] | None = None) -> None:
        """显示成绩页."""
        if results is None:
            results = self.exam.grade()
        self.main_window.result_page.setup_results(results, self.exam)
        self.main_window.show_result()

    # ── 进度更新 ──────────────────────────────────────────────

    def _update_progress(self) -> None:
        if not self.exam:
            return
        total = self.exam.total_count
        answered = self.exam.answered_count
        if total > 0:
            self._progress_bar.setValue(int(answered / total * 100))
        if answered == total and total > 0:
            self.progress_label.setText("已完成，可以交卷！")
        else:
            self.progress_label.setText(f"已完成 {answered} / {total} 题")
        self._update_progress_label_style()

    def _save_exam_progress(self, status: str, score_pct: float | None = None) -> None:
        """保存考试进度到配置."""
        progress = self.main_window.config_mgr.load_progress()
        exam_key = os.path.basename(self.exam_file_path)
        entry = progress.get(exam_key, {})
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        if status == "completed":
            entry["status"] = "completed"
            entry["last_completed"] = now_str
            if score_pct is not None:
                prev_best = entry.get("best_score", 0)
                entry["best_score"] = max(prev_best, round(score_pct, 1))
        elif status == "started":
            if entry.get("status") != "completed":
                entry["status"] = "started"
            entry["last_started"] = now_str

        progress[exam_key] = entry
        self.main_window.config_mgr.save_progress(progress)