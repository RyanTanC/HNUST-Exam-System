"""更新对话框：展示版本信息，支持自动下载替换（frozen）或打开浏览器（开发模式）."""

from __future__ import annotations

import sys
import webbrowser

from PySide6.QtCore import Qt, QThread, QTimer, Signal
from PySide6.QtWidgets import (
    QDialog,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QFrame,
    QVBoxLayout,
    QWidget,
    QSizePolicy,
)

from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.constants import CURRENT_VERSION
from hnust_exam.utils.theme import Theme


class _DownloadThread(QThread):
    """后台下载线程，通过信号将进度派发到主线程."""

    progress = Signal(int, float, int)
    finished = Signal(bool, str)

    def __init__(self, url: str, save_path: str, expected_size: int) -> None:
        super().__init__()
        self._url = url
        self._save_path = save_path
        self._expected_size = expected_size
        self._stop_flag = False

    def request_stop(self) -> None:
        """请求线程停止."""
        self._stop_flag = True

    def run(self) -> None:
        from hnust_exam.services.auto_updater import download_file

        def _on_progress(percent: int, speed_mb: float, remaining: int) -> None:
            if self._stop_flag:
                return
            self.progress.emit(percent, speed_mb, remaining)

        if self._stop_flag:
            self.finished.emit(False, "下载已取消")
            return

        ok, msg = download_file(
            self._url, self._save_path, _on_progress, self._expected_size,
        )
        if self._stop_flag and not ok:
            self.finished.emit(False, "下载已取消")
            return
        self.finished.emit(ok, msg)


class UpdateDialog(QDialog):
    """发现新版本对话框."""

    def __init__(self, info: dict, config_mgr: ConfigManager, parent=None) -> None:
        super().__init__(parent)
        self._info = info
        self._config_mgr = config_mgr
        self._download_thread: _DownloadThread | None = None
        self.setWindowTitle("发现新版本")
        self.setMinimumSize(520, 480)
        self.resize(560, 560)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self._build_ui()

    # ──────────────────────────────────────────────
    #  UI 构建
    # ──────────────────────────────────────────────

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._make_header(c))
        body = self._make_body(c)
        root.addWidget(body, 1)
        root.addWidget(self._make_footer(c))

    # ───────── 顶部：彩色头部 ─────────

    def _make_header(self, c: dict) -> QFrame:
        header = QFrame()
        header.setStyleSheet(f"background-color: {c['PRIMARY']};")
        header.setMinimumHeight(120)

        lay = QVBoxLayout(header)
        lay.setContentsMargins(28, 24, 28, 20)
        lay.setSpacing(6)

        badge = QLabel("✦ 发现新版本")
        badge.setStyleSheet(
            "color: rgba(255,255,255,0.85); font-size: 10pt; font-weight: 600; "
            "letter-spacing: 1px; background: rgba(255,255,255,0.15); "
            "border-radius: 10px; padding: 3px 12px;"
        )
        badge.setFixedWidth(badge.sizeHint().width())
        lay.addWidget(badge)

        current = self._info.get("current_ver", CURRENT_VERSION).lstrip("v")
        latest = self._info["latest_ver"].lstrip("v")

        version_label = QLabel(f"v{latest}")
        version_label.setStyleSheet(
            "color: white; font-size: 28pt; font-weight: 800;"
        )
        lay.addWidget(version_label)

        sub = QLabel(f"当前版本 v{current}  →  最新版本 v{latest}")
        sub.setStyleSheet(
            f"color: {c['HEADER_SUB_TEXT']}; font-size: 10pt;"
        )
        lay.addWidget(sub)

        return header

    # ───────── 主体：信息卡片 + 更新日志 + 进度 ─────────

    def _make_body(self, c: dict) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setStyleSheet(
            f"QScrollArea {{ background-color: {c['BG']}; border: none; }}"
            f"QScrollBar:vertical {{"
            f"  width: 6px; background: transparent; margin: 4px 2px 4px 0;"
            f"}}"
            f"QScrollBar::handle:vertical {{"
            f"  background: {c['BORDER']}; border-radius: 3px; min-height: 24px;"
            f"}}"
            f"QScrollBar::handle:vertical:hover {{"
            f"  background: {c['MUTED']};"
            f"}}"
            f"QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{"
            f"  height: 0; background: none;"
            f"}}"
        )

        container = QWidget()
        container.setStyleSheet(f"background-color: {c['BG']};")
        body_layout = QVBoxLayout(container)
        body_layout.setContentsMargins(24, 20, 24, 16)
        body_layout.setSpacing(14)

        # ── 版本信息芯片 ──
        chips_widget = QWidget()
        chips_widget.setStyleSheet("background: transparent;")
        chips = QHBoxLayout(chips_widget)
        chips.setContentsMargins(0, 0, 0, 0)
        chips.setSpacing(10)

        info = self._info
        current = info.get("current_ver", CURRENT_VERSION).lstrip("v")
        latest = info["latest_ver"].lstrip("v")
        published = info.get("published_at", "")

        chip_data = [
            ("当前版本", f"v{current}", c["TEXT"]),
            ("最新版本", f"v{latest}", c["SUCCESS"]),
        ]
        if published:
            chip_data.append(("发布时间", published, c["MUTED"]))

        for label, value, color in chip_data:
            chip = QFrame()
            chip.setObjectName("infoChip")
            chip.setStyleSheet(
                f"QFrame#infoChip {{"
                f"  background-color: {c['WHITE']};"
                f"  border: 1px solid {c['BORDER']};"
                f"  border-radius: 8px;"
                f"}}"
            )
            chip_lay = QVBoxLayout(chip)
            chip_lay.setContentsMargins(14, 10, 14, 10)
            chip_lay.setSpacing(2)

            lbl = QLabel(label)
            lbl.setStyleSheet(f"color: {c['MUTED']}; font-size: 8pt; border: none;")
            chip_lay.addWidget(lbl)

            val = QLabel(value)
            val.setStyleSheet(f"color: {color}; font-size: 13pt; font-weight: 700; border: none;")
            chip_lay.addWidget(val)

            chips.addWidget(chip, 1)

        body_layout.addWidget(chips_widget)

        # ── 更新日志卡片 ──
        notes_card = QFrame()
        notes_card.setObjectName("notesCard")
        notes_card.setStyleSheet(
            f"QFrame#notesCard {{"
            f"  background-color: {c['WHITE']};"
            f"  border: 1px solid {c['BORDER']};"
            f"  border-radius: 8px;"
            f"}}"
        )
        notes_card_lay = QVBoxLayout(notes_card)
        notes_card_lay.setContentsMargins(0, 0, 0, 0)
        notes_card_lay.setSpacing(0)

        notes_header = QLabel("  更新日志")
        notes_header.setStyleSheet(
            f"color: {c['TEXT']}; font-size: 11pt; font-weight: 700;"
            f"padding: 12px 16px 6px 16px; border: none;"
        )
        notes_card_lay.addWidget(notes_header)

        notes_sep = QFrame()
        notes_sep.setFrameShape(QFrame.Shape.HLine)
        notes_sep.setStyleSheet(f"border: none; border-top: 1px solid {c['BORDER']}; margin: 0 16px;")
        notes_card_lay.addWidget(notes_sep)

        notes_scroll = QScrollArea()
        notes_scroll.setWidgetResizable(True)
        notes_scroll.setFrameShape(QFrame.Shape.NoFrame)
        notes_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        notes_scroll.setMinimumHeight(120)
        notes_scroll.setStyleSheet(
            f"QScrollArea {{ background: transparent; border: none; }}"
            f"QScrollBar:vertical {{ width: 5px; background: transparent; margin: 2px 0; }}"
            f"QScrollBar::handle:vertical {{"
            f"  background: {c['BORDER']}; border-radius: 2px; min-height: 20px;"
            f"}}"
            f"QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{"
            f"  height: 0;"
            f"}}"
        )

        notes_content = QWidget()
        notes_content.setStyleSheet(f"background: transparent;")
        notes_layout = QVBoxLayout(notes_content)
        notes_layout.setContentsMargins(16, 10, 16, 10)
        notes_layout.setSpacing(0)

        release_notes = info.get("release_notes", "暂无更新日志")
        for line in release_notes.splitlines():
            line = line.strip()
            if not line:
                notes_layout.addSpacing(6)
                continue
            if line.startswith("### "):
                lbl = QLabel(line[4:])
                lbl.setWordWrap(True)
                lbl.setStyleSheet(
                    f"color: {c['TEXT']}; font-size: 10pt; font-weight: 700; "
                    f"padding: 6px 0 2px 0; border: none;"
                )
                notes_layout.addWidget(lbl)
            elif line.startswith("## "):
                lbl = QLabel(line[3:])
                lbl.setWordWrap(True)
                lbl.setStyleSheet(
                    f"color: {c['TEXT']}; font-size: 11pt; font-weight: 700; "
                    f"padding: 8px 0 2px 0; border: none;"
                )
                notes_layout.addWidget(lbl)
            elif line.startswith("- ") or line.startswith("* "):
                lbl = QLabel(f"  •  {line[2:]}")
                lbl.setWordWrap(True)
                lbl.setStyleSheet(
                    f"color: {c['TEXT']}; font-size: 10pt; "
                    f"padding: 1px 0 1px 12px; border: none;"
                )
                notes_layout.addWidget(lbl)
            else:
                lbl = QLabel(line)
                lbl.setWordWrap(True)
                lbl.setStyleSheet(
                    f"color: {c['TEXT']}; font-size: 10pt; "
                    f"padding: 1px 0; border: none;"
                )
                notes_layout.addWidget(lbl)

        notes_layout.addStretch()
        notes_scroll.setWidget(notes_content)
        notes_card_lay.addWidget(notes_scroll, 1)

        body_layout.addWidget(notes_card, 1)

        # ── 下载进度 ──
        self._progress_frame = QFrame()
        self._progress_frame.setObjectName("progressFrame")
        self._progress_frame.setStyleSheet(
            f"QFrame#progressFrame {{"
            f"  background-color: {c['WHITE']};"
            f"  border: 1px solid {c['BORDER']};"
            f"  border-radius: 8px;"
            f"  padding: 14px 16px;"
            f"}}"
        )
        progress_layout = QVBoxLayout(self._progress_frame)
        progress_layout.setContentsMargins(0, 0, 0, 0)
        progress_layout.setSpacing(8)

        progress_header_row = QHBoxLayout()
        progress_header_row.setContentsMargins(0, 0, 0, 0)

        progress_title = QLabel("下载进度")
        progress_title.setStyleSheet(f"color: {c['TEXT']}; font-size: 10pt; font-weight: 600; border: none;")
        progress_header_row.addWidget(progress_title)
        progress_header_row.addStretch()

        self._progress_percent = QLabel("0%")
        self._progress_percent.setStyleSheet(f"color: {c['PRIMARY']}; font-size: 10pt; font-weight: 700; border: none;")
        progress_header_row.addWidget(self._progress_percent)

        progress_layout.addLayout(progress_header_row)

        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 100)
        self._progress_bar.setValue(0)
        self._progress_bar.setTextVisible(False)
        self._progress_bar.setFixedHeight(6)
        self._progress_bar.setStyleSheet(
            f"QProgressBar {{"
            f"  background-color: {c['PROGRESS_BG']};"
            f"  border-radius: 3px; border: none;"
            f"}}"
            f"QProgressBar::chunk {{"
            f"  background-color: {c['PRIMARY']};"
            f"  border-radius: 3px;"
            f"}}"
        )
        progress_layout.addWidget(self._progress_bar)

        self._status_label = QLabel("准备下载...")
        self._status_label.setStyleSheet(f"color: {c['MUTED']}; font-size: 9pt; border: none;")
        progress_layout.addWidget(self._status_label)

        body_layout.addWidget(self._progress_frame)
        self._progress_frame.hide()

        # ── 底部提示 ──
        hint = QLabel("建议更新到最新版本以获得最佳体验（也可以稍后更新）")
        hint.setStyleSheet(f"color: {c['MUTED']}; font-size: 8pt;")
        body_layout.addWidget(hint)

        scroll.setWidget(container)
        return scroll

    # ───────── 底部按钮栏 ─────────

    def _make_footer(self, c: dict) -> QFrame:
        footer = QFrame()
        footer.setStyleSheet(
            f"background-color: {c['WHITE']};"
            f"border-top: 1px solid {c['BORDER']};"
        )

        lay = QHBoxLayout(footer)
        lay.setContentsMargins(24, 12, 24, 14)

        self._skip_btn = QPushButton("跳过此版本")
        self._skip_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._skip_btn.setStyleSheet(
            f"QPushButton {{"
            f"  background: transparent; color: {c['MUTED']}; font-size: 9pt;"
            f"  border: none; padding: 6px 10px;"
            f"}}"
            f"QPushButton:hover {{ color: {c['TEXT']}; }}"
        )
        self._skip_btn.clicked.connect(self._on_skip)
        lay.addWidget(self._skip_btn)

        self._later_btn = QPushButton("暂不更新")
        self._later_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._later_btn.setStyleSheet(
            f"QPushButton {{"
            f"  background-color: {c['SURFACE']}; color: {c['MUTED']};"
            f"  font-size: 10pt; padding: 7px 18px;"
            f"  border: 1px solid {c['BORDER']}; border-radius: 6px;"
            f"}}"
            f"QPushButton:hover {{"
            f"  background-color: {c['BORDER']}; color: {c['TEXT']};"
            f"}}"
        )
        self._later_btn.clicked.connect(self.reject)
        lay.addWidget(self._later_btn)

        lay.addStretch()

        btn_text = "  前往下载  " if not self._info.get("download_available", False) else "  立即更新  "
        self._update_btn = QPushButton(btn_text)
        self._update_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._update_btn.setStyleSheet(
            f"QPushButton {{"
            f"  background-color: {c['PRIMARY']}; color: white; font-size: 11pt;"
            f"  font-weight: 700; padding: 9px 28px;"
            f"  border: none; border-radius: 6px;"
            f"}}"
            f"QPushButton:hover {{"
            f"  background-color: {c['PRIMARY_HOVER']};"
            f"}}"
            f"QPushButton:disabled {{"
            f"  background-color: {c['BORDER']}; color: {c['MUTED']};"
            f"}}"
        )
        self._update_btn.clicked.connect(self._on_update)
        lay.addWidget(self._update_btn)

        return footer

    # ──────────────────────────────────────────────
    #  按钮事件
    # ──────────────────────────────────────────────

    def _on_update(self) -> None:
        url = self._info.get("download_url", "")

        # 没有可下载的 exe asset 时，打开降级页（或 Release 页）让用户手动下载
        if not self._info.get("download_available", False):
            fallback_url = self._info.get("fallback_url", "")
            target_url = fallback_url or self._info.get("release_url", url)
            if target_url:
                webbrowser.open(target_url)
            self.accept()
            return

        if not url:
            return

        if not getattr(sys, "frozen", False):
            webbrowser.open(url)
            self.accept()
            return

        from hnust_exam.services.auto_updater import (
            check_write_permission,
            get_app_exe_path,
            get_temp_download_path,
            replace_and_restart,
        )

        exe_path = get_app_exe_path()
        if not check_write_permission(exe_path):
            self._status_label.setText("权限不足，请以管理员身份运行")
            self._status_label.setStyleSheet("color: #e74c3c; font-size: 9pt;")
            self._progress_frame.show()
            return

        self._set_downloading_state()

        temp_path = get_temp_download_path()
        expected_size = self._info.get("expected_size", 0)

        self._download_thread = _DownloadThread(url, temp_path, expected_size)
        self._download_thread.progress.connect(self._on_download_progress)
        self._download_thread.finished.connect(
            lambda ok, msg: self._on_download_finished(ok, msg, temp_path),
        )
        self._download_thread.start()

    def _on_skip(self) -> None:
        self._config_mgr.save_skip_version(self._info["latest_ver"])
        self.reject()

    # ──────────────────────────────────────────────
    #  下载状态管理
    # ──────────────────────────────────────────────

    def _set_downloading_state(self) -> None:
        self._update_btn.setEnabled(False)
        self._skip_btn.setEnabled(False)
        self._later_btn.setEnabled(False)
        self._update_btn.setText("  下载中...  ")
        self._progress_frame.show()
        self._progress_bar.setValue(0)
        self._progress_percent.setText("0%")
        self._status_label.setText("准备下载...")

    def _on_download_progress(self, percent: int, speed_mb: float, remaining: int) -> None:
        self._progress_bar.setValue(percent)
        self._progress_percent.setText(f"{percent}%")
        self._status_label.setText(
            f"{speed_mb:.1f} MB/s  ·  剩余约 {remaining} 秒"
        )

    def _on_download_finished(
        self, success: bool, error_msg: str, temp_path: str,
    ) -> None:
        if not success:
            self._on_download_failed(error_msg)
            return

        self._status_label.setText("下载完成，正在替换并重启...")
        self._progress_bar.setValue(100)
        self._progress_percent.setText("100%")

        from hnust_exam.services.auto_updater import replace_and_restart

        ok, msg = replace_and_restart(temp_path)
        if not ok:
            self._on_download_failed(msg)

    def _on_download_failed(self, error_msg: str) -> None:
        c = Theme.get_current_colors()
        self._update_btn.setEnabled(True)
        self._skip_btn.setEnabled(True)
        self._later_btn.setEnabled(True)
        self._update_btn.setText("   重试更新   ")
        self._status_label.setText(f"更新失败：{error_msg}")
        self._status_label.setStyleSheet(f"color: #e74c3c; font-size: 9pt;")
        self._progress_bar.setStyleSheet(
            f"QProgressBar {{"
            f"  background-color: {c['PROGRESS_BG']};"
            f"  border-radius: 3px; border: none;"
            f"}}"
            f"QProgressBar::chunk {{"
            f"  background-color: #e74c3c;"
            f"  border-radius: 3px;"
            f"}}"
        )

    def closeEvent(self, event) -> None:
        if self._download_thread and self._download_thread.isRunning():
            self._download_thread.request_stop()
            if not self._download_thread.wait(3000):
                self._download_thread.terminate()
                self._download_thread.wait()
        super().closeEvent(event)
