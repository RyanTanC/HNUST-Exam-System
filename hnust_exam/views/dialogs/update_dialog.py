"""更新对话框."""

from __future__ import annotations

import webbrowser

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QFrame,
)

from hnust_exam.services.config_manager import ConfigManager
from hnust_exam.utils.constants import CURRENT_VERSION
from hnust_exam.utils.theme import Theme


class UpdateDialog(QDialog):
    """发现新版本对话框."""

    def __init__(self, info: dict, config_mgr: ConfigManager, parent=None) -> None:
        super().__init__(parent)
        self._info = info
        self._config_mgr = config_mgr
        self.setWindowTitle("发现新版本")
        self.setMinimumSize(500, 400)
        self._build_ui()

    def _build_ui(self) -> None:
        c = Theme.get_current_colors()
        info = self._info

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 顶部
        header = QFrame()
        header.setStyleSheet(f"background-color: {c['PRIMARY']}; padding: 22px 30px;")
        header_layout = QHBoxLayout(header)

        # 图标
        icon_circle = QLabel("↑")
        icon_circle.setFixedSize(56, 56)
        icon_circle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_circle.setStyleSheet(
            f"background-color: white; color: {c['PRIMARY']}; font-size: 22pt; "
            f"font-weight: bold; border-radius: 28px;"
        )
        header_layout.addWidget(icon_circle)

        header_text = QVBoxLayout()
        title = QLabel("发现新版本")
        title.setStyleSheet("color: white; font-size: 18pt; font-weight: bold;")
        header_text.addWidget(title)
        ver_line = QLabel(f"v{CURRENT_VERSION.lstrip('v')} → {info['latest_ver']}")
        ver_line.setStyleSheet(f"color: {c['HEADER_SUB_TEXT']}; font-size: 11pt;")
        header_text.addWidget(ver_line)
        header_layout.addLayout(header_text)
        header_layout.addStretch()

        layout.addWidget(header)

        # 主体
        body = QFrame()
        body.setStyleSheet(f"background-color: {c['WHITE']}; padding: 20px 24px;")
        body_layout = QVBoxLayout(body)

        # 版本信息
        info_card = QFrame()
        info_card.setStyleSheet(
            f"background-color: {c['SURFACE']}; "
            f"border: 1px solid {c['BORDER']}; border-radius: 4px; padding: 14px 16px;"
        )
        info_layout = QVBoxLayout(info_card)

        items = [
            ("当前版本", CURRENT_VERSION, c["MUTED"]),
            ("最新版本", info["latest_ver"], c["SUCCESS"]),
        ]
        if info.get("published_at"):
            items.append(("发布时间", info["published_at"], c["MUTED"]))

        for label, value, color in items:
            row = QFrame()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 3, 0, 3)
            lbl = QLabel(f"{label}：")
            lbl.setStyleSheet(f"color: {c['MUTED']}; font-size: 10pt; min-width: 80px;")
            row_layout.addWidget(lbl)
            val = QLabel(value)
            val.setStyleSheet(f"color: {color}; font-size: 10pt; font-weight: bold;")
            row_layout.addWidget(val)
            row_layout.addStretch()
            info_layout.addWidget(row)

        body_layout.addWidget(info_card)

        # 更新日志
        log_title = QLabel("更新日志")
        log_title.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold; padding: 16px 0 6px 0;")
        body_layout.addWidget(log_title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"QScrollArea {{ border: 1px solid {c['BORDER']}; border-radius: 4px; }}")

        notes_content = QWidget()
        notes_content.setStyleSheet(f"background-color: {c['NOTES_BG']}; padding: 12px;")
        notes_layout = QVBoxLayout(notes_content)

        release_notes = info.get("release_notes", "暂无更新日志")
        for line in release_notes.splitlines():
            line = line.strip()
            if not line:
                notes_layout.addWidget(QLabel(""))
                continue
            if line.startswith("### "):
                lbl = QLabel(line[4:])
                lbl.setWordWrap(True)
                lbl.setStyleSheet(f"color: {c['TEXT']}; font-size: 10pt; font-weight: bold; padding: 8px 0 2px 0;")
                notes_layout.addWidget(lbl)
            elif line.startswith("## "):
                lbl = QLabel(line[3:])
                lbl.setWordWrap(True)
                lbl.setStyleSheet(f"color: {c['TEXT']}; font-size: 11pt; font-weight: bold; padding: 8px 0 2px 0;")
                notes_layout.addWidget(lbl)
            elif line.startswith("- ") or line.startswith("* "):
                lbl = QLabel(f"  •  {line[2:]}")
                lbl.setWordWrap(True)
                lbl.setStyleSheet(f"color: {c['TEXT']}; font-size: 10pt; padding: 1px 0 1px 16px;")
                notes_layout.addWidget(lbl)
            else:
                lbl = QLabel(line)
                lbl.setWordWrap(True)
                lbl.setStyleSheet(f"color: {c['TEXT']}; font-size: 10pt; padding: 1px 0;")
                notes_layout.addWidget(lbl)

        scroll.setWidget(notes_content)
        body_layout.addWidget(scroll, 1)

        hint = QLabel("建议更新到最新版本以获得最佳体验（也可以稍后更新）")
        hint.setStyleSheet(f"color: {c['MUTED']}; font-size: 8pt; padding-top: 4px;")
        body_layout.addWidget(hint)

        layout.addWidget(body, 1)

        # 底部按钮
        btn_bar = QFrame()
        btn_bar.setStyleSheet(f"background-color: {c['WHITE']}; border-top: 1px solid {c['BORDER']}; padding: 12px 24px;")
        btn_layout = QHBoxLayout(btn_bar)

        skip_btn = QPushButton("  跳过此版本  ")
        skip_btn.setStyleSheet(
            f"background-color: transparent; color: {c['MUTED']}; font-size: 9pt; border: none;"
        )
        skip_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        skip_btn.clicked.connect(self._on_skip)
        btn_layout.addWidget(skip_btn)

        later_btn = QPushButton("  暂不更新  ")
        later_btn.setStyleSheet(
            f"background-color: {c['SURFACE']}; color: {c['MUTED']}; font-size: 10pt; padding: 6px 14px; border-radius: 4px;"
        )
        later_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        later_btn.clicked.connect(self.reject)
        btn_layout.addWidget(later_btn)

        btn_layout.addStretch()

        update_btn = QPushButton("   立即更新   ")
        update_btn.setStyleSheet(
            f"background-color: {c['PRIMARY']}; color: white; font-size: 11pt; "
            f"font-weight: bold; padding: 8px 24px; border: none; border-radius: 4px;"
        )
        update_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        update_btn.clicked.connect(self._on_update)
        btn_layout.addWidget(update_btn)

        layout.addWidget(btn_bar)

    def _on_update(self) -> None:
        url = self._info.get("download_url", "")
        if url:
            webbrowser.open(url)
        self.accept()

    def _on_skip(self) -> None:
        self._config_mgr.save_skip_version(self._info["latest_ver"])
        self.reject()
