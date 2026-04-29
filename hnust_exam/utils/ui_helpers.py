"""主题适配的对话框工具函数."""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtWidgets import QMessageBox

from hnust_exam.utils.theme import Theme

if TYPE_CHECKING:
    from PySide6.QtWidgets import QWidget


def _apply_msgbox_style(dlg: QMessageBox) -> None:
    """为 QMessageBox 应用当前主题样式."""
    c = Theme.get_current_colors()
    dlg.setStyleSheet(
        f"QMessageBox {{"
        f"  background-color: {c['BG']};"
        f"  color: {c['TEXT']};"
        f"}}"
        f"QMessageBox QLabel {{"
        f"  color: {c['TEXT']};"
        f"  background: transparent;"
        f"}}"
        f"QPushButton {{"
        f"  background-color: {c['SURFACE']};"
        f"  color: {c['TEXT']};"
        f"  border: 1px solid {c['BORDER']};"
        f"  padding: 6px 20px;"
        f"  border-radius: 4px;"
        f"  min-width: 70px;"
        f"}}"
        f"QPushButton:hover {{"
        f"  background-color: {c['PRIMARY']};"
        f"  color: white;"
        f"  border: none;"
        f"}}"
    )


def themed_question(
    parent: QWidget | None,
    title: str,
    text: str,
    buttons: QMessageBox.StandardButton = QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
    default: QMessageBox.StandardButton = QMessageBox.StandardButton.No,
) -> QMessageBox.StandardButton:
    """主题适配的确认对话框."""
    dlg = QMessageBox(parent)
    dlg.setWindowTitle(title)
    dlg.setText(text)
    dlg.setStandardButtons(buttons)
    dlg.setDefaultButton(default)
    dlg.setIcon(QMessageBox.Icon.Question)
    _apply_msgbox_style(dlg)
    return dlg.exec()


def themed_info(
    parent: QWidget | None,
    title: str,
    text: str,
) -> None:
    """主题适配的提示对话框."""
    dlg = QMessageBox(parent)
    dlg.setWindowTitle(title)
    dlg.setText(text)
    dlg.setStandardButtons(QMessageBox.StandardButton.Ok)
    dlg.setIcon(QMessageBox.Icon.Information)
    _apply_msgbox_style(dlg)
    dlg.exec()


def themed_warning(
    parent: QWidget | None,
    title: str,
    text: str,
) -> None:
    """主题适配的警告对话框."""
    dlg = QMessageBox(parent)
    dlg.setWindowTitle(title)
    dlg.setText(text)
    dlg.setStandardButtons(QMessageBox.StandardButton.Ok)
    dlg.setIcon(QMessageBox.Icon.Warning)
    _apply_msgbox_style(dlg)
    dlg.exec()


def themed_critical(
    parent: QWidget | None,
    title: str,
    text: str,
) -> None:
    """主题适配的错误对话框."""
    dlg = QMessageBox(parent)
    dlg.setWindowTitle(title)
    dlg.setText(text)
    dlg.setStandardButtons(QMessageBox.StandardButton.Ok)
    dlg.setIcon(QMessageBox.Icon.Critical)
    _apply_msgbox_style(dlg)
    dlg.exec()
