"""Consistent user-facing dialogs for the CNC program-sheet application.

The dialogs preserve the existing actions and log information while making the
first screen readable for operators. Technical details remain available on
demand instead of overwhelming normal error messages.
"""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from PySide6.QtCore import QPoint, QPropertyAnimation, QTimer, Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


_DIALOG_COLORS = {
    "info": ("i", "#0F6CBD", "#EAF4FF"),
    "warning": ("!", "#A15C00", "#FFF4CE"),
    "error": ("!", "#D83B01", "#FFF0ED"),
    "question": ("?", "#5B5B5B", "#F2F4F7"),
}


class ApplicationDialog(QDialog):
    """One compact dialog surface for information, warnings and decisions."""

    def __init__(
        self,
        title: str,
        message: str,
        *,
        kind: str = "info",
        details: str = "",
        buttons: Iterable[tuple[str, str]] = (("关闭", "close"),),
        checkbox_text: str = "",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(480)
        self.setModal(True)
        self.choice = ""
        self.checkbox: QCheckBox | None = None
        self._details = details.strip()

        icon_text, color, tint = _DIALOG_COLORS.get(kind, _DIALOG_COLORS["info"])
        root = QVBoxLayout(self)
        root.setContentsMargins(22, 20, 22, 18)
        root.setSpacing(12)

        heading = QHBoxLayout()
        icon = QLabel(icon_text)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setFixedSize(36, 36)
        icon.setStyleSheet(
            f"background: {color}; color: white; border-radius: 18px; "
            "font-size: 21px; font-weight: 700;"
        )
        heading.addWidget(icon, 0, Qt.AlignmentFlag.AlignTop)
        title_label = QLabel(title)
        title_label.setStyleSheet("color: #1F2937; font-size: 16px; font-weight: 700;")
        title_label.setWordWrap(True)
        heading.addWidget(title_label, 1)
        root.addLayout(heading)

        message_label = QLabel(message)
        message_label.setWordWrap(True)
        message_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        message_label.setStyleSheet(
            f"color: #1F2937; background: {tint}; border-radius: 5px; padding: 9px 11px;"
        )
        root.addWidget(message_label)

        self.details_view = QPlainTextEdit()
        self.details_view.setReadOnly(True)
        self.details_view.setPlainText(self._details)
        self.details_view.setMinimumHeight(100)
        self.details_view.setMaximumHeight(170)
        self.details_view.setVisible(False)
        if self._details:
            self.details_button = QPushButton("查看详细信息")
            self.details_button.clicked.connect(self._toggle_details)
            root.addWidget(self.details_button, 0, Qt.AlignmentFlag.AlignLeft)
            root.addWidget(self.details_view)

        if checkbox_text:
            self.checkbox = QCheckBox(checkbox_text)
            root.addWidget(self.checkbox)

        actions = QDialogButtonBox()
        action_list = list(buttons)
        for index, (label, value) in enumerate(action_list):
            button = QPushButton(label)
            if index == 0:
                button.setObjectName("primaryButton")
                button.setDefault(True)
            button.clicked.connect(lambda _checked=False, selected=value: self._choose(selected))
            actions.addButton(button, QDialogButtonBox.ButtonRole.ActionRole)
        root.addWidget(actions)

    def _toggle_details(self) -> None:
        visible = not self.details_view.isVisible()
        self.details_view.setVisible(visible)
        self.details_button.setText("收起详细信息" if visible else "查看详细信息")
        self.adjustSize()

    def _choose(self, value: str) -> None:
        self.choice = value
        self.accept()

    @property
    def checkbox_checked(self) -> bool:
        return bool(self.checkbox and self.checkbox.isChecked())


class ErrorDialog(ApplicationDialog):
    """Error dialog with an optional, respectful repeated-error reminder."""

    def __init__(
        self,
        title: str,
        message: str,
        log_path: Path,
        *,
        technical_details: str = "",
        repeated_count: int = 1,
        show_repeat_nudge: bool = False,
        buttons: Iterable[tuple[str, str]] = (("关闭", "close"),),
        parent: QWidget | None = None,
    ) -> None:
        details = technical_details.strip()
        log_text = f"详细日志：{log_path}"
        details = f"{details}\n\n{log_text}" if details else log_text
        super().__init__(
            title,
            message,
            kind="error",
            details=details,
            buttons=buttons,
            parent=parent,
        )
        self._tap_icon: QLabel | None = None
        self._tap_animation: QPropertyAnimation | None = None
        if show_repeat_nudge and repeated_count >= 2:
            self._add_repeat_nudge(repeated_count)

    def _add_repeat_nudge(self, repeated_count: int) -> None:
        root = self.layout()
        if not isinstance(root, QVBoxLayout):
            return
        nudge = QFrame()
        nudge.setObjectName("repeatErrorNudge")
        nudge.setStyleSheet(
            "QFrame#repeatErrorNudge { background: #FFF4CE; border: 1px solid #E8B931; "
            "border-radius: 6px; }"
        )
        nudge_layout = QHBoxLayout(nudge)
        nudge_layout.setContentsMargins(12, 8, 12, 8)
        self._tap_icon = QLabel("✋")
        self._tap_icon.setFixedWidth(34)
        self._tap_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._tap_icon.setStyleSheet("font-size: 22px;")
        nudge_layout.addWidget(self._tap_icon)
        nudge_text = QLabel(
            f"趣味拍拍提醒：同一问题已连续出现 {repeated_count} 次。\n"
            "请先检查提示内容，再试一次。"
        )
        nudge_text.setWordWrap(True)
        nudge_text.setStyleSheet("color: #4A3300;")
        nudge_layout.addWidget(nudge_text, 1)
        root.insertWidget(max(0, root.count() - 1), nudge)
        QTimer.singleShot(60, self._start_tap_animation)

    def _start_tap_animation(self) -> None:
        if self._tap_icon is None or not self.isVisible():
            return
        origin = self._tap_icon.pos()
        near = origin + QPoint(12, 0)
        animation = QPropertyAnimation(self._tap_icon, b"pos", self)
        animation.setDuration(520)
        animation.setKeyValueAt(0.0, origin)
        animation.setKeyValueAt(0.24, near)
        animation.setKeyValueAt(0.42, origin + QPoint(5, 0))
        animation.setKeyValueAt(0.66, near)
        animation.setKeyValueAt(1.0, origin)
        animation.setLoopCount(2)
        self._tap_animation = animation
        animation.start()
