from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from utils.file_utils import open_folder


def card(*widgets: QWidget) -> QFrame:
    box = QFrame()
    box.setObjectName("card")
    layout = QVBoxLayout(box)
    layout.setContentsMargins(20, 18, 20, 18)
    layout.setSpacing(12)
    for widget in widgets:
        layout.addWidget(widget)
    return box


def title_label(text: str, name: str = "pageTitle") -> QLabel:
    label = QLabel(text)
    label.setObjectName(name)
    return label


def muted(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("muted")
    label.setWordWrap(True)
    return label


class PrimaryButton(QPushButton):
    def __init__(self, text: str, parent=None) -> None:
        super().__init__(text, parent)
        self.setObjectName("primary")
        self.setCursor(Qt.CursorShape.PointingHandCursor)


class SecondaryButton(QPushButton):
    def __init__(self, text: str, parent=None) -> None:
        super().__init__(text, parent)
        self.setObjectName("secondary")
        self.setCursor(Qt.CursorShape.PointingHandCursor)


class GhostButton(QPushButton):
    def __init__(self, text: str, parent=None) -> None:
        super().__init__(text, parent)
        self.setObjectName("ghost")
        self.setCursor(Qt.CursorShape.PointingHandCursor)


class DropZone(QFrame):
    files_dropped = Signal(list)

    def __init__(self, hint: str = "将 Excel 文件拖到这里", parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setAcceptDrops(True)
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label = QLabel(hint)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setObjectName("muted")
        layout.addWidget(self.label)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            self.setProperty("active", True)
            self.style().unpolish(self)
            self.style().polish(self)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self.setProperty("active", False)
        self.style().unpolish(self)
        self.style().polish(self)
        event.accept()

    def dropEvent(self, event: QDropEvent) -> None:
        self.setProperty("active", False)
        self.style().unpolish(self)
        self.style().polish(self)
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile()]
        if paths:
            self.files_dropped.emit(paths)
        event.acceptProposedAction()


class StatusBox(QFrame):
    retry_clicked = Signal()
    open_clicked = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setVisible(False)
        self.layout_box = QVBoxLayout(self)
        self.title = QLabel()
        self.detail = QLabel()
        self.detail.setWordWrap(True)
        self.detail.setObjectName("muted")
        self.retry_btn = SecondaryButton("重新选择文件")
        self.open_btn = SecondaryButton("打开输出文件夹")
        self.retry_btn.clicked.connect(self.retry_clicked.emit)
        self.open_btn.clicked.connect(self.open_clicked.emit)
        self.layout_box.addWidget(self.title)
        self.layout_box.addWidget(self.detail)
        self.layout_box.addWidget(self.retry_btn, alignment=Qt.AlignmentFlag.AlignLeft)
        self.layout_box.addWidget(self.open_btn, alignment=Qt.AlignmentFlag.AlignLeft)

    def show_error(self, title: str, hint: str) -> None:
        self.setObjectName("errorBox")
        self.title.setObjectName("errorTitle")
        self.title.setText(f"出错了：{title}")
        self.detail.setText(hint)
        self.retry_btn.setVisible(True)
        self.open_btn.setVisible(False)
        self.setVisible(True)
        self._refresh()

    def show_ok(self, title: str, hint: str, folder: str | Path | None = None) -> None:
        self.setObjectName("okBox")
        self.title.setObjectName("okTitle")
        self.title.setText(title)
        self.detail.setText(hint)
        self.retry_btn.setVisible(False)
        self.open_btn.setVisible(bool(folder))
        self._folder = Path(folder) if folder else None
        self.setVisible(True)
        self._refresh()

    def clear(self) -> None:
        self.setVisible(False)

    def _refresh(self) -> None:
        self.style().unpolish(self)
        self.style().polish(self)
        self.title.style().unpolish(self.title)
        self.title.style().polish(self.title)


class ProgressPanel(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.label = QLabel("正在处理……")
        self.bar = QProgressBar()
        self.bar.setRange(0, 100)
        self.bar.setValue(0)
        self.bar.setFormat("%p%")
        layout.addWidget(self.label)
        layout.addWidget(self.bar)
        self.setVisible(False)

    def start(self, text: str = "正在处理……") -> None:
        self.label.setText(text)
        self.bar.setValue(0)
        self.setVisible(True)

    def update_progress(self, value: int, text: str) -> None:
        self.label.setText(text)
        self.bar.setValue(value)

    def finish(self) -> None:
        self.label.setText("处理完成！")
        self.bar.setValue(100)


class FieldSelect(QWidget):
    def __init__(self, label: str, parent=None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        self.label = QLabel(label)
        self.combo = QComboBox()
        self.combo.addItem("请选择", "")
        layout.addWidget(self.label)
        layout.addWidget(self.combo)

    def set_fields(self, fields: list[str], selected: str | None = None) -> None:
        self.combo.clear()
        self.combo.addItem("请选择", "")
        for field in fields:
            self.combo.addItem(field, field)
        if selected and selected in fields:
            self.combo.setCurrentText(selected)

    def value(self) -> str:
        return str(self.combo.currentData() or "")


def checkbox(text: str) -> QCheckBox:
    box = QCheckBox(text)
    return box


def open_output(path: str | Path) -> None:
    open_folder(path)
