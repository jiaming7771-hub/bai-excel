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
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from utils.file_utils import open_file, open_folder


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
        self.setFixedHeight(40)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)


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

    def __init__(self, text: str, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setAcceptDrops(True)
        layout = QVBoxLayout(self)
        self.label = QLabel(text)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setObjectName("muted")
        layout.addWidget(self.label)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            self.setProperty("active", True)
            self.style().unpolish(self)
            self.style().polish(self)
            event.acceptProposedAction()

    def dragLeaveEvent(self, event) -> None:
        self.setProperty("active", False)
        self.style().unpolish(self)
        self.style().polish(self)

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
    report_clicked = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setVisible(False)
        self._folder: Path | None = None
        self._report: Path | None = None
        self.layout_box = QVBoxLayout(self)
        self.title = QLabel()
        self.detail = QLabel()
        self.detail.setWordWrap(True)
        self.detail.setObjectName("muted")
        self.retry_btn = SecondaryButton("重新选择文件")
        self.open_btn = SecondaryButton("打开结果文件")
        self.report_btn = SecondaryButton("打开处理报告")
        self.retry_btn.clicked.connect(self.retry_clicked.emit)
        self.open_btn.clicked.connect(self.open_clicked.emit)
        self.report_btn.clicked.connect(self.report_clicked.emit)
        self.layout_box.addWidget(self.title)
        self.layout_box.addWidget(self.detail)
        self.layout_box.addWidget(self.retry_btn, alignment=Qt.AlignmentFlag.AlignLeft)
        self.layout_box.addWidget(self.open_btn, alignment=Qt.AlignmentFlag.AlignLeft)
        self.layout_box.addWidget(self.report_btn, alignment=Qt.AlignmentFlag.AlignLeft)

    def show_error(self, title: str, hint: str) -> None:
        self.setObjectName("errorBox")
        self.title.setObjectName("errorTitle")
        self.title.setText(f"出错了：{title}")
        self.detail.setText(hint)
        self.retry_btn.setVisible(True)
        self.open_btn.setVisible(False)
        self.report_btn.setVisible(False)
        self.setVisible(True)
        self._refresh()

    def show_ok(self, title: str, hint: str, folder: str | Path | None = None, report: str | Path | None = None) -> None:
        self.setObjectName("okBox")
        self.title.setObjectName("okTitle")
        self.title.setText(title)
        self.detail.setText(hint)
        self.retry_btn.setVisible(False)
        self._folder = Path(folder) if folder else None
        self._report = Path(report) if report else None
        self.open_btn.setVisible(bool(self._folder))
        self.report_btn.setVisible(bool(self._report))
        self.setVisible(True)
        self._refresh()

    def clear(self) -> None:
        self.setVisible(False)

    def _refresh(self) -> None:
        self.style().unpolish(self)
        self.style().polish(self)
        self.title.style().unpolish(self.title)
        self.title.style().polish(self.title)


class ProgressPanel(QFrame):
    """独立进度卡片：与上方表格分离，避免文字叠在表头上。"""

    cancel_clicked = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("progressPanel")
        self.setMinimumHeight(78)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        self.label = QLabel("正在处理……")
        self.label.setObjectName("progressLabel")
        self.label.setWordWrap(True)

        row = QHBoxLayout()
        row.setSpacing(10)
        self.bar = QProgressBar()
        self.bar.setRange(0, 100)
        self.bar.setValue(0)
        self.bar.setTextVisible(True)
        self.bar.setFormat("%p%")
        self.bar.setFixedHeight(16)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setObjectName("secondary")
        self.cancel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.cancel_btn.setFixedWidth(72)
        self.cancel_btn.clicked.connect(self.cancel_clicked.emit)
        row.addWidget(self.bar, stretch=1)
        row.addWidget(self.cancel_btn, alignment=Qt.AlignmentFlag.AlignVCenter)

        layout.addWidget(self.label)
        layout.addLayout(row)
        self.setVisible(False)

    def start(self, text: str = "正在处理……") -> None:
        self.label.setText(text)
        self.bar.setValue(0)
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setVisible(True)
        self.setVisible(True)

    def update_progress(self, value: int, text: str) -> None:
        self.label.setText(text)
        self.bar.setValue(max(0, min(100, int(value))))
        if not self.isVisible():
            self.start(text)
        self.cancel_btn.setEnabled(True)

    def finish(self, text: str = "处理完成！") -> None:
        self.label.setText(text)
        self.bar.setValue(100)
        self.cancel_btn.setEnabled(False)

    def hide_panel(self) -> None:
        self.cancel_btn.setEnabled(False)
        self.setVisible(False)


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
    target = Path(path)
    if target.is_file():
        open_file(target)
    else:
        open_folder(target)
