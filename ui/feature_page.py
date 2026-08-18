from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QFrame, QScrollArea, QVBoxLayout, QWidget

from ui.task_runner import TaskHost
from ui.widgets import GhostButton, ProgressPanel, StatusBox, muted, title_label


class FeaturePage(QWidget):
    back_requested = Signal()

    def __init__(self, title: str, hint: str, parent=None) -> None:
        super().__init__(parent)
        self.tasks = TaskHost(self)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        self.layout_box = QVBoxLayout(inner)
        self.layout_box.setContentsMargins(32, 24, 32, 24)
        self.layout_box.setSpacing(14)
        scroll.setWidget(inner)
        outer.addWidget(scroll)

        self.back_btn = GhostButton("← 返回首页")
        self.back_btn.clicked.connect(self.back_requested.emit)
        self.layout_box.addWidget(self.back_btn)
        self.layout_box.addWidget(title_label(title))
        self.layout_box.addWidget(muted(hint))

        self.progress = ProgressPanel()
        self.status = StatusBox()
        self.status.retry_clicked.connect(self.on_retry)
        self.progress.cancel_clicked.connect(self.tasks.cancel)

    def attach_status(self) -> None:
        self.layout_box.addWidget(self.progress)
        self.layout_box.addWidget(self.status)
        self.layout_box.addStretch()

    def busy(self) -> bool:
        return self.tasks.running()

    def on_retry(self) -> None:
        self.status.clear()

    def set_busy(self, busy: bool) -> None:
        if busy:
            self.progress.start("正在处理……")
            self.status.clear()
        else:
            self.progress.hide_panel()
