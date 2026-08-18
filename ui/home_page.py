from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QFrame, QGridLayout, QLabel, QVBoxLayout, QWidget

from ui.widgets import PrimaryButton, muted


class HomePage(QWidget):
    open_feature = Signal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 28, 32, 28)
        layout.setSpacing(18)

        title = QLabel("Excel小工具箱")
        title.setObjectName("appTitle")
        subtitle = QLabel("批量处理 Excel，让重复工作一键完成")
        subtitle.setObjectName("appSubtitle")
        hint = muted("打开后选择文件，点一下按钮就能完成原来要大量复制粘贴的工作。全程在您的电脑本地处理，不联网，也不上传文件。")

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(hint)

        grid = QGridLayout()
        grid.setSpacing(16)
        cards = [
            ("merge", "Excel 批量合并", "多个 Excel 合并成一个总表"),
            ("split", "Excel 数据拆分", "按照指定字段自动拆成多个 Excel"),
            ("clean", "重复数据清洗", "一键删除重复数据和空白行"),
            ("sales", "销售数据汇总", "自动统计销售额、数量和人员排名"),
        ]
        for index, (key, name, desc) in enumerate(cards):
            grid.addWidget(self._card(key, name, desc), index // 2, index % 2)
        layout.addLayout(grid)
        layout.addStretch()

    def _card(self, key: str, name: str, desc: str) -> QFrame:
        box = QFrame()
        box.setObjectName("card")
        inner = QVBoxLayout(box)
        inner.setContentsMargins(22, 20, 22, 20)
        inner.setSpacing(8)
        heading = QLabel(name)
        heading.setObjectName("pageTitle")
        heading.setStyleSheet("font-size: 18px;")
        inner.addWidget(heading)
        inner.addWidget(muted(desc))
        inner.addStretch()
        btn = PrimaryButton("立即使用")
        btn.clicked.connect(lambda: self.open_feature.emit(key))
        inner.addWidget(btn, alignment=Qt.AlignmentFlag.AlignLeft)
        return box
