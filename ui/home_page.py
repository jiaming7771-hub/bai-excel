from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QFrame, QGridLayout, QLabel, QSizePolicy, QVBoxLayout, QWidget

from app_config import APP_VERSION
from ui.widgets import PrimaryButton, muted

# 销售汇总先不放首页：透视一下就能做，卖点不够强
SHOW_SALES = False

ALL_FEATURE_CARDS = [
    ("merge", "Excel 批量合并", "多表合并，自动对齐列名，带来源文件"),
    ("split", "Excel 数据拆分", "按字段 / 工作表 / 行数拆成多个文件"),
    ("clean", "数据清洗", "去重可保留最新，异常标出，修复可追溯"),
    ("compare", "Excel 两表对比", "找出仅在A、仅在B、以及改动字段"),
    ("sales", "销售数据汇总", "自动统计销售额、数量和人员排名"),
]


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
        layout.addWidget(self._whats_new())

        grid = QGridLayout()
        grid.setSpacing(16)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        cards = [c for c in ALL_FEATURE_CARDS if SHOW_SALES or c[0] != "sales"]
        for index, (key, name, desc) in enumerate(cards):
            card = self._card(key, name, desc)
            card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
            grid.addWidget(card, index // 2, index % 2)
        for row in range((len(cards) + 1) // 2):
            grid.setRowStretch(row, 1)
        layout.addLayout(grid)
        layout.addStretch()
        layout.addWidget(muted(f"版本 {APP_VERSION}  ·  本地离线处理，不上传文件"))

    def _whats_new(self) -> QFrame:
        box = QFrame()
        box.setObjectName("card")
        inner = QVBoxLayout(box)
        inner.setContentsMargins(18, 14, 18, 14)
        inner.setSpacing(6)
        heading = QLabel(f"本版更新（v{APP_VERSION}）")
        heading.setObjectName("pageTitle")
        heading.setStyleSheet("font-size: 16px;")
        body = muted(
            "1. 合并：自动对齐相近列名，结果会写明来自哪个文件\n"
            "2. 拆分：除了按字段，还能按工作表、按行数拆\n"
            "3. 清洗：去重可留最新一条，也能用两列一起判断重复\n"
            "4. 新增：两表对比，一眼看出多了谁、少了谁、改了啥"
        )
        inner.addWidget(heading)
        inner.addWidget(body)
        return box

    def _card(self, key: str, name: str, desc: str) -> QFrame:
        box = QFrame()
        box.setObjectName("card")
        box.setMinimumHeight(172)
        inner = QVBoxLayout(box)
        inner.setContentsMargins(22, 20, 22, 22)
        inner.setSpacing(8)
        heading = QLabel(name)
        heading.setObjectName("pageTitle")
        heading.setStyleSheet("font-size: 18px;")
        inner.addWidget(heading)
        inner.addWidget(muted(desc))
        inner.addStretch(1)
        inner.addSpacing(4)
        btn = PrimaryButton("立即使用")
        btn.clicked.connect(lambda *_, feature_key=key: self.open_feature.emit(feature_key))
        inner.addWidget(btn, alignment=Qt.AlignmentFlag.AlignLeft)
        return box
