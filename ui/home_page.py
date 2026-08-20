from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QFrame, QGridLayout, QLabel, QVBoxLayout, QWidget

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

        grid = QGridLayout()
        grid.setSpacing(16)
        cards = [c for c in ALL_FEATURE_CARDS if SHOW_SALES or c[0] != "sales"]
        for index, (key, name, desc) in enumerate(cards):
            grid.addWidget(self._card(key, name, desc), index // 2, index % 2)
        layout.addLayout(grid)
        layout.addStretch()
        layout.addWidget(muted(f"版本 {APP_VERSION}  ·  本地离线处理，不上传文件"))

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
        btn.clicked.connect(lambda *_, feature_key=key: self.open_feature.emit(feature_key))
        inner.addWidget(btn, alignment=Qt.AlignmentFlag.AlignLeft)
        return box
