from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QLabel,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app_config import APP_VERSION
from ui.widgets import PrimaryButton, muted

# 销售汇总先不放首页：透视一下就能做，卖点不够强
SHOW_SALES = False

FEATURE_CARD_HEIGHT = 152

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
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(36, 28, 36, 20)
        layout.setSpacing(20)

        header = QVBoxLayout()
        header.setSpacing(6)
        title = QLabel("Excel小工具箱")
        title.setObjectName("appTitle")
        subtitle = QLabel("批量处理 Excel，让重复工作一键完成")
        subtitle.setObjectName("appSubtitle")
        header.addWidget(title)
        header.addWidget(subtitle)
        layout.addLayout(header)

        layout.addWidget(
            muted(
                "打开后选择文件，点一下按钮就能完成原来要大量复制粘贴的工作。"
                "全程在您的电脑本地处理，不联网，也不上传文件。"
            )
        )
        layout.addWidget(self._whats_new())

        grid_host = QWidget()
        grid = QGridLayout(grid_host)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setSpacing(14)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        grid.setAlignment(Qt.AlignmentFlag.AlignTop)

        cards = [c for c in ALL_FEATURE_CARDS if SHOW_SALES or c[0] != "sales"]
        for index, (key, name, desc) in enumerate(cards):
            grid.addWidget(self._card(key, name, desc), index // 2, index % 2)

        layout.addWidget(grid_host)
        layout.addSpacing(4)
        footer = muted(f"版本 {APP_VERSION}  ·  本地离线处理，不上传文件")
        footer.setObjectName("homeFooter")
        layout.addWidget(footer)

        scroll.setWidget(content)
        outer.addWidget(scroll)

    def _whats_new(self) -> QFrame:
        box = QFrame()
        box.setObjectName("highlightBox")
        inner = QVBoxLayout(box)
        inner.setContentsMargins(16, 12, 16, 12)
        inner.setSpacing(4)
        heading = QLabel(f"本版更新（v{APP_VERSION}）")
        heading.setObjectName("highlightTitle")
        body = muted(
            "合并对齐列名并标注来源 · 拆分支持工作表/行数 · 清洗可组合去重 · 新增两表对比"
        )
        body.setObjectName("highlightBody")
        inner.addWidget(heading)
        inner.addWidget(body)
        return box

    def _card(self, key: str, name: str, desc: str) -> QFrame:
        box = QFrame()
        box.setObjectName("featureCard")
        box.setFixedHeight(FEATURE_CARD_HEIGHT)
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        inner = QVBoxLayout(box)
        inner.setContentsMargins(18, 16, 18, 16)
        inner.setSpacing(0)

        heading = QLabel(name)
        heading.setObjectName("featureTitle")
        inner.addWidget(heading)

        description = muted(desc)
        description.setObjectName("featureDesc")
        description.setFixedHeight(40)
        inner.addWidget(description)

        inner.addSpacing(10)
        btn = PrimaryButton("立即使用")
        btn.clicked.connect(lambda *_, feature_key=key: self.open_feature.emit(feature_key))
        inner.addWidget(btn, alignment=Qt.AlignmentFlag.AlignLeft)
        return box
