from __future__ import annotations

from PySide6.QtWidgets import QMainWindow, QStackedWidget, QVBoxLayout, QWidget

from app_config import APP_NAME
from ui.clean_page import CleanPage
from ui.compare_page import ComparePage
from ui.home_page import HomePage
from ui.merge_page import MergePage
from ui.sales_page import SalesPage
from ui.split_page import SplitPage
from ui.styles import APP_STYLE


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(920, 740)
        self.setMinimumSize(820, 640)
        self.setStyleSheet(APP_STYLE)

        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)

        self.stack = QStackedWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.stack)

        self.home = HomePage()
        self.merge = MergePage()
        self.split = SplitPage()
        self.clean = CleanPage()
        self.compare = ComparePage()
        self.sales = SalesPage()

        self.pages = {
            "home": self.home,
            "merge": self.merge,
            "split": self.split,
            "clean": self.clean,
            "compare": self.compare,
            "sales": self.sales,
        }
        for page in self.pages.values():
            self.stack.addWidget(page)

        self.home.open_feature.connect(self.show_page)
        self.merge.back_requested.connect(lambda: self.show_page("home"))
        self.split.back_requested.connect(lambda: self.show_page("home"))
        self.clean.back_requested.connect(lambda: self.show_page("home"))
        self.compare.back_requested.connect(lambda: self.show_page("home"))
        self.sales.back_requested.connect(lambda: self.show_page("home"))
        self.setAcceptDrops(True)

    def show_page(self, key: str) -> None:
        page = self.pages.get(key, self.home)
        self.stack.setCurrentWidget(page)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile()]
        if not paths:
            return
        current = self.stack.currentWidget()
        if current is self.home:
            self.show_page("merge")
            self.merge.add_paths(paths)
        elif hasattr(current, "add_paths"):
            current.add_paths(paths)
        event.acceptProposedAction()
