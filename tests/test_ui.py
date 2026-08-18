from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QPushButton

from ui.home_page import SHOW_SALES, HomePage
from ui.main_window import MainWindow


def test_main_window_builds():
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    assert window.windowTitle() == "Excel小工具箱"
    assert window.stack.count() == 5
    app.processEvents()


def test_home_cards_open_the_right_feature():
    app = QApplication.instance() or QApplication([])
    page = HomePage()
    seen: list[str] = []
    page.open_feature.connect(seen.append)
    buttons = page.findChildren(QPushButton)
    expected = ["merge", "split", "clean", "sales"] if SHOW_SALES else ["merge", "split", "clean"]
    assert [btn.text() for btn in buttons] == ["立即使用"] * len(expected)
    for btn in buttons:
        btn.click()
    assert seen == expected
    app.processEvents()
