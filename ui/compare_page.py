from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel

from core.excel_compare import CompareResult, compare_excels
from ui.feature_page import FeaturePage
from ui.widgets import DropZone, FieldSelect, PrimaryButton, SecondaryButton
from utils.excel_utils import list_columns, read_excel
from utils.error_handler import AppError, friendly_error
from utils.file_utils import collect_excel_files, open_file


class ComparePage(FeaturePage):
    def __init__(self, parent=None) -> None:
        super().__init__(
            "Excel 两表对比",
            "选两份表和主键（如手机号/订单号），自动找出：仅在 A、仅在 B、两边都有但字段变了。",
            parent,
        )
        self.path_a: Path | None = None
        self.path_b: Path | None = None
        self.output_path: Path | None = None
        self.report_path: Path | None = None
        self.columns: list[str] = []

        self.drop = DropZone("把两个 Excel 拖到这里（先 A 后 B，也可分别点选）")
        self.drop.setMinimumHeight(100)
        self.drop.files_dropped.connect(self.add_paths)
        self.layout_box.addWidget(self.drop)

        buttons = QHBoxLayout()
        self.pick_a_btn = SecondaryButton("选择表 A")
        self.pick_b_btn = SecondaryButton("选择表 B")
        self.start_btn = PrimaryButton("开始对比")
        self.pick_a_btn.clicked.connect(lambda: self.pick_file("a"))
        self.pick_b_btn.clicked.connect(lambda: self.pick_file("b"))
        self.start_btn.clicked.connect(self.start)
        buttons.addWidget(self.pick_a_btn)
        buttons.addWidget(self.pick_b_btn)
        buttons.addWidget(self.start_btn)
        buttons.addStretch()
        self.layout_box.addLayout(buttons)

        self.file_label = QLabel("还没有选择文件")
        self.file_label.setObjectName("fileInfo")
        self.layout_box.addWidget(self.file_label)

        self.key1 = FieldSelect("对比主键 1（必选）")
        self.key2 = FieldSelect("对比主键 2（可选，做组合键）")
        self.layout_box.addWidget(self.key1)
        self.layout_box.addWidget(self.key2)

        self.status.open_clicked.connect(self.open_output)
        self.status.report_clicked.connect(self.open_report)
        self.attach_status()

    def add_paths(self, paths: list[str]) -> None:
        try:
            files = collect_excel_files(paths)
        except AppError as exc:
            self.status.show_error(exc.title, exc.hint)
            return
        if len(files) >= 1:
            self.load_file(files[0], "a")
        if len(files) >= 2:
            self.load_file(files[1], "b")

    def pick_file(self, which: str) -> None:
        path, _ = QFileDialog.getOpenFileName(self, f"选择表 {which.upper()}", "", "Excel 文件 (*.xlsx *.xls)")
        if path:
            self.load_file(Path(path), which)

    def on_retry(self) -> None:
        self.pick_file("a")

    def _refresh_label(self) -> None:
        a = self.path_a.name if self.path_a else "未选"
        b = self.path_b.name if self.path_b else "未选"
        self.file_label.setText(f"表 A：{a}\n表 B：{b}")

    def _refresh_keys(self) -> None:
        if not self.columns:
            return
        preferred = next((c for c in self.columns if any(h in str(c) for h in ("手机", "订单", "编号", "邮箱"))), None)
        self.key1.set_fields(self.columns, preferred or self.columns[0])
        self.key2.set_fields(self.columns, None)

    def load_file(self, path: Path, which: str) -> None:
        try:
            df = read_excel(path)
        except Exception as exc:
            err = friendly_error(exc)
            self.status.show_error(err.title, err.hint)
            return
        if which == "a":
            self.path_a = path
        else:
            self.path_b = path

        # 主键选项取两表列交集；只有一张时用该表列
        try:
            cols_a = list_columns(read_excel(self.path_a)) if self.path_a else []
            cols_b = list_columns(read_excel(self.path_b)) if self.path_b else []
        except Exception as exc:
            err = friendly_error(exc)
            self.status.show_error(err.title, err.hint)
            return
        if cols_a and cols_b:
            self.columns = [c for c in cols_a if c in set(cols_b)]
            if not self.columns:
                self.status.show_error("两表没有同名列", "请确认两边表头一致，至少有一个共同字段可作主键。")
                self._refresh_label()
                return
        else:
            self.columns = cols_a or cols_b

        self._refresh_label()
        self._refresh_keys()
        self.status.clear()

    def start(self) -> None:
        if self.busy():
            return
        if not self.path_a or not self.path_b:
            self.status.show_error("请先选择两份 Excel", "分别选择表 A 和表 B，或一次拖入两个文件。")
            return
        key1 = self.key1.value()
        key2 = self.key2.value()
        if not key1:
            self.status.show_error("请选择对比主键", "例如手机号、订单号。")
            return
        keys = [key1]
        if key2 and key2 != key1:
            keys.append(key2)

        self.set_busy(True)
        self.start_btn.setEnabled(False)
        path_a = self.path_a
        path_b = self.path_b
        self.tasks.start(
            lambda cb: compare_excels(path_a, path_b, keys, progress_cb=cb),
            self.on_ok,
            self.on_fail,
            self.progress.update_progress,
        )

    def on_ok(self, result: CompareResult) -> None:
        self.start_btn.setEnabled(True)
        self.progress.finish()
        self.output_path = result.output_path
        self.report_path = result.report_path
        self.status.show_ok(
            "对比完成！",
            (
                f"{result.detail_text}\n\n"
                "结果文件为差异清单，可外传；处理报告含分表详情与标色。"
            ),
            result.output_path,
            report=result.report_path,
        )

    def on_fail(self, title: str, hint: str) -> None:
        self.start_btn.setEnabled(True)
        self.progress.setVisible(False)
        self.status.show_error(title, hint)

    def open_output(self) -> None:
        if self.output_path:
            open_file(self.output_path)

    def open_report(self) -> None:
        if self.report_path:
            open_file(self.report_path)
