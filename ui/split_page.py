from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel

from core.excel_split import SplitResult, split_excel
from ui.feature_page import FeaturePage
from ui.widgets import DropZone, FieldSelect, PrimaryButton, SecondaryButton
from utils.excel_utils import list_columns, read_excel
from utils.error_handler import AppError, friendly_error
from utils.file_utils import collect_excel_files, open_folder


class SplitPage(FeaturePage):
    def __init__(self, parent=None) -> None:
        super().__init__(
            "Excel 数据拆分",
            "选择一个 Excel，再选择要拆分的字段。例如按“部门”拆分后，每个部门会生成一个独立的 Excel。",
            parent,
        )
        self.file_path: Path | None = None
        self.output_dir: Path | None = None

        self.drop = DropZone("将 Excel 文件拖到这里")
        self.drop.setMinimumHeight(110)
        self.drop.files_dropped.connect(self.add_paths)
        self.layout_box.addWidget(self.drop)

        buttons = QHBoxLayout()
        self.pick_btn = SecondaryButton("选择文件")
        self.start_btn = PrimaryButton("开始拆分")
        self.pick_btn.clicked.connect(self.pick_file)
        self.start_btn.clicked.connect(self.start)
        buttons.addWidget(self.pick_btn)
        buttons.addWidget(self.start_btn)
        buttons.addStretch()
        self.layout_box.addLayout(buttons)

        self.file_label = QLabel("还没有选择文件")
        self.layout_box.addWidget(self.file_label)

        self.field = FieldSelect("按照哪个字段拆分？")
        self.layout_box.addWidget(self.field)

        self.status.open_clicked.connect(self.open_output)
        self.attach_status()

    def add_paths(self, paths: list[str]) -> None:
        try:
            files = collect_excel_files(paths)
        except AppError as exc:
            self.status.show_error(exc.title, exc.hint)
            return
        self.load_file(files[0])

    def pick_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if path:
            self.load_file(Path(path))

    def on_retry(self) -> None:
        self.pick_file()

    def load_file(self, path: Path) -> None:
        try:
            df = read_excel(path)
        except Exception as exc:
            err = friendly_error(exc)
            self.status.show_error(err.title, err.hint)
            return
        self.file_path = path
        self.file_label.setText(f"文件：{path.name}    数据量：{len(df)} 行")
        self.field.set_fields(list_columns(df))
        self.status.clear()

    def start(self) -> None:
        if self.busy():
            return
        if not self.file_path:
            self.status.show_error("请先选择 Excel 文件", "把文件拖进来，或点击“选择文件”。")
            return
        column = self.field.value()
        if not column:
            self.status.show_error("请选择拆分字段", "例如：部门、城市、销售人员。")
            return
        self.set_busy(True)
        self.start_btn.setEnabled(False)
        path = self.file_path
        self.tasks.start(
            lambda cb: split_excel(path, column, progress_cb=cb),
            self.on_ok,
            self.on_fail,
            self.progress.update_progress,
        )

    def on_ok(self, result: SplitResult) -> None:
        self.start_btn.setEnabled(True)
        self.progress.finish()
        self.output_dir = result.output_dir
        self.status.show_ok(
            "处理完成！",
            f"拆分完成，共生成 {result.file_count} 个文件。",
            result.output_dir,
        )

    def on_fail(self, title: str, hint: str) -> None:
        self.start_btn.setEnabled(True)
        self.progress.setVisible(False)
        self.status.show_error(title, hint)

    def open_output(self) -> None:
        if self.output_dir:
            open_folder(self.output_dir)
