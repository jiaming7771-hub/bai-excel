from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel

from core.excel_clean import CleanResult, clean_excel
from ui.feature_page import FeaturePage
from ui.widgets import DropZone, FieldSelect, PrimaryButton, SecondaryButton, checkbox
from utils.excel_utils import list_columns, read_excel
from utils.error_handler import AppError, friendly_error
from utils.file_utils import collect_excel_files, open_folder


COMMON_FIELDS = ["姓名", "手机号", "身份证号", "邮箱", "订单号", "客户编号"]


class CleanPage(FeaturePage):
    def __init__(self, parent=None) -> None:
        super().__init__(
            "重复数据清洗",
            "删除完全重复的行、空白行，或按手机号、姓名等字段只保留第一条。",
            parent,
        )
        self.file_path: Path | None = None
        self.output_path: Path | None = None
        self.columns: list[str] = []

        self.drop = DropZone("将 Excel 文件拖到这里")
        self.drop.setMinimumHeight(110)
        self.drop.files_dropped.connect(self.add_paths)
        self.layout_box.addWidget(self.drop)

        buttons = QHBoxLayout()
        self.pick_btn = SecondaryButton("选择文件")
        self.start_btn = PrimaryButton("开始清洗")
        self.pick_btn.clicked.connect(self.pick_file)
        self.start_btn.clicked.connect(self.start)
        buttons.addWidget(self.pick_btn)
        buttons.addWidget(self.start_btn)
        buttons.addStretch()
        self.layout_box.addLayout(buttons)

        self.file_label = QLabel("还没有选择文件")
        self.layout_box.addWidget(self.file_label)

        self.opt_dup = checkbox("删除完全重复行")
        self.opt_blank = checkbox("删除空白行")
        self.opt_field = checkbox("按指定字段去重")
        self.opt_dup.setChecked(True)
        self.opt_blank.setChecked(True)
        self.layout_box.addWidget(self.opt_dup)
        self.layout_box.addWidget(self.opt_blank)
        self.layout_box.addWidget(self.opt_field)

        self.field = FieldSelect("去重字段（默认保留第一条）")
        self.layout_box.addWidget(self.field)
        self.opt_field.toggled.connect(self.field.setEnabled)
        self.field.setEnabled(False)

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
        self.columns = list_columns(df)
        self.file_label.setText(f"文件：{path.name}\n数据量：{len(df)} 行")
        preferred = [name for name in COMMON_FIELDS if name in self.columns] + [
            col for col in self.columns if col not in COMMON_FIELDS
        ]
        selected = next((name for name in COMMON_FIELDS if name in self.columns), preferred[0] if preferred else None)
        self.field.set_fields(preferred or self.columns, selected)
        self.status.clear()

    def start(self) -> None:
        if self.busy():
            return
        if not self.file_path:
            self.status.show_error("请先选择 Excel 文件", "把文件拖进来，或点击“选择文件”。")
            return
        dedupe_column = self.field.value() if self.opt_field.isChecked() else None
        if self.opt_field.isChecked() and not dedupe_column:
            self.status.show_error("请选择去重字段", "例如选择手机号、姓名、订单号。")
            return
        self.set_busy(True)
        self.start_btn.setEnabled(False)
        path = self.file_path
        drop_dup = self.opt_dup.isChecked()
        drop_blank = self.opt_blank.isChecked()
        self.tasks.start(
            lambda cb: clean_excel(
                path,
                drop_duplicates=drop_dup,
                drop_blank_rows=drop_blank,
                dedupe_column=dedupe_column,
                progress_cb=cb,
            ),
            self.on_ok,
            self.on_fail,
            self.progress.update_progress,
        )

    def on_ok(self, result: CleanResult) -> None:
        self.start_btn.setEnabled(True)
        self.progress.finish()
        self.output_path = result.output_path
        self.status.show_ok(
            "处理完成！",
            f"原始数据：{result.original_rows} 行\n清洗后：{result.cleaned_rows} 行\n删除：{result.deleted_rows} 行",
            result.output_path,
        )

    def on_fail(self, title: str, hint: str) -> None:
        self.start_btn.setEnabled(True)
        self.progress.setVisible(False)
        self.status.show_error(title, hint)

    def open_output(self) -> None:
        if self.output_path:
            open_folder(self.output_path)
