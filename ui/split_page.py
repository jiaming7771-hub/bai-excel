from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QComboBox, QFileDialog, QHBoxLayout, QLabel, QSpinBox

from core.excel_split import SplitResult, split_excel
from ui.feature_page import FeaturePage
from ui.widgets import DropZone, FieldSelect, PrimaryButton, SecondaryButton
from utils.excel_utils import list_columns, list_sheet_names, read_excel
from utils.error_handler import AppError, friendly_error
from utils.file_utils import collect_excel_files, open_file, open_folder


class SplitPage(FeaturePage):
    def __init__(self, parent=None) -> None:
        super().__init__(
            "Excel 数据拆分",
            "支持三种方式：按字段拆、按工作表拆、按行数拆。结果会生成多个 Excel 文件。",
            parent,
        )
        self.file_path: Path | None = None
        self.output_dir: Path | None = None
        self.report_path: Path | None = None
        self.columns: list[str] = []

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
        self.file_label.setObjectName("fileInfo")
        self.layout_box.addWidget(self.file_label)

        self.layout_box.addWidget(QLabel("拆分方式"))
        self.mode = QComboBox()
        self.mode.addItem("按字段拆分（如部门）", "column")
        self.mode.addItem("按工作表拆分（每个 Sheet 一个文件）", "sheet")
        self.mode.addItem("按行数拆分（大表切成多份）", "rows")
        self.mode.currentIndexChanged.connect(self._sync_mode_ui)
        self.layout_box.addWidget(self.mode)

        self.field = FieldSelect("按照哪个字段拆分？")
        self.layout_box.addWidget(self.field)

        self.rows_label = QLabel("每个文件多少行？")
        self.rows_spin = QSpinBox()
        self.rows_spin.setRange(1, 1_000_000)
        self.rows_spin.setValue(1000)
        self.rows_spin.setSingleStep(100)
        self.layout_box.addWidget(self.rows_label)
        self.layout_box.addWidget(self.rows_spin)

        self.status.open_clicked.connect(self.open_output)
        self.status.report_clicked.connect(self.open_report)
        self.attach_status()
        self._sync_mode_ui()

    def _mode_value(self) -> str:
        return str(self.mode.currentData() or "column")

    def _sync_mode_ui(self) -> None:
        mode = self._mode_value()
        self.field.setVisible(mode == "column")
        self.rows_label.setVisible(mode == "rows")
        self.rows_spin.setVisible(mode == "rows")

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
            sheets = list_sheet_names(path)
            df = read_excel(path)
        except Exception as exc:
            err = friendly_error(exc)
            self.status.show_error(err.title, err.hint)
            return
        self.file_path = path
        self.columns = list_columns(df)
        self.file_label.setText(f"文件：{path.name}    数据量：{len(df)} 行    工作表：{len(sheets)} 个")
        self.field.set_fields(self.columns)
        self.status.clear()

    def start(self) -> None:
        if self.busy():
            return
        if not self.file_path:
            self.status.show_error("请先选择 Excel 文件", "把文件拖进来，或点击“选择文件”。")
            return
        mode = self._mode_value()
        column = self.field.value() if mode == "column" else None
        if mode == "column" and not column:
            self.status.show_error("请选择拆分字段", "例如：部门、城市、销售人员。")
            return
        rows_per_file = int(self.rows_spin.value())
        self.set_busy(True)
        self.start_btn.setEnabled(False)
        path = self.file_path
        self.tasks.start(
            lambda cb: split_excel(
                path,
                column,
                mode=mode,
                rows_per_file=rows_per_file,
                progress_cb=cb,
            ),
            self.on_ok,
            self.on_fail,
            self.progress.update_progress,
        )

    def on_ok(self, result: SplitResult) -> None:
        self.start_btn.setEnabled(True)
        self.progress.finish()
        self.output_dir = result.output_dir
        self.report_path = result.report_path
        detail = result.detail_text or f"拆分完成，共生成 {result.file_count} 个文件。"
        self.status.show_ok(
            "处理完成！",
            detail,
            result.output_dir,
            report=result.report_path,
        )

    def on_fail(self, title: str, hint: str) -> None:
        self.start_btn.setEnabled(True)
        self.progress.setVisible(False)
        self.status.show_error(title, hint)

    def open_output(self) -> None:
        if self.output_dir:
            open_folder(self.output_dir)

    def open_report(self) -> None:
        if self.report_path:
            open_file(self.report_path)
