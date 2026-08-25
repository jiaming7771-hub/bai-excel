from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QListWidget

from core.excel_merge import MergeResult, merge_excels
from ui.feature_page import FeaturePage
from ui.widgets import DropZone, PrimaryButton, SecondaryButton
from utils.error_handler import AppError
from utils.file_utils import collect_excel_files, open_file


class MergePage(FeaturePage):
    def __init__(self, parent=None) -> None:
        super().__init__(
            "Excel 批量合并",
            "请选择需要合并的 Excel。列名相近会自动对齐；干净结果可外传，列对齐详情在「处理报告」。",
            parent,
        )
        self.files: list[Path] = []
        self.output_path: Path | None = None
        self.report_path: Path | None = None

        self.drop = DropZone("将 Excel 文件拖到这里，也可以选择文件或文件夹")
        self.drop.setMinimumHeight(110)
        self.drop.files_dropped.connect(self.add_paths)
        self.layout_box.addWidget(self.drop)

        buttons = QHBoxLayout()
        self.pick_files_btn = SecondaryButton("选择文件")
        self.pick_folder_btn = SecondaryButton("选择文件夹")
        self.start_btn = PrimaryButton("开始合并")
        self.pick_files_btn.clicked.connect(self.pick_files)
        self.pick_folder_btn.clicked.connect(self.pick_folder)
        self.start_btn.clicked.connect(self.start)
        buttons.addWidget(self.pick_files_btn)
        buttons.addWidget(self.pick_folder_btn)
        buttons.addWidget(self.start_btn)
        buttons.addStretch()
        self.layout_box.addLayout(buttons)

        self.file_list = QListWidget()
        self.file_list.setMinimumHeight(150)
        self.layout_box.addWidget(QLabel("已选择的文件"))
        self.layout_box.addWidget(self.file_list)

        self.status.open_clicked.connect(self.open_output)
        self.status.report_clicked.connect(self.open_report)
        self.attach_status()

    def add_paths(self, paths: list[str]) -> None:
        try:
            found = collect_excel_files(paths)
        except AppError as exc:
            self.status.show_error(exc.title, exc.hint)
            return
        existing = {item.resolve() for item in self.files}
        for path in found:
            if path.resolve() not in existing:
                self.files.append(path)
        self.file_list.clear()
        for path in self.files:
            self.file_list.addItem(path.name)
        self.status.clear()

    def pick_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if files:
            self.add_paths(files)

    def pick_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择包含 Excel 的文件夹")
        if folder:
            self.add_paths([folder])

    def on_retry(self) -> None:
        self.pick_files()

    def start(self) -> None:
        if self.busy():
            return
        if not self.files:
            self.status.show_error("请先选择 Excel 文件", "可以一次选择多个文件，也可以把文件拖到上面的框里。")
            return
        self.set_busy(True)
        self.start_btn.setEnabled(False)
        paths = list(self.files)
        self.tasks.start(
            lambda cb: merge_excels(paths, progress_cb=cb),
            self.on_ok,
            self.on_fail,
            self.progress.update_progress,
        )

    def on_ok(self, result: MergeResult) -> None:
        self.start_btn.setEnabled(True)
        self.progress.finish()
        self.output_path = result.output_path
        self.report_path = result.report_path
        self.status.show_ok(
            "处理完成！",
            result.detail_text or f"合并完成，共处理 {result.file_count} 个文件，共 {result.row_count} 条数据。",
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
