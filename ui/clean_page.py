from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QComboBox, QFileDialog, QHBoxLayout, QLabel

from core.excel_clean import CleanResult, clean_excel
from ui.feature_page import FeaturePage
from ui.widgets import DropZone, FieldSelect, PrimaryButton, SecondaryButton, checkbox
from utils.excel_utils import list_columns, read_excel
from utils.error_handler import AppError, friendly_error
from utils.file_utils import collect_excel_files, open_file


COMMON_FIELDS = ["手机号", "姓名", "身份证号", "邮箱", "订单号", "客户编号"]
DATE_HINTS = ("日期", "下单", "时间", "入职", "更新", "创建")


class CleanPage(FeaturePage):
    def __init__(self, parent=None) -> None:
        super().__init__(
            "数据清洗",
            "空白/重复可删；字段去重支持组合键、保留最新；能修的自动修并标黄。",
            parent,
        )
        self.file_path: Path | None = None
        self.output_path: Path | None = None
        self.report_path: Path | None = None
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
        self.file_label.setObjectName("fileInfo")
        self.layout_box.addWidget(self.file_label)

        self.opt_dup = checkbox("删除完全重复行")
        self.opt_blank = checkbox("删除空白行")
        self.opt_shifted = checkbox("尝试归位窜行（能修标黄，修不了待核对）")
        self.opt_phone = checkbox("尝试修复手机号（修不了的待核对）")
        self.opt_email = checkbox("尝试修复邮箱（修不了的待核对）")
        self.opt_date = checkbox("尝试修复日期（修不了的待核对）")
        self.opt_required = checkbox("关键字段缺失标出待核对（姓名/手机）")
        self.opt_field = checkbox("按指定字段去重")
        for opt in (
            self.opt_dup,
            self.opt_blank,
            self.opt_shifted,
            self.opt_phone,
            self.opt_email,
            self.opt_date,
            self.opt_required,
        ):
            opt.setChecked(True)
            self.layout_box.addWidget(opt)

        self.opt_field.setChecked(True)
        self.layout_box.addWidget(self.opt_field)

        self.field = FieldSelect("去重字段 1（必选，建议手机号）")
        self.field2 = FieldSelect("去重字段 2（可选，做组合键，如姓名）")
        self.layout_box.addWidget(self.field)
        self.layout_box.addWidget(self.field2)

        self.layout_box.addWidget(QLabel("重复时保留哪一条？"))
        self.keep_mode = QComboBox()
        self.keep_mode.addItem("保留第一条", "first")
        self.keep_mode.addItem("保留最后一条", "last")
        self.keep_mode.addItem("按日期保留最新", "latest")
        self.keep_mode.currentIndexChanged.connect(self._sync_keep_ui)
        self.layout_box.addWidget(self.keep_mode)

        self.latest_field = FieldSelect("日期字段（按最新保留时使用）")
        self.layout_box.addWidget(self.latest_field)

        self.opt_field.toggled.connect(self._sync_dedupe_ui)
        self.status.open_clicked.connect(self.open_output)
        self.status.report_clicked.connect(self.open_report)
        self.attach_status()
        self._sync_dedupe_ui()

    def _sync_dedupe_ui(self) -> None:
        enabled = self.opt_field.isChecked()
        self.field.setEnabled(enabled)
        self.field2.setEnabled(enabled)
        self.keep_mode.setEnabled(enabled)
        self._sync_keep_ui()

    def _sync_keep_ui(self) -> None:
        enabled = self.opt_field.isChecked() and str(self.keep_mode.currentData() or "") == "latest"
        self.latest_field.setEnabled(enabled)
        self.latest_field.setVisible(True)

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

    def _preferred_fields(self) -> list[str]:
        return [name for name in COMMON_FIELDS if name in self.columns] + [
            col for col in self.columns if col not in COMMON_FIELDS
        ]

    def _refresh_field_choices(self) -> None:
        preferred = self._preferred_fields()
        selected = next((c for c in self.columns if "手机" in str(c)), None)
        if not selected:
            selected = next((name for name in COMMON_FIELDS if name in self.columns), preferred[0] if preferred else None)
        self.field.set_fields(preferred or self.columns, selected)
        self.field2.set_fields(preferred or self.columns, None)

        date_cols = [c for c in self.columns if any(h in str(c) for h in DATE_HINTS)]
        date_selected = date_cols[0] if date_cols else (preferred[0] if preferred else None)
        self.latest_field.set_fields(self.columns, date_selected)

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
        self._refresh_field_choices()
        self.status.clear()

    def start(self) -> None:
        if self.busy():
            return
        if not self.file_path:
            self.status.show_error("请先选择 Excel 文件", "把文件拖进来，或点击“选择文件”。")
            return

        dedupe_columns: list[str] | None = None
        dedupe_keep = "first"
        dedupe_latest_column = None
        if self.opt_field.isChecked():
            col1 = self.field.value()
            col2 = self.field2.value()
            if not col1:
                self.status.show_error("请选择去重字段", "例如选择手机号；也可再选姓名做组合键。")
                return
            dedupe_columns = [col1]
            if col2 and col2 != col1:
                dedupe_columns.append(col2)
            dedupe_keep = str(self.keep_mode.currentData() or "first")
            if dedupe_keep == "latest":
                dedupe_latest_column = self.latest_field.value()
                if not dedupe_latest_column:
                    self.status.show_error("请选择日期字段", "「按日期保留最新」需要选一列日期/更新时间。")
                    return

        self.set_busy(True)
        self.start_btn.setEnabled(False)
        path = self.file_path
        options = dict(
            drop_duplicates=self.opt_dup.isChecked(),
            drop_blank_rows=self.opt_blank.isChecked(),
            dedupe_columns=dedupe_columns,
            dedupe_keep=dedupe_keep,
            dedupe_latest_column=dedupe_latest_column,
            check_shifted_rows=self.opt_shifted.isChecked(),
            fix_phone=self.opt_phone.isChecked(),
            fix_email=self.opt_email.isChecked(),
            fix_dates=self.opt_date.isChecked(),
            check_required_fields=self.opt_required.isChecked(),
        )
        self.tasks.start(
            lambda cb: clean_excel(path, progress_cb=cb, **options),
            self.on_ok,
            self.on_fail,
            self.progress.update_progress,
        )

    def on_ok(self, result: CleanResult) -> None:
        self.start_btn.setEnabled(True)
        self.progress.finish()
        self.output_path = result.output_path
        self.report_path = result.report_path
        self.status.show_ok(
            "清洗完成！",
            (
                f"{result.detail_text}\n"
                f"{result.quality_summary}\n\n"
                "结果文件可直接外传；处理报告含核对视图、已修复、待核对、已删除。"
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
