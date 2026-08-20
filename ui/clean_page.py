from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel

from core.excel_clean import CleanResult, clean_excel
from ui.feature_page import FeaturePage
from ui.widgets import DropZone, FieldSelect, PrimaryButton, SecondaryButton, checkbox
from utils.excel_utils import list_columns, read_excel
from utils.error_handler import AppError, friendly_error
from utils.file_utils import collect_excel_files, open_file


COMMON_FIELDS = ["手机号", "姓名", "身份证号", "邮箱", "订单号", "客户编号"]


class CleanPage(FeaturePage):
    def __init__(self, parent=None) -> None:
        super().__init__(
            "数据清洗",
            "空白/重复可删；能修的自动修并标黄，修不了的进待核对。结果带清洗报告，方便抽查。",
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
        self.file_label.setObjectName("fileInfo")
        self.layout_box.addWidget(self.file_label)

        # 一套默认规则，全部默认勾选
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

        self.field = FieldSelect("去重字段（默认保留第一条，建议选手机号）")
        self.layout_box.addWidget(self.field)
        self.opt_field.toggled.connect(self.field.setEnabled)

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

    def _refresh_field_choices(self) -> None:
        preferred = [name for name in COMMON_FIELDS if name in self.columns] + [
            col for col in self.columns if col not in COMMON_FIELDS
        ]
        selected = next((c for c in self.columns if "手机" in str(c)), None)
        if not selected:
            selected = next((name for name in COMMON_FIELDS if name in self.columns), preferred[0] if preferred else None)
        self.field.set_fields(preferred or self.columns, selected)

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
        dedupe_column = self.field.value() if self.opt_field.isChecked() else None
        if self.opt_field.isChecked() and not dedupe_column:
            self.status.show_error("请选择去重字段", "例如选择手机号、姓名、订单号。")
            return
        self.set_busy(True)
        self.start_btn.setEnabled(False)
        path = self.file_path
        options = dict(
            drop_duplicates=self.opt_dup.isChecked(),
            drop_blank_rows=self.opt_blank.isChecked(),
            dedupe_column=dedupe_column,
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
        self.status.show_ok(
            "清洗完成！",
            (
                f"{result.detail_text}\n"
                f"{result.quality_summary}\n\n"
                "请打开结果，按这个顺序看：\n"
                "1）「清洗报告」——健康度与问题清单\n"
                "2）「已自动修复」——原值→新值，抽查对不对\n"
                "3）「待人工核对」——修不了的，需你处理\n"
                "4）「清洗结果」——黄=已修，橙=待核"
            ),
            result.output_path,
        )

    def on_fail(self, title: str, hint: str) -> None:
        self.start_btn.setEnabled(True)
        self.progress.setVisible(False)
        self.status.show_error(title, hint)

    def open_output(self) -> None:
        if self.output_path:
            open_file(self.output_path)
