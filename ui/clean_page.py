from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QComboBox, QFileDialog, QHBoxLayout, QLabel, QVBoxLayout, QWidget

from core.excel_clean import CleanResult, clean_excel
from ui.feature_page import FeaturePage
from ui.widgets import DropZone, FieldSelect, PrimaryButton, SecondaryButton, checkbox, muted
from utils.excel_utils import list_columns, read_excel
from utils.error_handler import AppError, friendly_error
from utils.file_utils import collect_excel_files, open_file


COMMON_FIELDS = ["手机号", "姓名", "身份证号", "邮箱", "订单号", "客户编号"]

# 场景预设：少勾选、按用途一键清洗
PRESETS = {
    "customer": {
        "label": "联系人 / 客户表（推荐）",
        "hint": "窜行归位 + 手机/邮箱/日期修复 + 关键字段缺失标出 + 按手机去重",
        "drop_duplicates": True,
        "drop_blank_rows": True,
        "check_shifted_rows": True,
        "fix_phone": True,
        "fix_email": True,
        "fix_dates": True,
        "check_required_fields": True,
        "use_field_dedupe": True,
        "prefer_dedupe": "手机号",
    },
    "order": {
        "label": "订单 / 明细表",
        "hint": "空白与重复清理 + 日期规整；有手机/邮箱也会检查",
        "drop_duplicates": True,
        "drop_blank_rows": True,
        "check_shifted_rows": False,
        "fix_phone": True,
        "fix_email": True,
        "fix_dates": True,
        "check_required_fields": False,
        "use_field_dedupe": True,
        "prefer_dedupe": "订单号",
    },
    "custom": {
        "label": "自定义规则",
        "hint": "自己勾选下面规则（适合特殊表）",
        "drop_duplicates": True,
        "drop_blank_rows": True,
        "check_shifted_rows": True,
        "fix_phone": True,
        "fix_email": True,
        "fix_dates": True,
        "check_required_fields": True,
        "use_field_dedupe": False,
        "prefer_dedupe": "手机号",
    },
}


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

        self.layout_box.addWidget(muted("使用场景"))
        self.preset = QComboBox()
        for key, meta in PRESETS.items():
            self.preset.addItem(meta["label"], key)
        self.preset.currentIndexChanged.connect(self.apply_preset)
        self.layout_box.addWidget(self.preset)

        self.preset_hint = muted(PRESETS["customer"]["hint"])
        self.layout_box.addWidget(self.preset_hint)

        self.field = FieldSelect("按字段去重（默认保留第一条，建议选手机号/订单号）")
        self.layout_box.addWidget(self.field)

        self.advanced_box = QWidget()
        adv = QVBoxLayout(self.advanced_box)
        adv.setContentsMargins(0, 8, 0, 0)
        adv.setSpacing(8)
        adv.addWidget(muted("高级规则（仅自定义场景需要改）"))
        self.opt_dup = checkbox("删除完全重复行")
        self.opt_blank = checkbox("删除空白行")
        self.opt_shifted = checkbox("尝试归位窜行")
        self.opt_phone = checkbox("尝试修复手机号")
        self.opt_email = checkbox("尝试修复邮箱")
        self.opt_date = checkbox("尝试修复日期")
        self.opt_required = checkbox("关键字段缺失标出待核对（姓名/手机）")
        self.opt_field = checkbox("启用按字段去重")
        for opt in (
            self.opt_dup,
            self.opt_blank,
            self.opt_shifted,
            self.opt_phone,
            self.opt_email,
            self.opt_date,
            self.opt_required,
            self.opt_field,
        ):
            adv.addWidget(opt)
        self.layout_box.addWidget(self.advanced_box)

        self.opt_field.toggled.connect(self._sync_field_enabled)
        self.status.open_clicked.connect(self.open_output)
        self.attach_status()
        self.apply_preset()

    def _sync_field_enabled(self) -> None:
        self.field.setEnabled(self.opt_field.isChecked())

    def apply_preset(self) -> None:
        key = str(self.preset.currentData() or "customer")
        meta = PRESETS[key]
        self.preset_hint.setText(meta["hint"])
        self.opt_dup.setChecked(meta["drop_duplicates"])
        self.opt_blank.setChecked(meta["drop_blank_rows"])
        self.opt_shifted.setChecked(meta["check_shifted_rows"])
        self.opt_phone.setChecked(meta["fix_phone"])
        self.opt_email.setChecked(meta["fix_email"])
        self.opt_date.setChecked(meta["fix_dates"])
        self.opt_required.setChecked(meta["check_required_fields"])
        self.opt_field.setChecked(meta["use_field_dedupe"])
        custom = key == "custom"
        self.advanced_box.setEnabled(custom)
        self._sync_field_enabled()
        if self.columns:
            self._refresh_field_choices(meta.get("prefer_dedupe"))

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

    def _refresh_field_choices(self, prefer: str | None = None) -> None:
        preferred = [name for name in COMMON_FIELDS if name in self.columns] + [
            col for col in self.columns if col not in COMMON_FIELDS
        ]
        prefer = prefer or PRESETS[str(self.preset.currentData() or "customer")].get("prefer_dedupe")
        selected = None
        if prefer:
            selected = next((c for c in self.columns if prefer in str(c)), None)
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
