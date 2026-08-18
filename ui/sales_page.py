from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel

from core.sales_report import SalesResult, auto_detect_columns, build_sales_report
from ui.feature_page import FeaturePage
from ui.widgets import DropZone, FieldSelect, PrimaryButton, SecondaryButton
from utils.excel_utils import list_columns, read_excel
from utils.error_handler import AppError, friendly_error
from utils.file_utils import collect_excel_files, open_folder


class SalesPage(FeaturePage):
    def __init__(self, parent=None) -> None:
        super().__init__(
            "销售数据汇总",
            "选择销售明细表，软件会尽量自动识别销售人员、销售额、数量、产品和日期。如果对不上，您可以手动选择。",
            parent,
        )
        self.file_path: Path | None = None
        self.output_path: Path | None = None

        self.drop = DropZone("将 Excel 文件拖到这里")
        self.drop.setMinimumHeight(110)
        self.drop.files_dropped.connect(self.add_paths)
        self.layout_box.addWidget(self.drop)

        buttons = QHBoxLayout()
        self.pick_btn = SecondaryButton("选择文件")
        self.start_btn = PrimaryButton("生成报表")
        self.pick_btn.clicked.connect(self.pick_file)
        self.start_btn.clicked.connect(self.start)
        buttons.addWidget(self.pick_btn)
        buttons.addWidget(self.start_btn)
        buttons.addStretch()
        self.layout_box.addLayout(buttons)

        self.file_label = QLabel("还没有选择文件")
        self.file_label.setObjectName("fileInfo")
        self.layout_box.addWidget(self.file_label)

        self.person = FieldSelect("销售人员字段")
        self.amount = FieldSelect("销售额字段")
        self.qty = FieldSelect("数量字段")
        self.product = FieldSelect("产品字段")
        self.date = FieldSelect("日期字段")
        for widget in (self.person, self.amount, self.qty, self.product, self.date):
            self.layout_box.addWidget(widget)

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
        columns = list_columns(df)
        mapping = auto_detect_columns(columns)
        self.file_label.setText(f"文件：{path.name}    数据量：{len(df)} 行")
        self.person.set_fields(columns, mapping["person"])
        self.amount.set_fields(columns, mapping["amount"])
        self.qty.set_fields(columns, mapping["qty"])
        self.product.set_fields(columns, mapping["product"])
        self.date.set_fields(columns, mapping["date"])
        self.status.clear()

    def start(self) -> None:
        if self.busy():
            return
        if not self.file_path:
            self.status.show_error("请先选择 Excel 文件", "把文件拖进来，或点击“选择文件”。")
            return
        person = self.person.value()
        amount = self.amount.value()
        if not person or not amount:
            self.status.show_error("请选择必要字段", "至少需要选择销售人员和销售额。")
            return
        self.set_busy(True)
        self.start_btn.setEnabled(False)
        path = self.file_path
        qty = self.qty.value() or None
        product = self.product.value() or None
        date = self.date.value() or None
        self.tasks.start(
            lambda cb: build_sales_report(
                path,
                person_col=person,
                amount_col=amount,
                qty_col=qty,
                product_col=product,
                date_col=date,
                progress_cb=cb,
            ),
            self.on_ok,
            self.on_fail,
            self.progress.update_progress,
        )

    def on_ok(self, result: SalesResult) -> None:
        self.start_btn.setEnabled(True)
        self.progress.finish()
        self.output_path = result.output_path
        extra = f"已汇总 {result.people_count} 位销售人员"
        if result.product_count:
            extra += f"、{result.product_count} 个产品"
        self.status.show_ok(
            "处理完成！",
            f"{extra}。报表已保存为 {result.output_path.name}",
            result.output_path,
        )

    def on_fail(self, title: str, hint: str) -> None:
        self.start_btn.setEnabled(True)
        self.progress.setVisible(False)
        self.status.show_error(title, hint)

    def open_output(self) -> None:
        if self.output_path:
            open_folder(self.output_path)
