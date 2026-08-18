from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from app_config import ensure_output_dir
from utils.excel_utils import read_excel, write_excel
from utils.error_handler import AppError


@dataclass
class CleanResult:
    output_path: Path
    original_rows: int
    cleaned_rows: int
    deleted_rows: int


def _is_blank_row(row: pd.Series) -> bool:
    values = row.astype(str).str.strip()
    return bool(((values == "") | (values.str.lower() == "nan") | (values.str.lower() == "none")).all())


def clean_excel(
    path: Path,
    drop_duplicates: bool = False,
    drop_blank_rows: bool = False,
    dedupe_column: str | None = None,
    output_name: str = "cleaned_result.xlsx",
    progress_cb=None,
) -> CleanResult:
    if not any([drop_duplicates, drop_blank_rows, dedupe_column]):
        raise AppError(
            "请先选择清洗方式",
            "至少勾选一项：删除完全重复行、删除空白行，或按指定字段去重。",
        )

    if progress_cb:
        progress_cb(10, "正在读取 Excel")
    df = read_excel(path)
    original = len(df)

    if drop_blank_rows:
        if progress_cb:
            progress_cb(35, "正在删除空白行")
        mask = df.apply(_is_blank_row, axis=1)
        df = df.loc[~mask].copy()

    if drop_duplicates:
        if progress_cb:
            progress_cb(55, "正在删除完全重复行")
        df = df.drop_duplicates(keep="first")

    if dedupe_column:
        if progress_cb:
            progress_cb(75, f"正在按 {dedupe_column} 去重")
        if dedupe_column not in df.columns:
            raise AppError("找不到去重字段", "请选择表格里实际存在的字段，例如手机号、姓名或订单号。")
        df = df.drop_duplicates(subset=[dedupe_column], keep="first")

    df = df.reset_index(drop=True)
    output = ensure_output_dir() / output_name
    if progress_cb:
        progress_cb(90, "正在保存结果")
    write_excel(df, output)
    cleaned = len(df)
    if progress_cb:
        progress_cb(100, "处理完成！")
    return CleanResult(
        output_path=output,
        original_rows=original,
        cleaned_rows=cleaned,
        deleted_rows=max(0, original - cleaned),
    )
