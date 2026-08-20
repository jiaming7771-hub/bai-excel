from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from app_config import ensure_output_dir
from utils.excel_utils import list_sheet_names, read_excel, write_excel
from utils.error_handler import AppError
from utils.file_utils import sanitize_filename


@dataclass
class SplitResult:
    output_dir: Path
    file_count: int
    row_count: int
    mode: str = "column"


def _unique_filename(base: str, used_names: dict[str, int]) -> str:
    safe = sanitize_filename(base, fallback="空值")
    count = used_names.get(safe, 0) + 1
    used_names[safe] = count
    return f"{safe}.xlsx" if count == 1 else f"{safe}_{count}.xlsx"


def split_excel(
    path: Path,
    column: str | None = None,
    *,
    mode: str = "column",
    rows_per_file: int = 1000,
    progress_cb=None,
) -> SplitResult:
    """mode: column | sheet | rows"""
    mode = (mode or "column").strip().lower()
    if mode not in {"column", "sheet", "rows"}:
        raise AppError("不支持的拆分方式", "请选择：按字段、按工作表、或按行数。")

    output_dir = ensure_output_dir() / f"split_{path.stem}"
    output_dir.mkdir(parents=True, exist_ok=True)
    used_names: dict[str, int] = {}

    if mode == "sheet":
        return _split_by_sheets(path, output_dir, used_names, progress_cb)
    if mode == "rows":
        return _split_by_rows(path, output_dir, used_names, rows_per_file, progress_cb)
    if not column:
        raise AppError("请选择拆分字段", "例如：部门、城市、销售人员。")
    return _split_by_column(path, column, output_dir, used_names, progress_cb)


def _split_by_column(
    path: Path,
    column: str,
    output_dir: Path,
    used_names: dict[str, int],
    progress_cb,
) -> SplitResult:
    if progress_cb:
        progress_cb(5, "正在读取 Excel")
    df = read_excel(path)
    if column not in df.columns:
        raise AppError("找不到拆分字段", "请重新选择一个表格里实际存在的字段。")

    groups = list(df.groupby(column, dropna=False, sort=False))
    if not groups:
        raise AppError("没有可以拆分的数据", "请确认表格里至少有一行数据。")

    total = len(groups)
    for index, (value, group) in enumerate(groups, start=1):
        filename = _unique_filename(str(value), used_names)
        write_excel(group.reset_index(drop=True), output_dir / filename)
        if progress_cb:
            progress_cb(10 + int(index / total * 85), f"正在生成 {filename}")

    if progress_cb:
        progress_cb(100, "处理完成！")
    return SplitResult(output_dir=output_dir, file_count=len(groups), row_count=len(df), mode="column")


def _split_by_sheets(
    path: Path,
    output_dir: Path,
    used_names: dict[str, int],
    progress_cb,
) -> SplitResult:
    sheets = list_sheet_names(path)
    if not sheets:
        raise AppError("没有找到工作表", "请确认 Excel 里至少有一个 Sheet。")

    total_rows = 0
    written = 0
    total = len(sheets)
    for index, sheet in enumerate(sheets, start=1):
        if progress_cb:
            progress_cb(5 + int(index / total * 80), f"正在拆分工作表：{sheet}")
        try:
            df = read_excel(path, sheet_name=sheet)
        except AppError as exc:
            if "空" in exc.title:
                continue
            raise
        filename = _unique_filename(sheet, used_names)
        write_excel(df, output_dir / filename, sheet_name=sheet[:31])
        total_rows += len(df)
        written += 1

    if written == 0:
        raise AppError("没有可拆分的工作表数据", "请确认每个 Sheet 至少有一行表头和数据。")
    if progress_cb:
        progress_cb(100, "处理完成！")
    return SplitResult(output_dir=output_dir, file_count=written, row_count=total_rows, mode="sheet")


def _split_by_rows(
    path: Path,
    output_dir: Path,
    used_names: dict[str, int],
    rows_per_file: int,
    progress_cb,
) -> SplitResult:
    if rows_per_file < 1:
        raise AppError("每文件行数无效", "请输入大于 0 的整数，例如 1000。")
    if progress_cb:
        progress_cb(5, "正在读取 Excel")
    df = read_excel(path)
    if df.empty:
        raise AppError("没有可以拆分的数据", "请确认表格里至少有一行数据。")

    chunks = [df.iloc[i : i + rows_per_file] for i in range(0, len(df), rows_per_file)]
    total = len(chunks)
    for index, chunk in enumerate(chunks, start=1):
        filename = _unique_filename(f"第{index}批_{rows_per_file}行", used_names)
        write_excel(chunk.reset_index(drop=True), output_dir / filename)
        if progress_cb:
            progress_cb(10 + int(index / total * 85), f"正在生成 {filename}")

    if progress_cb:
        progress_cb(100, "处理完成！")
    return SplitResult(output_dir=output_dir, file_count=total, row_count=len(df), mode="rows")
