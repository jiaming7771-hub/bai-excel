from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app_config import ensure_output_dir
from utils.excel_utils import read_excel, write_excel
from utils.error_handler import AppError
from utils.file_utils import sanitize_filename


@dataclass
class SplitResult:
    output_dir: Path
    file_count: int
    row_count: int


def split_excel(
    path: Path,
    column: str,
    progress_cb=None,
) -> SplitResult:
    if progress_cb:
        progress_cb(5, "正在读取 Excel")
    df = read_excel(path)
    if column not in df.columns:
        raise AppError("找不到拆分字段", "请重新选择一个表格里实际存在的字段。")

    groups = list(df.groupby(column, dropna=False, sort=False))
    if not groups:
        raise AppError("没有可以拆分的数据", "请确认表格里至少有一行数据。")

    output_dir = ensure_output_dir() / f"split_{path.stem}"
    output_dir.mkdir(parents=True, exist_ok=True)

    used_names: dict[str, int] = {}
    total = len(groups)
    for index, (value, group) in enumerate(groups, start=1):
        base = sanitize_filename(value, fallback="空值")
        count = used_names.get(base, 0) + 1
        used_names[base] = count
        filename = f"{base}.xlsx" if count == 1 else f"{base}_{count}.xlsx"
        write_excel(group.reset_index(drop=True), output_dir / filename)
        if progress_cb:
            progress_cb(10 + int(index / total * 85), f"正在生成 {filename}")

    if progress_cb:
        progress_cb(100, "处理完成！")
    return SplitResult(output_dir=output_dir, file_count=len(groups), row_count=len(df))
