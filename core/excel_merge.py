from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from app_config import ensure_output_dir
from utils.excel_utils import read_excel, write_excel
from utils.error_handler import AppError


@dataclass
class MergeResult:
    output_path: Path
    file_count: int
    row_count: int


def merge_excels(
    paths: list[Path],
    output_name: str = "merged_result.xlsx",
    progress_cb=None,
) -> MergeResult:
    if not paths:
        raise AppError("请先选择 Excel 文件", "可以一次选择多个文件，也可以选择一个包含 Excel 的文件夹。")

    frames: list[pd.DataFrame] = []
    columns: list[str] = []
    seen: set[str] = set()
    used_files = 0
    total = max(len(paths), 1)

    for index, path in enumerate(paths, start=1):
        if progress_cb:
            progress_cb(int(index / (total + 1) * 80), f"正在读取 {path.name}")
        try:
            df = read_excel(path)
        except AppError as exc:
            if "空" in exc.title:
                continue
            raise
        used_files += 1
        frames.append(df)
        for col in df.columns:
            if col not in seen:
                seen.add(col)
                columns.append(col)

    if not frames:
        raise AppError("这些文件都没有可合并的数据", "请确认每个 Excel 至少有一行表头和数据。")

    if progress_cb:
        progress_cb(88, "正在按列名对齐并合并")

    aligned = [frame.reindex(columns=columns) for frame in frames]
    merged = pd.concat(aligned, ignore_index=True, sort=False)
    output = ensure_output_dir() / output_name
    if progress_cb:
        progress_cb(95, "正在保存结果")
    write_excel(merged, output)
    if progress_cb:
        progress_cb(100, "处理完成！")
    return MergeResult(output_path=output, file_count=used_files, row_count=len(merged))
