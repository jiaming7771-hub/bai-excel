from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from app_config import ensure_output_dir
from utils.excel_utils import normalize_header_name, read_excel, write_excel
from utils.error_handler import AppError

# 同义词映射：把常见别名对齐到统一列名，减少「手机 vs 手机号」对不齐
HEADER_ALIASES = {
    "手机": "手机号",
    "电话": "手机号",
    "联系电话": "手机号",
    "mobile": "手机号",
    "phone": "手机号",
    "tel": "手机号",
    "客户名": "姓名",
    "联系人": "姓名",
    "名字": "姓名",
    "邮件": "邮箱",
    "email": "邮箱",
    "mail": "邮箱",
    "e-mail": "邮箱",
}


@dataclass
class MergeResult:
    output_path: Path
    file_count: int
    row_count: int


def _align_headers(df: pd.DataFrame) -> pd.DataFrame:
    renamed: dict[str, str] = {}
    used: set[str] = set()
    for col in df.columns:
        raw = normalize_header_name(col)
        mapped = HEADER_ALIASES.get(raw.lower(), HEADER_ALIASES.get(raw, raw))
        name = mapped
        # 避免重名列互相覆盖
        if name in used:
            index = 2
            while f"{name}_{index}" in used:
                index += 1
            name = f"{name}_{index}"
        used.add(name)
        renamed[col] = name
    out = df.rename(columns=renamed)
    out.columns = [str(c).strip() for c in out.columns]
    return out


def merge_excels(
    paths: list[Path],
    output_name: str = "merged_result.xlsx",
    add_source_column: bool = True,
    progress_cb=None,
) -> MergeResult:
    if not paths:
        raise AppError("请先选择 Excel 文件", "可以一次选择多个文件，也可以选择一个包含 Excel 的文件夹。")

    frames: list[pd.DataFrame] = []
    columns: list[str] = []
    seen: set[str] = set()
    used_files = 0
    total = max(len(paths), 1)
    source_col = "来源文件"

    for index, path in enumerate(paths, start=1):
        if progress_cb:
            progress_cb(int(index / (total + 1) * 80), f"正在读取 {path.name}")
        try:
            df = _align_headers(read_excel(path))
        except AppError as exc:
            if "空" in exc.title:
                continue
            raise
        used_files += 1
        if add_source_column:
            # 若原表已有同名列，先让位
            if source_col in df.columns:
                df = df.rename(columns={source_col: f"{source_col}_原表"})
            df.insert(0, source_col, path.name)
        frames.append(df)
        for col in df.columns:
            if col not in seen:
                seen.add(col)
                columns.append(col)

    if not frames:
        raise AppError("这些文件都没有可合并的数据", "请确认每个 Excel 至少有一行表头和数据。")

    if progress_cb:
        progress_cb(88, "正在按列名对齐并合并")

    # 来源文件列固定在最前
    if add_source_column and source_col in columns:
        columns = [source_col] + [c for c in columns if c != source_col]

    aligned = [frame.reindex(columns=columns) for frame in frames]
    merged = pd.concat(aligned, ignore_index=True, sort=False)
    output = ensure_output_dir() / output_name
    if progress_cb:
        progress_cb(95, "正在保存结果")
    write_excel(merged, output)
    if progress_cb:
        progress_cb(100, "处理完成！")
    return MergeResult(output_path=output, file_count=used_files, row_count=len(merged))
