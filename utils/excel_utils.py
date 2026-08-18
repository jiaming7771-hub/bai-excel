from __future__ import annotations

from pathlib import Path

import pandas as pd

from utils.error_handler import AppError, friendly_error


def read_excel(path: str | Path) -> pd.DataFrame:
    file_path = Path(path)
    suffix = file_path.suffix.lower()
    try:
        if suffix == ".xls":
            df = pd.read_excel(file_path, engine="xlrd")
        else:
            df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as exc:
        raise friendly_error(exc) from exc

    if df is None or df.empty:
        raise AppError("Excel 文件是空的", "请确认表格里至少有一行数据。")
    df.columns = [str(c).strip() for c in df.columns]
    if any(col.startswith("Unnamed") for col in df.columns) and len(df.columns) == 1:
        raise AppError("没有识别到表头", "请确认第一行是表头，并且表格内容完整。")
    return df


def write_excel(df: pd.DataFrame, path: str | Path, sheet_name: str = "Sheet1") -> Path:
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_excel(output, index=False, sheet_name=sheet_name, engine="openpyxl")
    except PermissionError as exc:
        raise AppError(
            "无法保存结果文件",
            "请确认输出文件没有被 Excel 打开，并且您有权限写入这个文件夹。",
        ) from exc
    except Exception as exc:
        raise friendly_error(exc) from exc
    return output


def list_columns(df: pd.DataFrame) -> list[str]:
    return [str(c) for c in df.columns]
