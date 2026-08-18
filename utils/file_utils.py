from __future__ import annotations

import os
import re
import subprocess
import sys
from pathlib import Path

from app_config import EXCEL_SUFFIXES
from utils.error_handler import AppError


def is_excel_file(path: str | Path) -> bool:
    return Path(path).suffix.lower() in EXCEL_SUFFIXES


def collect_excel_files(paths: list[str | Path]) -> list[Path]:
    files: list[Path] = []
    seen: set[Path] = set()
    for raw in paths:
        path = Path(raw).expanduser().resolve()
        if not path.exists():
            continue
        candidates: list[Path]
        if path.is_dir():
            candidates = sorted(
                [p for p in path.iterdir() if p.is_file() and is_excel_file(p)],
                key=lambda p: p.name.lower(),
            )
        elif is_excel_file(path):
            candidates = [path]
        else:
            continue
        for item in candidates:
            if item not in seen:
                seen.add(item)
                files.append(item)
    if not files:
        raise AppError("没有找到 Excel 文件", "请选择 .xlsx 或 .xls 文件，也可以选择包含这些文件的文件夹。")
    return files


def sanitize_filename(name: str, fallback: str = "未命名") -> str:
    text = str(name).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        text = fallback
    text = re.sub(r'[\\/:*?"<>|]', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return (text[:80] or fallback)


def open_folder(path: str | Path) -> None:
    folder = Path(path).expanduser().resolve()
    if folder.is_file():
        folder = folder.parent
    folder.mkdir(parents=True, exist_ok=True)
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(folder)])
    elif os.name == "nt":
        os.startfile(str(folder))  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(folder)])


def open_file(path: str | Path) -> None:
    """用系统默认程序打开文件（如 WPS / Excel）。"""
    target = Path(path).expanduser().resolve()
    if not target.exists():
        raise AppError("找不到结果文件", "请确认文件还在，或重新运行一次处理。")
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(target)])
    elif os.name == "nt":
        os.startfile(str(target))  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(target)])
