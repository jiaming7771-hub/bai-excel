"""统一输出：结果文件（可外传）与处理报告（内部核对）分目录存放。"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from app_config import ensure_output_dir
from utils.file_utils import sanitize_filename


def _stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M")


def _task_root(task: str, stem: str) -> Path:
    safe = sanitize_filename(stem, fallback="output")
    path = ensure_output_dir() / task / safe
    path.mkdir(parents=True, exist_ok=True)
    return path


def merge_paths(stem: str) -> tuple[Path, Path]:
    root = _task_root("合并", stem)
    stamp = _stamp()
    return root / f"合并结果_{stamp}.xlsx", root / f"合并报告_{stamp}.xlsx"


def split_paths(stem: str) -> tuple[Path, Path]:
    root = _task_root("拆分", stem)
    stamp = _stamp()
    result_dir = root / f"拆分结果_{stamp}"
    result_dir.mkdir(parents=True, exist_ok=True)
    return result_dir, root / f"拆分报告_{stamp}.xlsx"


def clean_paths(stem: str) -> tuple[Path, Path]:
    root = _task_root("清洗", stem)
    stamp = _stamp()
    return root / f"清洗结果_{stamp}.xlsx", root / f"清洗报告_{stamp}.xlsx"


def compare_paths(stem: str) -> tuple[Path, Path]:
    root = _task_root("对比", stem)
    stamp = _stamp()
    return root / f"对比结果_{stamp}.xlsx", root / f"对比报告_{stamp}.xlsx"
