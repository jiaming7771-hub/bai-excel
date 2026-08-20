from pathlib import Path
import sys

APP_NAME = "Excel小工具箱"
APP_SUBTITLE = "批量处理 Excel，让重复工作一键完成"
APP_VERSION = "1.2.0"

EXCEL_SUFFIXES = {".xlsx", ".xls"}

ROOT_DIR = Path(__file__).resolve().parent
TESTDATA_DIR = ROOT_DIR / "testdata"


def default_output_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe = Path(sys.executable).resolve()
        if sys.platform == "darwin" and any(part.endswith(".app") for part in exe.parts):
            return Path.home() / "Desktop" / f"{APP_NAME}输出"
        return exe.parent / "output"
    return ROOT_DIR / "output"


OUTPUT_DIR = default_output_dir()


def ensure_output_dir() -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUT_DIR
