from pathlib import Path

APP_NAME = "Excel小工具箱"
APP_SUBTITLE = "批量处理 Excel，让重复工作一键完成"
APP_VERSION = "1.0.0"

EXCEL_SUFFIXES = {".xlsx", ".xls"}

ROOT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT_DIR / "output"
TESTDATA_DIR = ROOT_DIR / "testdata"


def ensure_output_dir() -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUT_DIR
