from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from core.excel_clean import clean_excel
from core.excel_merge import merge_excels
from core.excel_split import split_excel
from utils.error_handler import AppError, friendly_error
from utils.file_utils import collect_excel_files


def test_merge_requires_files(tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    with pytest.raises(AppError) as exc:
        merge_excels([])
    assert "请先选择" in exc.value.title
    assert "DataFrame" not in str(exc.value)
    assert "traceback" not in str(exc.value).lower()


def test_split_missing_column(testdata, tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    with pytest.raises(AppError) as exc:
        split_excel(testdata["split"], "不存在的字段")
    assert "找不到拆分字段" in exc.value.title


def test_clean_requires_option(testdata, tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    with pytest.raises(AppError) as exc:
        clean_excel(
            testdata["clean"],
            drop_duplicates=False,
            drop_blank_rows=False,
            check_shifted_rows=False,
            fix_phone=False,
            fix_email=False,
            fix_dates=False,
            check_required_fields=False,
            trim_spaces=False,
        )
    assert "清洗方式" in exc.value.title


def test_friendly_error_stays_chinese():
    err = friendly_error(PermissionError("Permission denied: being used by another process"))
    assert "Excel 文件读取失败" == err.title
    assert "占用" in err.hint
    assert "PermissionError" not in err.title
    assert "traceback" not in err.hint.lower()


def test_collect_excel_from_folder(testdata):
    folder = testdata["sales_a"].parent
    files = collect_excel_files([folder])
    names = {path.name for path in files}
    assert {"sales_a.xlsx", "sales_b.xlsx", "sales_c.xlsx"} <= names


def test_merge_keeps_chinese_headers(testdata, tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    result = merge_excels([testdata["sales_b"], testdata["sales_a"]])
    df = pd.read_excel(result.output_path)
    # file B starts with 销售额, but merge should still keep names, not positions
    assert "姓名" in df.columns
    assert "手机号" in df.columns
    assert "销售额" in df.columns
