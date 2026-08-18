from __future__ import annotations

import pandas as pd

from core.excel_clean import clean_excel


def test_clean_removes_duplicates_and_blanks(testdata, tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    original = pd.read_excel(testdata["clean"])
    result = clean_excel(
        testdata["clean"],
        drop_duplicates=True,
        drop_blank_rows=True,
        dedupe_column="手机号",
    )
    cleaned = pd.read_excel(result.output_path)
    assert result.original_rows == len(original)
    assert result.cleaned_rows == len(cleaned)
    assert result.deleted_rows == result.original_rows - result.cleaned_rows
    assert result.deleted_rows > 0
    assert cleaned["手机号"].dropna().duplicated().sum() == 0
    assert not cleaned.isna().all(axis=1).any()
