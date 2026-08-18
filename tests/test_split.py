from __future__ import annotations

import pandas as pd

from core.excel_split import split_excel


def test_split_by_department(testdata, tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    result = split_excel(testdata["split"], "部门")
    files = sorted(result.output_dir.glob("*.xlsx"))
    assert result.file_count >= 10
    assert len(files) == result.file_count

    original = pd.read_excel(testdata["split"])
    total = 0
    for path in files:
        part = pd.read_excel(path)
        assert list(part.columns) == list(original.columns)
        assert len(part) > 0
        total += len(part)
    assert total == len(original)
