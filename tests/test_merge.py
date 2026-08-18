from __future__ import annotations

import pandas as pd

from core.excel_merge import merge_excels


def test_merge_aligns_by_column_name(testdata, tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    result = merge_excels([testdata["sales_a"], testdata["sales_b"], testdata["sales_c"]])
    df = pd.read_excel(result.output_path)

    assert list(df.columns[:3]) == ["姓名", "手机号", "销售额"]
    assert result.file_count == 3
    assert result.row_count == 400 + 350 + 300
    assert len(df) == result.row_count
    # file C has no 手机号, those rows should stay empty instead of shifting
    missing_phone = df["手机号"].isna().sum()
    assert missing_phone == 300
    assert df["姓名"].notna().all()
    assert df["销售额"].notna().all()
