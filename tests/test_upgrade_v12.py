from __future__ import annotations

import pandas as pd
from openpyxl import Workbook

from core.excel_compare import compare_excels
from core.excel_merge import merge_excels
from core.excel_split import split_excel
from core.excel_clean import clean_excel


def test_merge_adds_source_and_aligns_aliases(tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"
    pd.DataFrame([{"姓名": "张三", "手机": "13800138001", "销售额": 100}]).to_excel(a, index=False)
    pd.DataFrame([{"姓名": "李四", "手机号": "13900139001", "销售额": 200}]).to_excel(b, index=False)

    result = merge_excels([a, b])
    df = pd.read_excel(result.output_path)
    assert "来源文件" in df.columns
    assert set(df["来源文件"].astype(str)) == {"a.xlsx", "b.xlsx"}
    assert "手机号" in df.columns
    assert df["手机号"].astype(str).tolist() == ["13800138001", "13900139001"]


def test_split_by_sheet_and_rows(tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    path = tmp_path / "multi.xlsx"
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "华东"
    ws1.append(["姓名", "销售额"])
    ws1.append(["张三", 100])
    ws2 = wb.create_sheet("华南")
    ws2.append(["姓名", "销售额"])
    ws2.append(["李四", 200])
    ws2.append(["王五", 300])
    wb.save(path)

    by_sheet = split_excel(path, mode="sheet")
    assert by_sheet.file_count == 2
    assert len(list(by_sheet.output_dir.glob("*.xlsx"))) == 2

    by_rows = split_excel(path, mode="rows", rows_per_file=1)
    # first sheet only is read by read_excel default - actually split by rows reads first sheet only
    # so 1 row from 华东 -> 1 file. Wait, read_excel reads sheet 0 which is 华东 with 1 data row.
    assert by_rows.file_count == 1
    assert by_rows.row_count == 1

    # dedicated single-sheet file for row split
    single = tmp_path / "single.xlsx"
    pd.DataFrame(
        [
            {"姓名": "A", "部门": "一"},
            {"姓名": "B", "部门": "二"},
            {"姓名": "C", "部门": "三"},
        ]
    ).to_excel(single, index=False)
    by_rows2 = split_excel(single, mode="rows", rows_per_file=2)
    assert by_rows2.file_count == 2
    assert by_rows2.row_count == 3


def test_clean_keep_latest_and_combo_key(tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    source = tmp_path / "dup.xlsx"
    pd.DataFrame(
        [
            {"姓名": "张三", "手机号": "13800138001", "更新时间": "2024-01-01", "备注": "旧"},
            {"姓名": "张三", "手机号": "13800138001", "更新时间": "2024-06-01", "备注": "新"},
            {"姓名": "李四", "手机号": "13900139001", "更新时间": "2024-02-01", "备注": "A"},
            {"姓名": "李四", "手机号": "13900139002", "更新时间": "2024-03-01", "备注": "B"},
        ]
    ).to_excel(source, index=False)

    result = clean_excel(
        source,
        drop_duplicates=False,
        drop_blank_rows=False,
        dedupe_columns=["姓名", "手机号"],
        dedupe_keep="latest",
        dedupe_latest_column="更新时间",
    )
    cleaned = pd.read_excel(result.output_path, sheet_name="清洗结果")
    # 张三同手机只留最新；李四手机不同都留
    assert len(cleaned) == 3
    zhang = cleaned[cleaned["姓名"] == "张三"].iloc[0]
    assert "新" in str(zhang["备注"])


def test_compare_two_tables(tmp_path, monkeypatch):
    import app_config

    monkeypatch.setattr(app_config, "OUTPUT_DIR", tmp_path)
    a = tmp_path / "old.xlsx"
    b = tmp_path / "new.xlsx"
    pd.DataFrame(
        [
            {"手机号": "13800138001", "姓名": "张三", "城市": "上海"},
            {"手机号": "13800138002", "姓名": "李四", "城市": "北京"},
            {"手机号": "13800138003", "姓名": "王五", "城市": "广州"},
        ]
    ).to_excel(a, index=False)
    pd.DataFrame(
        [
            {"手机号": "13800138001", "姓名": "张三", "城市": "杭州"},  # changed
            {"手机号": "13800138002", "姓名": "李四", "城市": "北京"},  # same
            {"手机号": "13800138004", "姓名": "赵六", "城市": "深圳"},  # only B
        ]
    ).to_excel(b, index=False)

    result = compare_excels(a, b, ["手机号"])
    assert result.only_a == 1  # 王五
    assert result.only_b == 1  # 赵六
    assert result.changed == 1  # 张三城市变了
    assert result.same == 1

    only_a = pd.read_excel(result.output_path, sheet_name="仅在表A")
    only_b = pd.read_excel(result.output_path, sheet_name="仅在表B")
    changed = pd.read_excel(result.output_path, sheet_name="有变化")
    assert len(only_a) == 1
    assert len(only_b) == 1
    assert len(changed) == 1
    assert "城市" in str(changed.iloc[0]["变化字段"])
