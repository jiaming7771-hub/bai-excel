from __future__ import annotations

import random
import sys
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd

from app_config import TESTDATA_DIR


DEPARTMENTS = [
    "华东",
    "华南",
    "华北",
    "华中",
    "西南",
    "西北",
    "东北",
    "总部",
    "电商部",
    "客服部",
    "渠道部",
    "海外部",
]

SALES_PEOPLE = [f"销售{i:02d}" for i in range(1, 21)]
PRODUCTS = [f"产品{i:02d}" for i in range(1, 31)]
FIRST_NAMES = ["张", "李", "王", "赵", "刘", "陈", "杨", "黄", "周", "吴"]
GIVEN_NAMES = ["伟", "芳", "娜", "敏", "静", "丽", "强", "磊", "洋", "勇", "艳", "杰"]


def _name(index: int) -> str:
    return FIRST_NAMES[index % len(FIRST_NAMES)] + GIVEN_NAMES[index % len(GIVEN_NAMES)] + str(index % 97)


def _phone(index: int) -> str:
    return f"138{index:08d}"[-11:]


def generate_all(root: Path | None = None) -> dict[str, Path]:
    testdata = root or TESTDATA_DIR
    merge_dir = testdata / "test_merge"
    merge_dir.mkdir(parents=True, exist_ok=True)

    rows_a = []
    for i in range(400):
        rows_a.append({"姓名": _name(i), "手机号": _phone(i), "销售额": 1000 + i * 3})
    sales_a = merge_dir / "sales_a.xlsx"
    pd.DataFrame(rows_a).to_excel(sales_a, index=False)

    rows_b = []
    for i in range(400, 750):
        rows_b.append({"销售额": 2000 + i, "姓名": _name(i), "手机号": _phone(i)})
    sales_b = merge_dir / "sales_b.xlsx"
    pd.DataFrame(rows_b).to_excel(sales_b, index=False)

    rows_c = []
    for i in range(750, 1050):
        rows_c.append({"姓名": _name(i), "销售额": 1500 + i * 2})
    sales_c = merge_dir / "sales_c.xlsx"
    pd.DataFrame(rows_c).to_excel(sales_c, index=False)

    split_rows = []
    rng = random.Random(42)
    start = date(2025, 1, 1)
    for i in range(2200):
        split_rows.append(
            {
                "姓名": _name(i),
                "部门": DEPARTMENTS[i % len(DEPARTMENTS)],
                "销售额": rng.randint(800, 20000),
                "日期": start + timedelta(days=i % 360),
            }
        )
    split_path = testdata / "test_split.xlsx"
    pd.DataFrame(split_rows).to_excel(split_path, index=False)

    clean_rows = []
    for i in range(2600):
        clean_rows.append(
            {
                "姓名": _name(i),
                "手机号": _phone(i),
                "邮箱": f"user{i}@example.com",
                "订单号": f"ORD{i:05d}",
                "客户编号": f"C{i:04d}",
            }
        )
    # duplicate complete rows
    clean_rows.extend(clean_rows[:200])
    # duplicate phone numbers with different names
    for i in range(150):
        clean_rows.append(
            {
                "姓名": f"重复{i}",
                "手机号": _phone(i),
                "邮箱": f"dup{i}@example.com",
                "订单号": f"DUP{i:05d}",
                "客户编号": f"D{i:04d}",
            }
        )
    # blank rows and partial blanks
    for _ in range(80):
        clean_rows.append({"姓名": None, "手机号": None, "邮箱": None, "订单号": None, "客户编号": None})
    for i in range(50):
        clean_rows.append(
            {
                "姓名": _name(3000 + i),
                "手机号": None,
                "邮箱": "",
                "订单号": f"BLANK{i:03d}",
                "客户编号": None,
            }
        )
    clean_path = testdata / "test_clean.xlsx"
    pd.DataFrame(clean_rows).to_excel(clean_path, index=False)

    sales_rows = []
    for i in range(5200):
        month = (i % 12) + 1
        sales_rows.append(
            {
                "日期": date(2025, month, (i % 28) + 1),
                "销售人员": SALES_PEOPLE[i % len(SALES_PEOPLE)],
                "产品": PRODUCTS[i % len(PRODUCTS)],
                "数量": rng.randint(1, 12),
                "销售额": rng.randint(200, 8000),
            }
        )
    sales_path = testdata / "test_sales.xlsx"
    pd.DataFrame(sales_rows).to_excel(sales_path, index=False)

    return {
        "sales_a": sales_a,
        "sales_b": sales_b,
        "sales_c": sales_c,
        "split": split_path,
        "clean": clean_path,
        "sales": sales_path,
    }


if __name__ == "__main__":
    generated = generate_all()
    for name, path in generated.items():
        print(f"{name}: {path}")
