from __future__ import annotations

import random
import shutil
import sys
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd

from app_config import TESTDATA_DIR
from utils.excel_utils import write_excel

CASES_DIR = TESTDATA_DIR / "cases"
DESKTOP_DIR = Path.home() / "Desktop" / "Excel小工具箱测试案例"

FIRST_NAMES = ["张", "李", "王", "赵", "刘", "陈", "杨", "黄", "周", "吴", "徐", "孙", "马", "朱", "胡", "郭", "何", "高", "林", "罗"]
GIVEN_NAMES = ["伟", "芳", "娜", "敏", "静", "丽", "强", "磊", "洋", "勇", "艳", "杰", "涛", "超", "霞", "平", "刚", "桂英", "秀英", "婷"]
CITIES = ["上海", "杭州", "南京", "苏州", "合肥", "宁波", "无锡", "温州", "金华", "嘉兴"]
CHANNELS = ["抖音", "小红书", "微信私域", "线下门店", "天猫", "拼多多", "转介绍", "展会"]
LEVELS = ["普通", "银卡", "金卡", "钻石"]
MANAGERS = ["张明", "李雪", "王鹏", "赵倩", "陈晨", "周杰"]
STATUSES = ["跟进中", "已成交", "已流失", "待联系"]
EXT_CITIES = CITIES + ["成都", "重庆", "武汉", "青岛", "厦门", "西安", "东莞"]
PRODUCTS = [
    ("SKU001", "护手霜", 49),
    ("SKU002", "面膜", 79),
    ("SKU003", "精华液", 129),
    ("SKU004", "口红", 89),
    ("SKU005", "眼影盘", 99),
    ("SKU006", "洁面乳", 59),
    ("SKU007", "防晒喷雾", 69),
    ("SKU008", "卸妆水", 39),
    ("SKU009", "身体乳", 109),
    ("SKU010", "洗发水", 119),
]
WAREHOUSES = ["杭州仓", "广州仓", "上海仓"]
SHIP_STATUSES = ["待发货", "已发货", "已签收", "已取消"]
DEPARTMENTS = ["华东", "华南", "华北", "西南", "线上", "客服中心"]
JOBS = ["销售", "主管", "客服", "运营"]
CATEGORIES = ["护肤", "彩妆", "个护", "香氛", "母婴"]
PAY_METHODS = ["微信", "支付宝", "银行卡", "现金", "对公转账"]

# ── 生产级数据规模（贴近真实业务：多文件合计上万行、单表数千行）──
PROD_MERGE_ROWS_PER_FILE = 1500
PROD_STAFF_ROWS = 5000
PROD_SHEET_ROWS = 800
PROD_BIG_ORDER_ROWS = 12000
PROD_BIG_ORDER_SPLIT = 2000
PROD_CLEAN_BASE = 5000
PROD_COMPARE_SHARED = 800
PROD_COMPARE_ONLY_OLD = 80
PROD_COMPARE_ONLY_NEW = 60
PROD_COMPARE_CHANGED = 120
PROD_ORDER_SHARED = 2000
PROD_ORDER_CHANGED = 250
PROD_ORDER_ONLY = 120

MERGE_STORES = [
    ("华东门店.xlsx", "east", PROD_MERGE_ROWS_PER_FILE, 1),
    ("华南门店.xlsx", "south", PROD_MERGE_ROWS_PER_FILE, 50000),
    ("华北门店.xlsx", "north", PROD_MERGE_ROWS_PER_FILE, 100000),
    ("西南门店.xlsx", "southwest", PROD_MERGE_ROWS_PER_FILE, 150000),
    ("华中门店.xlsx", "central", PROD_MERGE_ROWS_PER_FILE, 200000),
    ("线上渠道.xlsx", "online", PROD_MERGE_ROWS_PER_FILE, 250000),
    ("东北门店.xlsx", "east", PROD_MERGE_ROWS_PER_FILE, 300000),
    ("西北门店.xlsx", "north", PROD_MERGE_ROWS_PER_FILE, 350000),
]
SPLIT_SHEETS = ["华东", "华南", "华北", "西南", "线上", "汇总备查"]


def _save(df: pd.DataFrame, path: Path) -> Path:
    return write_excel(df, path)


def _person_name(index: int) -> str:
    return FIRST_NAMES[index % len(FIRST_NAMES)] + GIVEN_NAMES[index % len(GIVEN_NAMES)]


def _phone(index: int) -> str:
    return f"1{38 + (index % 50):02d}{index % 100000000:08d}"[-11:]


def _build_clean_rows(total: int = 1200) -> list[dict]:
    """生成更脏的客户表：重复、空白、窜行、错误手机号/邮箱、乱日期。源表不预写清洗标记。"""
    rng = random.Random(20260818)
    start = date(2024, 1, 1)
    rows: list[dict] = []

    for i in range(total):
        name = _person_name(i)
        phone = _phone(i)
        rows.append(
            {
                "姓名": name,
                "手机号": phone,
                "邮箱": f"user{i}@example.com",
                "订单号": f"ORD{i:05d}",
                "客户编号": f"C{i:04d}",
                "城市": CITIES[i % len(CITIES)],
                "来源渠道": CHANNELS[i % len(CHANNELS)],
                "会员等级": LEVELS[i % len(LEVELS)],
                "最近下单": start + timedelta(days=i % 500),
                "累计消费": rng.randint(50, 28000),
                "备注": "" if i % 7 else "老客户",
            }
        )

    # 整行完全重复
    rows.extend(rows[20:70])
    rows.extend(rows[120:150])

    # 同一手机号、不同订单
    for i in range(80):
        base = rows[i]
        rows.append(
            {
                "姓名": _person_name(5000 + i),
                "手机号": base["手机号"],
                "邮箱": f"client{i}@mail.com",
                "订单号": f"ORD9{i:04d}",
                "客户编号": f"C9{i:03d}",
                "城市": base["城市"],
                "来源渠道": CHANNELS[(i + 3) % len(CHANNELS)],
                "会员等级": base["会员等级"],
                "最近下单": base["最近下单"],
                "累计消费": base["累计消费"] + rng.randint(1, 300),
                "备注": "",
            }
        )

    # 完全空白行
    blank = {key: None for key in rows[0].keys()}
    for _ in range(45):
        rows.append(dict(blank))

    # 残缺字段
    for i in range(60):
        rows.append(
            {
                "姓名": _person_name(3000 + i),
                "手机号": None,
                "邮箱": "" if i % 2 == 0 else None,
                "订单号": f"ORD8{i:03d}",
                "客户编号": None,
                "城市": CITIES[i % len(CITIES)] if i % 3 else None,
                "来源渠道": CHANNELS[i % len(CHANNELS)],
                "会员等级": None,
                "最近下单": None,
                "累计消费": None,
                "备注": "",
            }
        )

    # 姓名空格 / 邮箱大小写
    for i in range(25):
        src = rows[200 + i]
        rows.append(
            {
                **src,
                "姓名": f" {src['姓名']} ",
                "邮箱": str(src["邮箱"]).upper(),
                "订单号": f"ORD7{i:03d}",
                "客户编号": f"C7{i:03d}",
                "备注": "",
            }
        )

    # 窜行：字段错位
    for i in range(40):
        base = rows[300 + i]
        rows.append(
            {
                "姓名": base["手机号"],  # 姓名里塞了手机号
                "手机号": base["邮箱"],  # 手机号里塞了邮箱
                "邮箱": base["姓名"],  # 邮箱里塞了姓名
                "订单号": f"ORD6{i:03d}",
                "客户编号": f"C6{i:03d}",
                "城市": base["城市"],
                "来源渠道": base["来源渠道"],
                "会员等级": base["会员等级"],
                "最近下单": base["最近下单"],
                "累计消费": base["累计消费"],
                "备注": "",
            }
        )

    # 手机号不正确
    bad_phones = [
        "12345",
        "23800138000",
        "1380013800a",
        "138-0013-8000",
        "01088886666",
        "8613800138000",
        "手机号暂无",
        "00000000000",
        "199999",
        "abc13800138000",
    ]
    for i, bad in enumerate(bad_phones * 4):  # 40 条
        rows.append(
            {
                "姓名": _person_name(6000 + i),
                "手机号": bad,
                "邮箱": f"badphone{i}@example.com",
                "订单号": f"ORD5{i:03d}",
                "客户编号": f"C5{i:03d}",
                "城市": CITIES[i % len(CITIES)],
                "来源渠道": CHANNELS[i % len(CHANNELS)],
                "会员等级": LEVELS[i % len(LEVELS)],
                "最近下单": start + timedelta(days=i),
                "累计消费": rng.randint(80, 9000),
                "备注": "",
            }
        )

    # 邮箱不正确
    bad_emails = [
        "not-an-email",
        "abc@",
        "@qq.com",
        "a@@b.com",
        "test#mail.com",
        "用户@公司",
        "hello world@test.com",
        "abc.def",
        "123456",
        "null",
    ]
    for i, bad in enumerate(bad_emails * 4):  # 40 条
        rows.append(
            {
                "姓名": _person_name(7000 + i),
                "手机号": _phone(7000 + i),
                "邮箱": bad,
                "订单号": f"ORD4{i:03d}",
                "客户编号": f"C4{i:03d}",
                "城市": CITIES[i % len(CITIES)],
                "来源渠道": CHANNELS[i % len(CHANNELS)],
                "会员等级": LEVELS[i % len(LEVELS)],
                "最近下单": start + timedelta(days=i + 10),
                "累计消费": rng.randint(80, 9000),
                "备注": "",
            }
        )

    # 日期格式混乱：可修正 + 不可修正
    messy_dates = [
        "2024/1/5",
        "2024.03.08",
        "20240315",
        "2024年4月9日",
        "03/20/2024",
        "2024-5-1 0:00:00",
        "2024/13/40",  # 无效
        "昨天",
        "null",
        "2024-99-99",
        "15-08-2024",
        "2024 6 18",
    ]
    for i, messy in enumerate(messy_dates * 5):  # 60 条
        rows.append(
            {
                "姓名": _person_name(8000 + i),
                "手机号": _phone(8000 + i),
                "邮箱": f"date{i}@example.com",
                "订单号": f"ORD3{i:03d}",
                "客户编号": f"C3{i:03d}",
                "城市": CITIES[i % len(CITIES)],
                "来源渠道": CHANNELS[i % len(CHANNELS)],
                "会员等级": LEVELS[i % len(LEVELS)],
                "最近下单": messy,
                "累计消费": rng.randint(80, 9000),
                "备注": "",
            }
        )

    rng.shuffle(rows)
    return rows


def _save_sheets(sheets: dict[str, pd.DataFrame], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    return path


def _merge_stats(paths: list[Path]):
    import tempfile

    import app_config
    from core.excel_merge import merge_excels

    tmp = Path(tempfile.mkdtemp())
    prev = app_config.OUTPUT_DIR
    try:
        app_config.OUTPUT_DIR = tmp
        result = merge_excels(paths)
        df = pd.read_excel(result.output_path)
        return result, df
    finally:
        app_config.OUTPUT_DIR = prev


def _split_stats(path: Path, mode: str, column: str | None = None, rows_per_file: int = 200):
    import tempfile

    import app_config
    from core.excel_split import split_excel

    tmp = Path(tempfile.mkdtemp())
    prev = app_config.OUTPUT_DIR
    try:
        app_config.OUTPUT_DIR = tmp
        return split_excel(path, column=column, mode=mode, rows_per_file=rows_per_file)
    finally:
        app_config.OUTPUT_DIR = prev


def _build_merge_store_df(style: str, count: int, base_i: int, rng: random.Random) -> pd.DataFrame:
    rows: list[dict] = []
    start = date(2024, 1, 1)
    for j in range(count):
        i = base_i + j
        common = {
            "订单日期": start + timedelta(days=j % 365),
            "商品类目": CATEGORIES[j % len(CATEGORIES)],
            "支付方式": PAY_METHODS[j % len(PAY_METHODS)],
            "门店编码": f"S{(i % 900) + 100:03d}",
        }
        if style == "east":
            rows.append(
                {
                    **common,
                    "姓名": _person_name(i),
                    "手机号": _phone(i),
                    "销售额": rng.randint(800, 58000),
                    "城市": EXT_CITIES[i % len(EXT_CITIES)],
                }
            )
        elif style == "south":
            rows.append(
                {
                    **common,
                    "销售额": rng.randint(800, 58000),
                    "姓名": _person_name(i),
                    "手机": _phone(i),
                    "城市": EXT_CITIES[i % len(EXT_CITIES)],
                }
            )
        elif style == "north":
            rows.append(
                {
                    **common,
                    "姓名": _person_name(i),
                    "销售额": rng.randint(800, 58000),
                    "联系电话": _phone(i),
                }
            )
        elif style == "southwest":
            rows.append(
                {
                    **common,
                    "客户名": _person_name(i),
                    "手机号": _phone(i),
                    "销售金额": rng.randint(800, 58000),
                }
            )
        elif style == "central":
            rows.append(
                {
                    **common,
                    "姓名": _person_name(i),
                    "销售额": rng.randint(800, 58000),
                }
            )
        else:
            rows.append(
                {
                    **common,
                    "姓名": _person_name(i),
                    "手机号": _phone(i),
                    "销售额": rng.randint(800, 58000),
                    "来源渠道": CHANNELS[i % len(CHANNELS)],
                }
            )
    return pd.DataFrame(rows)


def _build_staff_rows(total: int = 168) -> list[dict]:
    rng = random.Random(20260828)
    rows: list[dict] = []
    for i in range(total):
        rows.append(
            {
                "姓名": _person_name(i),
                "部门": DEPARTMENTS[i % len(DEPARTMENTS)],
                "岗位": JOBS[i % len(JOBS)],
                "入职日期": date(2020, 1, 1) + timedelta(days=i * 13),
                "销售额": rng.randint(0, 32000),
                "城市": EXT_CITIES[i % len(EXT_CITIES)],
            }
        )
    return rows


def _build_region_sheet_rows(region: str, count: int, base_i: int, rng: random.Random) -> list[dict]:
    rows: list[dict] = []
    for j in range(count):
        i = base_i + j
        rows.append(
            {
                "姓名": _person_name(i),
                "手机号": _phone(i),
                "销售额": rng.randint(2500, 24000),
                "城市": EXT_CITIES[i % len(EXT_CITIES)],
                "渠道": CHANNELS[i % len(CHANNELS)],
            }
        )
    return rows


def _build_big_order_rows(count: int = PROD_BIG_ORDER_ROWS) -> list[dict]:
    rng = random.Random(20260829)
    rows: list[dict] = []
    for i in range(count):
        sku, pname, price = PRODUCTS[i % len(PRODUCTS)]
        qty = rng.randint(1, 8)
        rows.append(
            {
                "订单号": f"ORD{2025}{i:05d}",
                "日期": date(2025, 1, 1) + timedelta(days=i % 120),
                "客户姓名": _person_name(i),
                "手机号": _phone(i),
                "商品": pname,
                "商品编码": sku,
                "数量": qty,
                "单价": price,
                "金额": qty * price,
                "仓库": WAREHOUSES[i % len(WAREHOUSES)],
            }
        )
    return rows


def _compare_stats(old_path: Path, new_path: Path, key_columns: list[str]):
    """生成时自测对比结果，写入怎么测.txt 用。"""
    import tempfile

    import app_config
    from core.excel_compare import compare_excels

    tmp = Path(tempfile.mkdtemp())
    prev = app_config.OUTPUT_DIR
    try:
        app_config.OUTPUT_DIR = tmp
        return compare_excels(old_path, new_path, key_columns, output_name="_verify.xlsx")
    finally:
        app_config.OUTPUT_DIR = prev


def _customer_pool_row(index: int, rng: random.Random, note: str = "") -> dict:
    start = date(2025, 1, 1)
    return {
        "客户编号": f"C{index:04d}",
        "手机号": _phone(index),
        "姓名": _person_name(index),
        "城市": EXT_CITIES[index % len(EXT_CITIES)],
        "渠道": CHANNELS[index % len(CHANNELS)],
        "会员等级": LEVELS[index % len(LEVELS)],
        "负责人": MANAGERS[index % len(MANAGERS)],
        "状态": STATUSES[index % len(STATUSES)],
        "累计消费": rng.randint(300, 28000),
        "最近跟进": (start + timedelta(days=index % 90)).isoformat(),
        "备注": note,
    }


def _build_compare_customer_pool() -> tuple[pd.DataFrame, pd.DataFrame]:
    """生产级客户池：约千行级，含删增改、重复主键、空主键。"""
    rng = random.Random(20260825)
    n_shared = PROD_COMPARE_SHARED
    shared = [_customer_pool_row(i, rng) for i in range(1, n_shared + 1)]
    only_old = [
        _customer_pool_row(i, rng, "旧表独有-已删除客户")
        for i in range(n_shared + 1, n_shared + PROD_COMPARE_ONLY_OLD + 1)
    ]

    dup = dict(shared[127])
    dup["客户编号"] = "C_dup1"
    dup["备注"] = "同手机号重复行-旧表多一行"

    empty_key = _customer_pool_row(900001, rng, "手机号为空")
    empty_key["客户编号"] = "C_empty"
    empty_key["手机号"] = ""

    old_rows = shared + only_old + [dup, empty_key]
    rng.shuffle(old_rows)

    change_idx = set(rng.sample(range(n_shared), PROD_COMPARE_CHANGED))
    new_shared: list[dict] = []
    for i, row in enumerate(shared):
        new_row = dict(row)
        if i in change_idx:
            new_row["城市"] = EXT_CITIES[(i + 5) % len(EXT_CITIES)]
            if i % 3 == 0:
                new_row["会员等级"] = LEVELS[(i + 2) % len(LEVELS)]
            elif i % 3 == 1:
                new_row["状态"] = STATUSES[(i + 1) % len(STATUSES)]
            else:
                new_row["累计消费"] = int(new_row["累计消费"]) + rng.randint(120, 800)
                new_row["负责人"] = MANAGERS[(i + 3) % len(MANAGERS)]
        new_shared.append(new_row)

    only_new_start = n_shared + PROD_COMPARE_ONLY_OLD + 1
    only_new = [
        _customer_pool_row(i, rng, "新版新增客户")
        for i in range(only_new_start, only_new_start + PROD_COMPARE_ONLY_NEW)
    ]
    new_rows = new_shared + only_new
    rng.shuffle(new_rows)
    return pd.DataFrame(old_rows), pd.DataFrame(new_rows)


def _build_compare_orders() -> tuple[pd.DataFrame, pd.DataFrame]:
    """生产级订单：约两千行，组合主键，一单多商品。"""
    rng = random.Random(20260826)

    def line(order_no: int, sku_idx: int, person_idx: int) -> dict:
        sku, _, price = PRODUCTS[sku_idx % len(PRODUCTS)]
        qty = rng.randint(1, 12)
        return {
            "订单号": f"SO2025{order_no:06d}",
            "商品编码": sku,
            "客户姓名": _person_name(person_idx),
            "手机号": _phone(person_idx),
            "数量": qty,
            "单价": price,
            "金额": qty * price,
            "仓库": WAREHOUSES[order_no % len(WAREHOUSES)],
            "发货状态": SHIP_STATUSES[order_no % len(SHIP_STATUSES)],
            "下单日期": date(2025, 1, 1) + timedelta(days=order_no % 180),
        }

    shared: list[dict] = []
    # 单笔订单
    for i in range(PROD_ORDER_SHARED - 240):
        shared.append(line(100000 + i, i, i + 10))
    # 一单多商品（每单 2 行）
    for i in range(120):
        order_no = 200000 + i
        shared.append(line(order_no, i, i + 5000))
        shared.append(line(order_no, (i + 5) % len(PRODUCTS), i + 5000))

    changed_old: list[dict] = []
    changed_new: list[dict] = []
    for j in range(PROD_ORDER_CHANGED):
        order_no = 300000 + j
        old_row = line(order_no, j + 2, j + 9000)
        new_row = dict(old_row)
        if j % 3 == 0:
            new_row["数量"] = int(old_row["数量"]) + 2
            new_row["金额"] = new_row["数量"] * new_row["单价"]
        elif j % 3 == 1:
            new_row["发货状态"] = SHIP_STATUSES[(j + 2) % len(SHIP_STATUSES)]
        else:
            new_row["仓库"] = WAREHOUSES[(j + 1) % len(WAREHOUSES)]
        changed_old.append(old_row)
        changed_new.append(new_row)

    only_old = [line(400000 + k, k + 5, k + 20000) for k in range(PROD_ORDER_ONLY)]
    only_new = [line(500000 + k, k + 7, k + 30000) for k in range(PROD_ORDER_ONLY)]

    old_rows = shared + changed_old + only_old
    new_rows = shared + changed_new + only_new
    rng.shuffle(old_rows)
    rng.shuffle(new_rows)
    return pd.DataFrame(old_rows), pd.DataFrame(new_rows)


def _write_compare_guide(path: Path, title: str, sections: list[str]) -> None:
    path.write_text(title + "\n\n" + "\n\n".join(sections) + "\n", encoding="utf-8")


def generate_cases(root: Path | None = None) -> dict[str, Path]:
    cases = root or CASES_DIR
    if cases.exists():
        shutil.rmtree(cases)
    cases.mkdir(parents=True, exist_ok=True)

    merge_dir = cases / "01-批量合并"
    split_dir = cases / "02-数据拆分"
    clean_dir = cases / "03-客户名单清洗"
    sales_dir = cases / "04-销售数据汇总"
    compare_dir = cases / "05-两表对比"

    merge_rng = random.Random(20260827)
    merge_paths: list[Path] = []
    merge_meta: list[str] = []
    merge_files: dict[str, Path] = {}
    merge_row_sum = 0
    for filename, style, count, base_i in MERGE_STORES:
        df = _build_merge_store_df(style, count, base_i, merge_rng)
        path = _save(df, merge_dir / filename)
        merge_paths.append(path)
        merge_files[style] = path
        merge_row_sum += len(df)
        merge_meta.append(f"  · {filename}（{len(df)} 行，列：{', '.join(df.columns)}）")

    merge_result, merged_df = _merge_stats(merge_paths)
    source_values = set(merged_df["来源文件"].astype(str))
    (merge_dir / "怎么测.txt").write_text(
        "【01-批量合并 · 生产级测试数据】\n\n"
        f"8 个区域/渠道表，每个约 {PROD_MERGE_ROWS_PER_FILE} 行，合计约 {merge_row_sum} 行。\n"
        "列名故意不统一（手机/联系电话/客户名/销售金额 等），并含订单日期、类目、支付方式。\n"
        + "\n".join(merge_meta)
        + "\n\n操作：选中整个文件夹 → 开始合并\n"
        f"自测：合并后 {len(merged_df)} 行，含「来源文件」「手机号」；"
        f"{len(source_values)} 个来源文件均可追溯。\n"
        "适合验证：万级合并速度、列名对齐、来源标注。",
        encoding="utf-8",
    )

    staff_df = pd.DataFrame(_build_staff_rows(PROD_STAFF_ROWS))
    staff = _save(staff_df, split_dir / "按部门-员工花名册.xlsx")
    staff_split = _split_stats(staff, "column", column="部门")

    sheet_rng = random.Random(20260830)
    sheet_map = {
        name: pd.DataFrame(
            _build_region_sheet_rows(name, PROD_SHEET_ROWS, 600 + idx * PROD_SHEET_ROWS, sheet_rng)
        )
        for idx, name in enumerate(SPLIT_SHEETS)
    }
    multi_sheet = _save_sheets(sheet_map, split_dir / "按工作表-区域销售.xlsx")
    sheet_split = _split_stats(multi_sheet, "sheet")

    big_orders = pd.DataFrame(_build_big_order_rows())
    big_table = _save(big_orders, split_dir / "按行数-大表订单.xlsx")
    row_split = _split_stats(big_table, "rows", rows_per_file=PROD_BIG_ORDER_SPLIT)

    dept_counts = staff_df["部门"].value_counts().to_dict()
    (split_dir / "怎么测.txt").write_text(
        "【02-数据拆分 · 生产级测试数据】\n\n"
        "一、按字段拆分（推荐录视频）\n"
        f"  文件：按部门-员工花名册.xlsx（{len(staff_df)} 行）\n"
        "  方式：按字段 → 部门\n"
        f"  自测：拆成 {staff_split.file_count} 个文件（"
        + "、".join(f"{k}{v}人" for k, v in sorted(dept_counts.items()))
        + "）\n\n"
        "二、按工作表拆分\n"
        f"  文件：按工作表-区域销售.xlsx（{len(SPLIT_SHEETS)} 个 Sheet，每表约 {PROD_SHEET_ROWS} 行）\n"
        f"  自测：拆成 {sheet_split.file_count} 个 xlsx\n\n"
        "三、按行数拆分\n"
        f"  文件：按行数-大表订单.xlsx（{len(big_orders)} 行）\n"
        f"  方式：按行数 → 每个文件 {PROD_BIG_ORDER_SPLIT} 行\n"
        f"  自测：拆成 {row_split.file_count} 个文件\n\n"
        "适合验证：五千/万级拆分、多 Sheet、大表切分。",
        encoding="utf-8",
    )

    clean_rows = _build_clean_rows(PROD_CLEAN_BASE)
    customers = _save(pd.DataFrame(clean_rows), clean_dir / "客户名单.xlsx")

    sales_rows = [
        {"日期": date(2025, 1, 8), "销售人员": "张明", "产品": "护手霜", "数量": 12, "销售额": 960},
        {"日期": date(2025, 1, 15), "销售人员": "李雪", "产品": "面膜", "数量": 20, "销售额": 2400},
        {"日期": date(2025, 1, 22), "销售人员": "王鹏", "产品": "精华液", "数量": 6, "销售额": 1800},
        {"日期": date(2025, 1, 28), "销售人员": "赵倩", "产品": "口红", "数量": 15, "销售额": 2250},
        {"日期": date(2025, 2, 5), "销售人员": "张明", "产品": "面膜", "数量": 18, "销售额": 2160},
        {"日期": date(2025, 2, 12), "销售人员": "李雪", "产品": "精华液", "数量": 8, "销售额": 2400},
        {"日期": date(2025, 2, 18), "销售人员": "王鹏", "产品": "护手霜", "数量": 30, "销售额": 2400},
        {"日期": date(2025, 2, 25), "销售人员": "赵倩", "产品": "眼影盘", "数量": 10, "销售额": 1990},
        {"日期": date(2025, 3, 3), "销售人员": "张明", "产品": "口红", "数量": 9, "销售额": 1350},
        {"日期": date(2025, 3, 9), "销售人员": "李雪", "产品": "护手霜", "数量": 25, "销售额": 2000},
        {"日期": date(2025, 3, 16), "销售人员": "王鹏", "产品": "面膜", "数量": 14, "销售额": 1680},
        {"日期": date(2025, 3, 21), "销售人员": "赵倩", "产品": "精华液", "数量": 5, "销售额": 1500},
        {"日期": date(2025, 4, 4), "销售人员": "张明", "产品": "眼影盘", "数量": 7, "销售额": 1393},
        {"日期": date(2025, 4, 11), "销售人员": "李雪", "产品": "口红", "数量": 11, "销售额": 1650},
        {"日期": date(2025, 4, 19), "销售人员": "王鹏", "产品": "精华液", "数量": 9, "销售额": 2700},
        {"日期": date(2025, 4, 27), "销售人员": "赵倩", "产品": "面膜", "数量": 16, "销售额": 1920},
        {"日期": date(2025, 5, 6), "销售人员": "张明", "产品": "护手霜", "数量": 22, "销售额": 1760},
        {"日期": date(2025, 5, 14), "销售人员": "李雪", "产品": "眼影盘", "数量": 8, "销售额": 1592},
        {"日期": date(2025, 5, 20), "销售人员": "王鹏", "产品": "口红", "数量": 13, "销售额": 1950},
        {"日期": date(2025, 5, 29), "销售人员": "赵倩", "产品": "护手霜", "数量": 19, "销售额": 1520},
        {"日期": date(2025, 6, 2), "销售人员": "张明", "产品": "面膜", "数量": 21, "销售额": 2520},
        {"日期": date(2025, 6, 10), "销售人员": "李雪", "产品": "精华液", "数量": 7, "销售额": 2100},
        {"日期": date(2025, 6, 18), "销售人员": "王鹏", "产品": "眼影盘", "数量": 6, "销售额": 1194},
        {"日期": date(2025, 6, 26), "销售人员": "赵倩", "产品": "口红", "数量": 18, "销售额": 2700},
    ]
    sales = _save(pd.DataFrame(sales_rows), sales_dir / "销售明细.xlsx")

    pool_old_df, pool_new_df = _build_compare_customer_pool()
    pool_old = _save(pool_old_df, compare_dir / "客户池_旧版.xlsx")
    pool_new = _save(pool_new_df, compare_dir / "客户池_新版.xlsx")
    pool_stats = _compare_stats(pool_old, pool_new, ["手机号"])

    order_old_df, order_new_df = _build_compare_orders()
    order_old = _save(order_old_df, compare_dir / "订单明细_旧版.xlsx")
    order_new = _save(order_new_df, compare_dir / "订单明细_新版.xlsx")
    order_stats = _compare_stats(order_old, order_new, ["订单号", "商品编码"])

    _write_compare_guide(
        compare_dir / "怎么测.txt",
        "【05-两表对比 · 生产级测试数据】",
        [
            "一、客户池（推荐录视频）\n"
            f"  表A：客户池_旧版.xlsx（{len(pool_old_df)} 行）\n"
            f"  表B：客户池_新版.xlsx（{len(pool_new_df)} 行）\n"
            "  主键：手机号\n"
            f"  自测：仅在A {pool_stats.only_a}｜仅在B {pool_stats.only_b}｜"
            f"有变化 {pool_stats.changed}｜完全相同 {pool_stats.same}\n"
            "  场景：八百级客户池、删增改、重复主键、空手机号。",
            "二、订单明细（组合主键）\n"
            f"  表A/B：各 {len(order_old_df)} 行\n"
            "  主键：订单号 + 商品编码\n"
            f"  自测：仅在A {order_stats.only_a}｜仅在B {order_stats.only_b}｜"
            f"有变化 {order_stats.changed}｜完全相同 {order_stats.same}\n"
            "  两千行级订单、一单多商品，必须用组合主键。",
            "三、操作\n"
            "  Excel 两表对比 → 选表 → 选主键 → 开始对比 → 看摘要/仅在A/B/有变化",
        ],
    )

    return {
        **merge_files,
        "staff": staff,
        "multi_sheet": multi_sheet,
        "big_table": big_table,
        "customers": customers,
        "sales": sales,
        "pool_old": pool_old,
        "pool_new": pool_new,
        "order_old": order_old,
        "order_new": order_new,
    }


def copy_to_desktop(generated: dict[str, Path]) -> Path:
    if DESKTOP_DIR.exists():
        shutil.rmtree(DESKTOP_DIR)
    shutil.copytree(CASES_DIR, DESKTOP_DIR)
    return DESKTOP_DIR


if __name__ == "__main__":
    import time

    t0 = time.perf_counter()
    files = generate_cases()
    desktop = copy_to_desktop(files)
    elapsed = time.perf_counter() - t0
    print(f"案例目录: {CASES_DIR}")
    print(f"桌面副本: {desktop}")
    print(f"生成耗时: {elapsed:.1f}s")
    for name, path in files.items():
        print(f"{name}: {path}")
    clean_path = files["customers"]
    df = pd.read_excel(clean_path)
    print(f"清洗案例行数: {len(df)}")
    print(f"列: {list(df.columns)}")
