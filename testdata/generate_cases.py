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


def generate_cases(root: Path | None = None) -> dict[str, Path]:
    cases = root or CASES_DIR
    if cases.exists():
        shutil.rmtree(cases)
    cases.mkdir(parents=True, exist_ok=True)

    merge_dir = cases / "01-批量合并"
    split_dir = cases / "02-按部门拆分"
    clean_dir = cases / "03-客户名单清洗"
    sales_dir = cases / "04-销售数据汇总"
    compare_dir = cases / "05-两表对比"

    east = _save(
        pd.DataFrame(
            [
                {"姓名": "张三", "手机号": "13800138001", "销售额": 12800},
                {"姓名": "李四", "手机号": "13800138002", "销售额": 9600},
                {"姓名": "王五", "手机号": "13800138003", "销售额": 15200},
                {"姓名": "赵六", "手机号": "13800138004", "销售额": 7800},
            ]
        ),
        merge_dir / "华东门店.xlsx",
    )
    south = _save(
        pd.DataFrame(
            [
                {"销售额": 21000, "姓名": "孙七", "手机号": "13900139001"},
                {"销售额": 13400, "姓名": "周八", "手机号": "13900139002"},
                {"销售额": 8900, "姓名": "吴九", "手机号": "13900139003"},
            ]
        ),
        merge_dir / "华南门店.xlsx",
    )
    north = _save(
        pd.DataFrame(
            [
                {"姓名": "郑十", "销售额": 17600},
                {"姓名": "钱十一", "销售额": 6400},
                {"姓名": "冯十二", "销售额": 11200},
            ]
        ),
        merge_dir / "华北门店.xlsx",
    )

    staff = _save(
        pd.DataFrame(
            [
                {"姓名": "陈晨", "部门": "华东", "岗位": "销售", "入职日期": date(2023, 3, 1), "销售额": 18000},
                {"姓名": "林小雨", "部门": "华南", "岗位": "销售", "入职日期": date(2022, 11, 12), "销售额": 22000},
                {"姓名": "黄伟", "部门": "华东", "岗位": "主管", "入职日期": date(2021, 6, 8), "销售额": 31000},
                {"姓名": "周杰", "部门": "华北", "岗位": "销售", "入职日期": date(2024, 1, 15), "销售额": 9600},
                {"姓名": "吴敏", "部门": "华南", "岗位": "客服", "入职日期": date(2023, 8, 20), "销售额": 0},
                {"姓名": "徐丽", "部门": "华东", "岗位": "销售", "入职日期": date(2024, 4, 2), "销售额": 12500},
                {"姓名": "孙浩", "部门": "西南", "岗位": "销售", "入职日期": date(2022, 5, 18), "销售额": 14700},
                {"姓名": "马芳", "部门": "华北", "岗位": "销售", "入职日期": date(2023, 9, 9), "销售额": 20300},
                {"姓名": "高强", "部门": "西南", "岗位": "主管", "入职日期": date(2020, 2, 1), "销售额": 28600},
                {"姓名": "何静", "部门": "华南", "岗位": "销售", "入职日期": date(2024, 7, 1), "销售额": 5400},
                {"姓名": "罗斌", "部门": "华北", "岗位": "客服", "入职日期": date(2021, 12, 3), "销售额": 0},
                {"姓名": "梁娜", "部门": "西南", "岗位": "销售", "入职日期": date(2023, 10, 21), "销售额": 16800},
            ]
        ),
        split_dir / "员工花名册.xlsx",
    )

    clean_rows = _build_clean_rows(1200)
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

    compare_old = _save(
        pd.DataFrame(
            [
                {"手机号": "13800138001", "姓名": "张三", "城市": "上海", "会员等级": "银卡", "备注": "老客户"},
                {"手机号": "13800138002", "姓名": "李四", "城市": "北京", "会员等级": "普通", "备注": ""},
                {"手机号": "13800138003", "姓名": "王五", "城市": "广州", "会员等级": "金卡", "备注": "要跟进"},
                {"手机号": "13800138004", "姓名": "赵六", "城市": "深圳", "会员等级": "普通", "备注": "仅在旧表"},
                {"手机号": "13800138005", "姓名": "钱七", "城市": "杭州", "会员等级": "钻石", "备注": "会改城市"},
            ]
        ),
        compare_dir / "客户名单_旧版.xlsx",
    )
    compare_new = _save(
        pd.DataFrame(
            [
                {"手机号": "13800138001", "姓名": "张三", "城市": "杭州", "会员等级": "银卡", "备注": "老客户"},
                {"手机号": "13800138002", "姓名": "李四", "城市": "北京", "会员等级": "普通", "备注": ""},
                {"手机号": "13800138003", "姓名": "王五", "城市": "广州", "会员等级": "钻石", "备注": "要跟进"},
                {"手机号": "13800138005", "姓名": "钱七", "城市": "宁波", "会员等级": "钻石", "备注": "会改城市"},
                {"手机号": "13800138006", "姓名": "孙八", "城市": "苏州", "会员等级": "普通", "备注": "新客户"},
            ]
        ),
        compare_dir / "客户名单_新版.xlsx",
    )
    (compare_dir / "怎么测.txt").write_text(
        "表A选旧版，表B选新版，主键选手机号。\n"
        "预期：仅在A=赵六；仅在B=孙八；有变化=张三/王五/钱七；相同=李四。\n",
        encoding="utf-8",
    )

    return {
        "east": east,
        "south": south,
        "north": north,
        "staff": staff,
        "customers": customers,
        "sales": sales,
        "compare_old": compare_old,
        "compare_new": compare_new,
    }


def copy_to_desktop(generated: dict[str, Path]) -> Path:
    if DESKTOP_DIR.exists():
        shutil.rmtree(DESKTOP_DIR)
    shutil.copytree(CASES_DIR, DESKTOP_DIR)
    return DESKTOP_DIR


if __name__ == "__main__":
    files = generate_cases()
    desktop = copy_to_desktop(files)
    print(f"案例目录: {CASES_DIR}")
    print(f"桌面副本: {desktop}")
    for name, path in files.items():
        print(f"{name}: {path}")
    clean_path = files["customers"]
    df = pd.read_excel(clean_path)
    print(f"清洗案例行数: {len(df)}")
    print(f"列: {list(df.columns)}")
