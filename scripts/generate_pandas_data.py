"""
2-2 用のサンプルデータを生成するスクリプト。
講座中は学生は使わない（生徒は出来上がった CSV/xlsx をそのまま読む）。

実行:
    python3 scripts/generate_pandas_data.py
"""
import os
import random
from datetime import date

import pandas as pd

random.seed(42)
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(HERE, "data", "pandas")
os.makedirs(OUT, exist_ok=True)


def make_employees_csv() -> None:
    employees = [
        {"id": 1, "name": "田中 太郎", "department": "営業1課", "salary": 350000, "hire_date": "2018-04-01"},
        {"id": 2, "name": "鈴木 花子", "department": "営業2課", "salary": 420000, "hire_date": "2015-10-01"},
        {"id": 3, "name": "佐藤 一郎", "department": "営業1課", "salary": 380000, "hire_date": "2019-04-01"},
        {"id": 4, "name": "山田 美咲", "department": "開発", "salary": 480000, "hire_date": "2014-04-01"},
        {"id": 5, "name": "高橋 健", "department": "総務", "salary": 320000, "hire_date": "2020-07-15"},
        {"id": 6, "name": "伊藤 春香", "department": "開発", "salary": 510000, "hire_date": "2012-04-01"},
        {"id": 7, "name": "渡辺 大輔", "department": "営業2課", "salary": 360000, "hire_date": "2021-04-01"},
        {"id": 8, "name": "中村 さくら", "department": "営業1課", "salary": 395000, "hire_date": "2017-09-01"},
        {"id": 9, "name": "小林 翔", "department": "開発", "salary": 440000, "hire_date": "2016-04-01"},
        {"id": 10, "name": "加藤 美穂", "department": "総務", "salary": 305000, "hire_date": "2022-04-01"},
    ]
    df = pd.DataFrame(employees)
    out_path = os.path.join(OUT, "employees.csv")
    df.to_csv(out_path, index=False, encoding="utf-8")
    print(f"wrote {out_path} ({len(df)} rows)")


def make_sales_xlsx() -> None:
    products = [
        ("P001", "りんご (1袋)", 480),
        ("P002", "ばなな (1房)", 280),
        ("P003", "みかん (1袋)", 380),
        ("P004", "メロン", 2800),
        ("P005", "ぶどう (1パック)", 980),
        ("P006", "いちご (1パック)", 680),
    ]
    rows = []
    for day in range(1, 31):
        # 1日あたり 1〜3 件のランダムな取引
        n_trans = random.randint(1, 3)
        for _ in range(n_trans):
            code, name, unit_price = random.choice(products)
            qty = random.randint(1, 12)
            rows.append({
                "date": date(2026, 1, day),
                "product_code": code,
                "product_name": name,
                "quantity": qty,
                "unit_price": unit_price,
                "amount": qty * unit_price,
            })
    df = pd.DataFrame(rows)
    out_path = os.path.join(OUT, "sales_2026-01.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="売上", index=False)
    print(f"wrote {out_path} ({len(df)} rows)")


if __name__ == "__main__":
    make_employees_csv()
    make_sales_xlsx()
