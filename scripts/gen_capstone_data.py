"""
第4章「月次売上レポート自動生成」用サンプルデータ生成スクリプト

生成物:
  data/capstone/sales_2026-01/sales_2026-01_<code>_<name>.xlsx ... 47ファイル (各 1,000 行)
  data/capstone/master_products.xlsx                              商品マスタ (6 行)
  data/capstone/master_branches.xlsx                              支店マスタ (47 行)

再現性のため random seed 固定。
"""
from __future__ import annotations

import random
from pathlib import Path

import pandas as pd

REPO = Path(__file__).resolve().parent.parent
OUT_DIR = REPO / "data" / "capstone"
SALES_DIR = OUT_DIR / "sales_2026-01"

ROWS_PER_FILE = 1000
RANDOM_SEED = 20260101

# ----- 商品マスタ -----
PRODUCTS = [
    # (code, name, category, unit_price, cost)
    ("P001", "りんご (1袋)", "定番果物", 480, 280),
    ("P002", "ばなな (1房)", "定番果物", 280, 130),
    ("P003", "みかん (1袋)", "定番果物", 380, 200),
    ("P004", "メロン", "高級果物", 2800, 1500),
    ("P005", "ぶどう (1パック)", "高級果物", 980, 550),
    ("P006", "いちご (1パック)", "高級果物", 680, 380),
]

# ----- 支店マスタ -----
# (都道府県コード, 英字名 (ローマ字), 都道府県名, 地域ブロック, 売上規模重み)
PREFECTURES = [
    ("01", "hokkaido", "北海道",   "北海道",   1.3),
    ("02", "aomori",   "青森県",   "東北",     0.5),
    ("03", "iwate",    "岩手県",   "東北",     0.5),
    ("04", "miyagi",   "宮城県",   "東北",     0.9),
    ("05", "akita",    "秋田県",   "東北",     0.4),
    ("06", "yamagata", "山形県",   "東北",     0.5),
    ("07", "fukushima","福島県",   "東北",     0.7),
    ("08", "ibaraki",  "茨城県",   "関東",     1.0),
    ("09", "tochigi",  "栃木県",   "関東",     0.8),
    ("10", "gunma",    "群馬県",   "関東",     0.8),
    ("11", "saitama",  "埼玉県",   "関東",     1.7),
    ("12", "chiba",    "千葉県",   "関東",     1.5),
    ("13", "tokyo",    "東京都",   "関東",     3.0),
    ("14", "kanagawa", "神奈川県", "関東",     2.0),
    ("15", "niigata",  "新潟県",   "中部",     0.8),
    ("16", "toyama",   "富山県",   "中部",     0.5),
    ("17", "ishikawa", "石川県",   "中部",     0.5),
    ("18", "fukui",    "福井県",   "中部",     0.4),
    ("19", "yamanashi","山梨県",   "中部",     0.4),
    ("20", "nagano",   "長野県",   "中部",     0.7),
    ("21", "gifu",     "岐阜県",   "中部",     0.7),
    ("22", "shizuoka", "静岡県",   "中部",     1.1),
    ("23", "aichi",    "愛知県",   "中部",     2.0),
    ("24", "mie",      "三重県",   "中部",     0.7),
    ("25", "shiga",    "滋賀県",   "近畿",     0.6),
    ("26", "kyoto",    "京都府",   "近畿",     1.0),
    ("27", "osaka",    "大阪府",   "近畿",     2.2),
    ("28", "hyogo",    "兵庫県",   "近畿",     1.5),
    ("29", "nara",     "奈良県",   "近畿",     0.5),
    ("30", "wakayama", "和歌山県", "近畿",     0.4),
    ("31", "tottori",  "鳥取県",   "中国",     0.3),
    ("32", "shimane",  "島根県",   "中国",     0.3),
    ("33", "okayama",  "岡山県",   "中国",     0.7),
    ("34", "hiroshima","広島県",   "中国",     0.9),
    ("35", "yamaguchi","山口県",   "中国",     0.5),
    ("36", "tokushima","徳島県",   "四国",     0.4),
    ("37", "kagawa",   "香川県",   "四国",     0.5),
    ("38", "ehime",    "愛媛県",   "四国",     0.6),
    ("39", "kochi",    "高知県",   "四国",     0.4),
    ("40", "fukuoka",  "福岡県",   "九州・沖縄", 1.4),
    ("41", "saga",     "佐賀県",   "九州・沖縄", 0.4),
    ("42", "nagasaki", "長崎県",   "九州・沖縄", 0.6),
    ("43", "kumamoto", "熊本県",   "九州・沖縄", 0.7),
    ("44", "oita",     "大分県",   "九州・沖縄", 0.5),
    ("45", "miyazaki", "宮崎県",   "九州・沖縄", 0.5),
    ("46", "kagoshima","鹿児島県", "九州・沖縄", 0.6),
    ("47", "okinawa",  "沖縄県",   "九州・沖縄", 0.5),
]


def make_products_master() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"product_code": p[0], "product_name": p[1], "category": p[2],
             "unit_price": p[3], "cost": p[4]}
            for p in PRODUCTS
        ]
    )


def make_branches_master() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"prefecture_code": p[0], "prefecture_name": p[2], "region": p[3]}
            for p in PREFECTURES
        ]
    )


def make_sales_for_prefecture(rng: random.Random, weight: float) -> pd.DataFrame:
    """1 都道府県・1,000 行の売上データを作る"""
    rows = []
    days = pd.date_range("2026-01-01", "2026-01-31", freq="D")
    # 商品の出現率（定番が多め）
    product_weights = [3.0, 3.0, 3.0, 0.5, 1.0, 2.0]   # 順番は PRODUCTS と対応

    for _ in range(ROWS_PER_FILE):
        date = rng.choice(days)
        p = rng.choices(PRODUCTS, weights=product_weights, k=1)[0]
        code, name, _cat, unit_price, _cost = p
        # 数量: 重み高い県ほど多めに買う傾向、メロンなど高単価は数量少なめ
        base_qty = max(1, int(rng.gauss(3.5 * weight, 1.5)))
        if code == "P004":   # メロンは控えめ
            base_qty = max(1, base_qty // 2)
        qty = max(1, base_qty)
        rows.append({
            "date": date,
            "product_code": code,
            "product_name": name,
            "quantity": qty,
            "unit_price": unit_price,
            "amount": qty * unit_price,
        })
    df = pd.DataFrame(rows).sort_values(["date", "product_code"]).reset_index(drop=True)
    return df


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    SALES_DIR.mkdir(parents=True, exist_ok=True)

    # マスタ
    make_products_master().to_excel(OUT_DIR / "master_products.xlsx", sheet_name="商品", index=False)
    make_branches_master().to_excel(OUT_DIR / "master_branches.xlsx", sheet_name="支店", index=False)
    print(f"wrote: master_products.xlsx, master_branches.xlsx")

    # 各都道府県の売上
    rng = random.Random(RANDOM_SEED)
    for code, romaji, name, _region, weight in PREFECTURES:
        df = make_sales_for_prefecture(rng, weight)
        fname = f"sales_2026-01_{code}_{romaji}.xlsx"
        df.to_excel(SALES_DIR / fname, sheet_name="売上", index=False)
        print(f"wrote: {fname}  (rows={len(df)}, weight={weight})")

    print(f"\nDone. Generated {len(PREFECTURES)} files in {SALES_DIR}")


if __name__ == "__main__":
    main()
