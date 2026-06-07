# 事務職のための Python 入門

〜Excel・VBA は使えるけど、Python は初めての方へ〜

## この講座について

普段 Excel や VBA で業務を自動化している方が、Python を使って **データの集計・加工・可視化** を一気通貫でできるようになるための講座です。

- **対象**: Python を触ったことがない方（Excel の関数や VBA の経験があるとスムーズです）
- **ゴール**: pandas で Excel/CSV を扱い、matplotlib でグラフ化、最終的に「月次レポート自動生成」レベルの実務スクリプトを書けるようになる
- **環境**: Google Colab（ブラウザだけで OK、インストール不要）
- **形式**: オンラインライブ・全4章構成（1日集中／ご相談の上 1ヶ月以内に分割も可）

## 学習の進め方

教材はすべて Jupyter Notebook (`.ipynb`) 形式で、Google Colab で実行できます。

### セットアップ（最初に1回だけ）

1. 下の目次から **オリエンテーション** の **Open in Colab** をクリック
2. オリエンの最初のコードセル（`drive.mount` + `git clone`）を実行
3. これで `マイドライブ/python-data-basics/` に教材一式がコピーされます

### それ以降の進め方

- 各章のノートブックは、**Google Drive 上のコピー** から開いてください
  （`マイドライブ → python-data-basics → 章フォルダ → .ipynb を右クリック → アプリで開く → Google Colaboratory`）
- Drive 上で開いたノートブックは編集内容が自動保存されます

> 💡 下の目次の「Open in Colab」バッジは **プレビュー用** です。実際に学習するときは、必ず Drive 上のコピーから開いてください。

1つのノートブックは 15〜30 分程度で消化できる粒度に区切ってあります。

## 講座目次

### はじめに

| # | タイトル | Colab |
|---|---|---|
| 0-1 | オリエンテーション | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Tomoaki-Kotsuka/python-data-basics/blob/main/00_introduction/00-1_orientation.ipynb) |

### 第1章 Python の基本

| # | タイトル | Colab |
|---|---|---|
| 1-1 | 変数と型 | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Tomoaki-Kotsuka/python-data-basics/blob/main/01_python_basics/01-1_variables.ipynb) |
| 1-2 | 文字列の操作 | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Tomoaki-Kotsuka/python-data-basics/blob/main/01_python_basics/01-2_strings.ipynb) |
| 1-3 | リストと辞書 | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Tomoaki-Kotsuka/python-data-basics/blob/main/01_python_basics/01-3_list_dict.ipynb) |
| 1-4 | 条件分岐とループ | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Tomoaki-Kotsuka/python-data-basics/blob/main/01_python_basics/01-4_control_flow.ipynb) |
| 1-5 | 関数 | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Tomoaki-Kotsuka/python-data-basics/blob/main/01_python_basics/01-5_functions.ipynb) |

### 第2章 pandas で Excel を操る

| # | タイトル | Colab |
|---|---|---|
| 2-1 | DataFrame 入門 | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Tomoaki-Kotsuka/python-data-basics/blob/main/02_pandas/02-1_dataframe.ipynb) |
| 2-2 | Excel と CSV を読み込む | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Tomoaki-Kotsuka/python-data-basics/blob/main/02_pandas/02-2_read_excel_csv.ipynb) |
| 2-3 | 集計と絞り込み | （準備中） |
| 2-4 | データの加工 | （準備中） |
| 2-5 | Excel への書き出し | （準備中） |

### 第3章 matplotlib で可視化

| # | タイトル | Colab |
|---|---|---|
| 3-1 | 棒グラフ | （準備中） |
| 3-2 | 折れ線と散布図 | （準備中） |
| 3-3 | pandas から直接グラフを描く | （準備中） |
| 3-4 | 体裁を整える（日本語フォント等） | （準備中） |

### 第4章 総集編：月次売上レポート自動生成

| # | タイトル | Colab |
|---|---|---|
| 4-1 | プロジェクトの全体像 | （準備中） |
| 4-2 | 複数ファイルの一括読み込み | （準備中） |
| 4-3 | マスタ結合と集計 | （準備中） |
| 4-4 | グラフ作成 | （準備中） |
| 4-5 | レポート用 Excel への書き出し | （準備中） |
| 4-6 | 全部つないだ完成版 | （準備中） |

## ディレクトリ構成

```
python-data-basics/
├── 00_introduction/    オリエンテーション
├── 01_python_basics/   第1章
├── 02_pandas/          第2章
├── 03_matplotlib/      第3章
├── 04_capstone/        第4章（総集編）
└── data/               講座で使用するサンプルデータ
```

## 質問・サポートについて

教材の内容や講座中の不明点については、ストアカのメッセージ機能からお問い合わせください。
