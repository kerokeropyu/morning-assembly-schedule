# 朝礼当番表 自動生成スクリプト

Excelテンプレートをもとに、月別の掃除担当表シートを自動生成します。

## ファイル構成

```
.
├── generate_schedule.py        # メインスクリプト
├── morning_assembly_config.json # 設定ファイル
├── template.xlsx               # テンプレートExcel（自分で用意・gitignore対象）
└── output/                     # 出力先（自動作成）
    └── 掃除担当表.xlsx
```

## セットアップ

```bash
pip install openpyxl jpholiday
```

## テンプレートExcelの準備

`template.xlsx` を実行ファイルと同じフォルダに配置してください。

テンプレートのレイアウト構成：
| 行 | 内容 |
|---|---|
| 1行目 | タイトル行（「掃除担当表」等）← 上書きしない |
| 2行目 | 週の開始日見出し（「8/7-」等）← 月初から書き込む |
| 3行目 | 列ヘッダー（フロア/番号/掃除場所/担当者）← そのまま残す |
| 4行目以降 | 清掃場所ごとの担当者行 ← 担当者名を書き込む |

列構成（`name_col` を起点に `col_step` おきに担当者列）：
| 列 | 内容 |
|---|---|
| A | フロア |
| B | 番号 |
| C | 掃除場所 |
| D | サポート情報 |
| E | 担当者（1週目）← `name_col: 5` |
| F | 空白列 |
| G | 担当者（2週目）|
| ... | 以降、2列おきに担当者列が続く |

## 設定ファイル（morning_assembly_config.json）

| キー | 説明 | デフォルト |
|---|---|---|
| `members` | メンバー名の配列 | 必須 |
| `template_file` | テンプレートExcelのパス | `./template.xlsx` |
| `template_sheet_name` | 雛形シート名 | 省略時は空シート生成 |
| `week_row` | 週見出しを書く行番号 | `2` |
| `header_row` | 列ヘッダーの行番号 | `3` |
| `data_row` | データ開始行番号 | `4` |
| `name_col` | 担当者列の先頭列番号 | `5` (E列) |
| `col_step` | 担当者列の間隔 | `2` |
| `num_locations` | 清掃場所の行数 | メンバー数と同じ |

## 実行方法

```bash
# 翌月分を1枚追加
python generate_schedule.py

# 指定年月のシートを1枚追加
python generate_schedule.py 2026 5
```

## 個人名の管理について

実際のメンバー名は `morning_assembly_config.json` に記載します。
このファイルは `.gitignore` で管理から除外することを推奨します。

```gitignore
# .gitignore
template.xlsx
morning_assembly_config.json
output/
```
