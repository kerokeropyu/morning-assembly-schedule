# 総務部 朝礼当番表 自動生成システム

## 概要

毎月の朝礼当番表を自動で生成するPythonスクリプトです。  
月末にバッチを実行すると、**1つのExcelファイルの中に翌月のシート（タブ）が1枚追加**されていきます。

```
総務部朝礼当番表.xlsx
  ├── 2026年04月  ← シート（タブ）
  ├── 2026年05月
  └── 2026年06月
```

## ファイル構成

```
.
├── generate_schedule.py      # メインスクリプト
├── morning_assembly_config.json  # 設定ファイル（メンバー管理）
├── run_monthly_batch.bat     # Windows用バッチ
├── run_monthly_batch.sh      # Linux/Mac用シェル
├── requirements.txt          # 必要ライブラリ
├── README.md                 # このファイル
├── README_for_staff.txt      # 事務員さん向け設定変更マニュアル
└── morning_assembly_schedules/   # Excel出力先（自動生成）
    └── 総務部朝礼当番表.xlsx
```

## セットアップ

```bash
pip install -r requirements.txt
```

## 使い方

### 翌月分を自動生成（月末バッチ用）

```bash
# Windows
run_monthly_batch.bat

# Linux / Mac
chmod +x run_monthly_batch.sh  # 初回のみ
./run_monthly_batch.sh

# Python直接実行
python generate_schedule.py
```

### 特定の月を手動生成

```bash
python generate_schedule.py 2026 4   # 2026年4月
python generate_schedule.py 2026 5   # 2026年5月
```

## 設定ファイル（morning_assembly_config.json）

```json
{
  "department_name": "総務部",
  "members": [
    "田中 太郎",
    "佐藤 花子",
    "鈴木 一郎",
    "高橋 美咲",
    "伊藤 健太"
  ],
  "output_directory": "./morning_assembly_schedules",
  "excel_filename": "総務部朝礼当番表.xlsx"
}
```

| 項目 | 説明 |
|------|------|
| `department_name` | 部署名 |
| `members` | 当番メンバーのリスト（この順番でローテーション） |
| `output_directory` | Excelの保存先フォルダ |
| `excel_filename` | Excelのファイル名（固定） |

メンバーの追加・削除方法は `README_for_staff.txt` を参照してください。

## 定期実行の設定

### Windows タスクスケジューラ

1. タスクスケジューラを開く
2. 「基本タスクの作成」を選択
3. トリガー：毎月最終日の23:00
4. 操作：`run_monthly_batch.bat` を実行

### Linux/Mac cron

```bash
crontab -e

# 毎月末日の23:00に実行
0 23 28-31 * * [ $(date -d tomorrow +\%d) -eq 1 ] && cd /path/to/script && ./run_monthly_batch.sh
```

## Excelの出力仕様

- **ファイル名**: 設定ファイルで指定した名前（固定）
- **シート名**: `2026年04月` のような形式
- **レイアウト**: 左列にメンバー名、上段に日付と曜日、当番の日に赤い◯マーク
- **当番の割り振り**: メンバーリストの順番で1日から順にローテーション
