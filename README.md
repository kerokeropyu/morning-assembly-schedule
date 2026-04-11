# 朝礼当番表 自動生成スクリプト

Excelファイルに月ごとのシートを自動追加します。

## セットアップ

```bash
pip install openpyxl jpholiday
# または uv を使う場合
uv pip install openpyxl jpholiday
```

## 実行方法

```bash
# 翌月分を1枚追加
python generate_schedule.py

# 指定年月のシートを1枚追加
python generate_schedule.py 2026 5
```

## 設定ファイル（morning_assembly_config.json）

| キー | 説明 |
|---|---|
| `members` | メンバー名の配列 |
| `output_directory` | 出力先フォルダ |
| `excel_filename` | 出力ファイル名 |
| `template_file` | テンプレートExcelのパス（省略可） |
| `template_sheet_name` | 雛形シート名（省略可） |
| `start_row` | 描画開始行番号（デフォルト: 1） |
| `start_col` | 描画開始列番号（デフォルト: 1） |
