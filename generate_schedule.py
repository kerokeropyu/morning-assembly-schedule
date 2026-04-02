#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
総務部 朝礼当番表 自動生成スクリプト
・1つのExcelファイルの中に、月ごとにシート（タブ）を追加していく方式
・ファイル名は固定（設定ファイルで指定）
・土日・祝日は当番なし（セルをグレーにして「休」と表示）
・通常実行：翌月分のシートを1枚追加
・引数あり：指定年月のシートを1枚追加（例: python generate_schedule.py 2026 4）
"""

import json
import sys
from datetime import date, datetime
from pathlib import Path
from calendar import monthrange

import jpholiday
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import openpyxl


def load_config(config_path: str = "morning_assembly_config.json"):
    """設定ファイルを読み込む"""
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"エラー: 設定ファイル '{config_path}' が見つかりません。")
        sys.exit(1)
    except json.JSONDecodeError:
        print(f"エラー: 設定ファイル '{config_path}' の内容が正しくありません（JSONエラー）。")
        sys.exit(1)


# =============================================================================
# 【テンプレート対応 修正箇所①】
# load_or_create_workbook() を以下のように変更する
#
# 変更前（現在）:
#   出力用Excelファイルが既にあれば開く、なければ空のWorkbookを新規作成する
#
# 変更後（テンプレート対応）:
#   出力用Excelファイルが既にあれば開く
#   なければ「テンプレートファイル」をコピーして新規作成する
#
# 具体的には、引数に template_path を追加して以下のように書き換える:
#
#   def load_or_create_workbook(output_dir: Path, filename: str, template_path: str):
#       output_dir.mkdir(parents=True, exist_ok=True)
#       filepath = output_dir / filename
#       if filepath.exists():
#           wb = load_workbook(filepath)
#       else:
#           # テンプレートファイルを読み込んで雛形として使う
#           wb = load_workbook(template_path)
#       return wb, filepath
#
# template_path は設定ファイル（morning_assembly_config.json）に
# "template_file": "./template.xlsx" のように追加して外部指定できるようにする
# （exe化したときにテンプレートを差し替えられるようにするため）
# =============================================================================
def load_or_create_workbook(output_dir: Path, filename: str):
    """Excelファイルを読み込む。なければ新規作成。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    filepath = output_dir / filename
    if filepath.exists():
        wb = load_workbook(filepath)
    else:
        wb = Workbook()
    return wb, filepath


def is_off_day(year: int, month: int, day: int) -> bool:
    """土日または祝日かどうか判定する"""
    d = date(year, month, day)
    return d.weekday() >= 5 or jpholiday.is_holiday(d)


# =============================================================================
# 【テンプレート対応 修正箇所②】
# create_month_sheet() の冒頭部分を以下のように変更する
#
# 変更前（現在）:
#   wb.create_sheet() で空のシートを新規作成し、スタイルをPythonコードで一から設定している
#
# 変更後（テンプレート対応）:
#   テンプレートに既に存在するシートをコピーして雛形として使い、そこに書き込む
#   例:
#     template_sheet_name = config.get("template_sheet_name", "テンプレート")
#     template_ws = wb[template_sheet_name]  # テンプレート内の雛形シートを取得
#     ws = wb.copy_worksheet(template_ws)    # コピーして新しいシートを作成
#     ws.title = sheet_title                 # シート名を「2026年04月」等に変更
#
# ※ テンプレートシート自体はユーザーが見えないように非表示にしておくとよい
#   ws.sheet_state = 'hidden'  # テンプレートシートを隠す場合
# =============================================================================


# =============================================================================
# 【テンプレート対応 修正箇所③】
# 描画の「始点」（書き込みを開始する行・列）を変数化する
#
# 変更前（現在）:
#   ヘッダーは row=1, column=1 から、名前は row=2 から決め打ちで書き込んでいる
#
# 変更後（テンプレート対応）:
#   テンプレートに既存のタイトルや装飾がある場合、
#   それを避けた行・列から書き始めたい
#   設定ファイルに以下のような項目を追加して外部から指定できるようにする:
#
#     "start_row": 3   # ヘッダー行を何行目から書き始めるか（デフォルト: 1）
#     "start_col": 2   # 名前列を何列目から書き始めるか（デフォルト: 1）
#
#   コード内では:
#     START_ROW = config.get("start_row", 1)
#     START_COL = config.get("start_col", 1)
#
#   そして以下の各行の row/column の数値を
#   START_ROW / START_COL を使った計算式に置き換える:
#
#     ws.cell(row=1,          column=1)          → ws.cell(row=START_ROW,       column=START_COL)
#     ws.cell(row=1,          column=day + 1)    → ws.cell(row=START_ROW,       column=START_COL + day)
#     ws.cell(row=idx,        column=1)          → ws.cell(row=START_ROW + idx, column=START_COL)
#     ws.cell(row=duty_row,   column=col)        → そのまま（duty_rowの算出式をSTART_ROW基準に変更）
#     ws.column_dimensions["A"]                  → get_column_letter(START_COL) に変更
#     ws.row_dimensions[1]                       → ws.row_dimensions[START_ROW] に変更
# =============================================================================
def create_month_sheet(wb, year: int, month: int, members):
    """指定年月のシートを新規作成（既にあれば作り直し）"""
    sheet_title = f"{year}年{month:02d}月"

    if sheet_title in wb.sheetnames:
        wb.remove(wb[sheet_title])

    ws = wb.create_sheet(title=sheet_title)

    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        try:
            wb.remove(wb["Sheet"])
        except KeyError:
            pass

    # スタイル定義
    header_fill  = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")  # 平日ヘッダー：ネイビー
    sat_fill     = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")  # 土曜：水色
    sun_hol_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")  # 日曜・祝日：薄赤
    off_fill     = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")  # 休日セル：グレー
    name_fill    = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")  # 名前列
    header_font  = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    center_align = Alignment(horizontal="center", vertical="center")

    num_members = len(members)
    _, num_days = monthrange(year, month)
    youbi = ["月", "火", "水", "木", "金", "土", "日"]

    # ヘッダー（左上セル）
    # 【修正箇所③】row=1, column=1 → row=START_ROW, column=START_COL に置き換える
    c = ws.cell(row=1, column=1, value="氏名 / 日付")
    c.fill = header_fill
    c.font = header_font
    c.border = border
    c.alignment = center_align

    # 日付行：平日はネイビー、土曜は水色、日曜・祝日は薄赤
    # 【修正箇所③】row=1, column=day + 1 → row=START_ROW, column=START_COL + day に置き換える
    for day in range(1, num_days + 1):
        d = date(year, month, day)
        w = youbi[d.weekday()]
        cell = ws.cell(row=1, column=day + 1, value=f"{day}日\n({w})")
        cell.font = header_font
        cell.border = border
        cell.alignment = center_align
        if d.weekday() == 5:
            cell.fill = sat_fill
        elif d.weekday() == 6 or jpholiday.is_holiday(d):
            cell.fill = sun_hol_fill
        else:
            cell.fill = header_fill

    # 名前列
    # 【修正箇所③】row=idx, column=1 → row=START_ROW + idx, column=START_COL に置き換える
    for idx, member in enumerate(members, start=2):
        cell = ws.cell(row=idx, column=1, value=member)
        cell.fill = name_fill
        cell.font = Font(bold=True, size=10)
        cell.border = border
        cell.alignment = Alignment(horizontal="left", vertical="center")

    # 当番 ◯ 割り振り
    # 土日祝日は「休」+グレー、平日は当番のメンバーだけ◯
    # 【修正箇所③】duty_row の算出式を START_ROW 基準に変更する
    duty_index = 0
    for day in range(1, num_days + 1):
        col = day + 1  # 【修正箇所③】col = START_COL + day に置き換える
        if is_off_day(year, month, day):
            # 休日：全メンバー行を「休」+グレー
            for row in range(2, num_members + 2):  # 【修正箇所③】range(START_ROW + 1, START_ROW + num_members + 1) に置き換える
                cell = ws.cell(row=row, column=col, value="休")
                cell.fill = off_fill
                cell.font = Font(size=10, color="FFFFFF")
                cell.border = border
                cell.alignment = center_align
        else:
            # 平日：当番のメンバーだけ◯、他は空白
            duty_row = (duty_index % num_members) + 2  # 【修正箇所③】+ 2 → + START_ROW + 1 に置き換える
            for row in range(2, num_members + 2):      # 【修正箇所③】range(START_ROW + 1, START_ROW + num_members + 1) に置き換える
                cell = ws.cell(row=row, column=col, value="")
                cell.border = border
                cell.alignment = center_align
            duty_cell = ws.cell(row=duty_row, column=col, value="◯")
            duty_cell.font = Font(size=14, bold=True, color="FF0000")
            duty_cell.border = border
            duty_cell.alignment = center_align
            duty_index += 1

    # 列幅・行高さ
    # 【修正箇所③】"A" → get_column_letter(START_COL) に、[1] → [START_ROW] に置き換える
    ws.column_dimensions["A"].width = 15
    for col in range(2, num_days + 2):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 6
    ws.row_dimensions[1].height = 30
    for row in range(2, num_members + 2):
        ws.row_dimensions[row].height = 20


def generate_schedule_for(year: int, month: int, config: dict):
    """指定年月のシートを1枚追加する（手動実行用）"""
    output_dir = Path(config["output_directory"])
    excel_filename = config["excel_filename"]

    # 【修正箇所①】template_path を config から取得して load_or_create_workbook() に渡す
    # 例: template_path = config.get("template_file", None)
    #     wb, filepath = load_or_create_workbook(output_dir, excel_filename, template_path)
    wb, filepath = load_or_create_workbook(output_dir, excel_filename)

    create_month_sheet(wb, year, month, config["members"])
    wb.save(filepath)
    print("============================================")
    print("朝礼当番表シートを作成しました")
    print(f"  ファイル: {filepath}")
    print(f"  シート名: {year}年{month:02d}月")
    print(f"  メンバー数: {len(config['members'])} 名")
    print("============================================")


def generate_next_month(config: dict):
    """翌月分のシートを1枚追加する（月末バッチ用）"""
    today = datetime.now()
    if today.month == 12:
        year = today.year + 1
        month = 1
    else:
        year = today.year
        month = today.month + 1
    generate_schedule_for(year, month, config)


if __name__ == "__main__":
    config = load_config()
    # 使い方：
    #   python generate_schedule.py           → 翌月分を1枚追加
    #   python generate_schedule.py 2026 4   → 2026年4月分を1枚追加
    if len(sys.argv) >= 3:
        y = int(sys.argv[1])
        m = int(sys.argv[2])
        generate_schedule_for(y, m, config)
    else:
        generate_next_month(config)
