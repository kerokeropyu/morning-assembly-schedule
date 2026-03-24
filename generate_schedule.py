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
        print(f"エラー: 設定ファイル '{config_path}' の内容が正しくありません（JSONエラー）。"
        sys.exit(1)


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
    c = ws.cell(row=1, column=1, value="氏名 / 日付")
    c.fill = header_fill
    c.font = header_font
    c.border = border
    c.alignment = center_align

    # 日付行：平日はネイビー、土曜は水色、日曜・祝日は薄赤
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
    for idx, member in enumerate(members, start=2):
        cell = ws.cell(row=idx, column=1, value=member)
        cell.fill = name_fill
        cell.font = Font(bold=True, size=10)
        cell.border = border
        cell.alignment = Alignment(horizontal="left", vertical="center")

    # 当番 ◯ 割り振り
    # 土日祝日は「休」+グレー、平日は当番のメンバーだけ◯
    # duty_indexは休日をスキップしてカウント
    duty_index = 0
    for day in range(1, num_days + 1):
        col = day + 1
        if is_off_day(year, month, day):
            # 休日：全メンバー行を「休」+グレー
            for row in range(2, num_members + 2):
                cell = ws.cell(row=row, column=col, value="休")
                cell.fill = off_fill
                cell.font = Font(size=10, color="FFFFFF")
                cell.border = border
                cell.alignment = center_align
        else:
            # 平日：当番のメンバーだけ◯、他は空白
            duty_row = (duty_index % num_members) + 2
            for row in range(2, num_members + 2):
                cell = ws.cell(row=row, column=col, value="")
                cell.border = border
                cell.alignment = center_align
            duty_cell = ws.cell(row=duty_row, column=col, value="◯")
            duty_cell.font = Font(size=14, bold=True, color="FF0000")
            duty_cell.border = border
            duty_cell.alignment = center_align
            duty_index += 1

    # 列幅・行高さ
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
