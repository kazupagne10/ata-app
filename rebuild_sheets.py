"""
スプレッドシートの全シートを削除して正しいヘッダー構成で再構築するスクリプト。
実行: python3 rebuild_sheets.py

戦略:
  1. 新しいシートを先に全て作成する
  2. 旧シートを削除する
  （Googleスプレッドシートは最低1シート必要なため、先に新シートを作ってから旧シートを削除）
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gspread
from ata_extract import (
    get_gspread_client, SPREADSHEET_ID,
    SHEET_MASTER, SHEET_TRADES, SHEET_DETAIL,
    MASTER_HEADERS, _set_dropdown_validation,
)
from config import BID_RESULT_OPTIONS

SHEET_CONDITIONS = "案件サマリー"

CONDITION_HEADERS = [
    "案件ID", "案件名",
    "工事時間帯", "立地タイプ", "工事状態", "指定業者範囲",
    "厨房区画", "室外機設置階数", "照明器具支給", "家具・什器含む",
    "業態", "フロア", "グリストラップ", "排気ダクト",
    "工期", "施工エリア", "防災工事", "その他備考",
]


def rebuild():
    print("Googleスプレッドシートに接続中...")
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    old_worksheets = sh.worksheets()
    old_titles = [ws.title for ws in old_worksheets]
    print(f"現在のシート一覧: {old_titles}")

    # 新しいシート名（作成予定）
    new_titles = [SHEET_MASTER, SHEET_TRADES, SHEET_DETAIL, SHEET_CONDITIONS]

    # ── 新しいシートを先に全て作成 ──
    print("\n新しいシートを作成中...")

    # ①案件マスタ
    title = SHEET_MASTER + "_new"
    ws_master = sh.add_worksheet(title=title, rows=500, cols=len(MASTER_HEADERS) + 5)
    ws_master.append_row(MASTER_HEADERS)
    ws_master.format(f"A1:{chr(64 + len(MASTER_HEADERS))}1", {"textFormat": {"bold": True}})
    _set_dropdown_validation(sh, ws_master.id, col_index=4, values=["新装", "改装"])
    bid_col_idx = MASTER_HEADERS.index("入札結果")
    _set_dropdown_validation(sh, ws_master.id, col_index=bid_col_idx, values=BID_RESULT_OPTIONS)
    print(f"  → {SHEET_MASTER} 作成（{len(MASTER_HEADERS)}列）")

    # ②工種別金額
    trades_headers = [
        "案件ID", "案件名", "業態", "坪数", "工種",
        "金額（税抜）", "坪単価",
        "全体合計（税抜）", "全体坪単価",
        "施工年月", "備考", "ドライブURL",
    ]
    ws_trades = sh.add_worksheet(title=SHEET_TRADES + "_new", rows=1000, cols=len(trades_headers) + 2)
    ws_trades.append_row(trades_headers)
    ws_trades.format("A1:L1", {"textFormat": {"bold": True}})
    print(f"  → {SHEET_TRADES} 作成（{len(trades_headers)}列）")

    # ③見積明細
    detail_headers = [
        "案件ID", "案件名", "業者名", "工種", "大分類",
        "項目名", "仕様・規格", "数量", "単位", "単価", "金額", "備考",
    ]
    ws_detail = sh.add_worksheet(title=SHEET_DETAIL + "_new", rows=2000, cols=len(detail_headers) + 2)
    ws_detail.append_row(detail_headers)
    ws_detail.format("A1:L1", {"textFormat": {"bold": True}})
    print(f"  → {SHEET_DETAIL} 作成（{len(detail_headers)}列）")

    # ④案件サマリー
    ws_cond = sh.add_worksheet(title=SHEET_CONDITIONS + "_new", rows=500, cols=len(CONDITION_HEADERS) + 2)
    ws_cond.append_row(CONDITION_HEADERS)
    ws_cond.format(f"A1:{chr(64 + len(CONDITION_HEADERS))}1", {"textFormat": {"bold": True}})
    print(f"  → {SHEET_CONDITIONS} 作成（{len(CONDITION_HEADERS)}列）")

    # ── 旧シートを全て削除 ──
    print("\n旧シートを削除中...")
    for ws in old_worksheets:
        print(f"  削除: {ws.title}")
        sh.del_worksheet(ws)

    # ── 新シートのタイトルを正式名称にリネーム ──
    print("\nシート名をリネーム中...")
    ws_master.update_title(SHEET_MASTER)
    ws_trades.update_title(SHEET_TRADES)
    ws_detail.update_title(SHEET_DETAIL)
    ws_cond.update_title(SHEET_CONDITIONS)

    print("\n=== 再構築完了 ===")
    print(f"スプレッドシートURL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
    print("\n作成されたシート:")
    for ws in sh.worksheets():
        print(f"  - {ws.title}")


if __name__ == "__main__":
    rebuild()
