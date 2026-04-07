"""
スプレッドシートの全シートを削除して正しいヘッダー構成で再構築するスクリプト。
実行: python3 rebuild_sheets.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gspread
from ata_extract import (
    get_gspread_client, SPREADSHEET_ID,
    SHEET_MASTER, SHEET_TRADES, SHEET_DETAIL,
    MASTER_HEADERS, _set_dropdown_validation
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

    worksheets = sh.worksheets()
    print(f"現在のシート一覧: {[ws.title for ws in worksheets]}")

    # 全シートを削除（シートが1枚しかない場合はGoogleが削除を拒否するため、
    # 一時シートを作成してから削除する）
    print("\n一時シートを作成中...")
    temp_ws = sh.add_worksheet(title="_temp_rebuild", rows=2, cols=2)

    print("既存シートを削除中...")
    for ws in worksheets:
        print(f"  削除: {ws.title}")
        sh.del_worksheet(ws)

    # ①案件マスタ を再作成
    print("\n案件マスタを作成中...")
    ws_master = sh.add_worksheet(title=SHEET_MASTER, rows=500, cols=len(MASTER_HEADERS) + 5)
    ws_master.append_row(MASTER_HEADERS)
    ws_master.format(f"A1:{chr(64 + len(MASTER_HEADERS))}1", {"textFormat": {"bold": True}})
    # 工事種別列（E列=index4）にプルダウン
    _set_dropdown_validation(sh, ws_master.id, col_index=4, values=["新装", "改装"])
    # 入札結果列にプルダウン
    bid_col_idx = MASTER_HEADERS.index("入札結果")
    _set_dropdown_validation(sh, ws_master.id, col_index=bid_col_idx, values=BID_RESULT_OPTIONS)
    print(f"  → 作成完了（{len(MASTER_HEADERS)}列）")

    # ②工種別金額 を再作成
    print("工種別金額を作成中...")
    trades_headers = [
        "案件ID", "案件名", "業態", "坪数", "工種",
        "金額（税抜）", "坪単価",
        "全体合計（税抜）", "全体坪単価",
        "施工年月", "備考", "ドライブURL",
    ]
    ws_trades = sh.add_worksheet(title=SHEET_TRADES, rows=1000, cols=len(trades_headers) + 2)
    ws_trades.append_row(trades_headers)
    ws_trades.format("A1:L1", {"textFormat": {"bold": True}})
    print(f"  → 作成完了（{len(trades_headers)}列）")

    # ③見積明細 を再作成
    print("見積明細を作成中...")
    detail_headers = [
        "案件ID", "案件名", "業者名", "工種", "大分類",
        "項目名", "仕様・規格", "数量", "単位", "単価", "金額", "備考",
    ]
    ws_detail = sh.add_worksheet(title=SHEET_DETAIL, rows=2000, cols=len(detail_headers) + 2)
    ws_detail.append_row(detail_headers)
    ws_detail.format("A1:L1", {"textFormat": {"bold": True}})
    print(f"  → 作成完了（{len(detail_headers)}列）")

    # ④案件サマリー を再作成
    print("案件サマリーを作成中...")
    ws_cond = sh.add_worksheet(title=SHEET_CONDITIONS, rows=500, cols=len(CONDITION_HEADERS) + 2)
    ws_cond.append_row(CONDITION_HEADERS)
    ws_cond.format(f"A1:{chr(64 + len(CONDITION_HEADERS))}1", {"textFormat": {"bold": True}})
    print(f"  → 作成完了（{len(CONDITION_HEADERS)}列）")

    # 一時シートを削除
    print("\n一時シートを削除中...")
    sh.del_worksheet(temp_ws)

    print("\n=== 再構築完了 ===")
    print(f"スプレッドシートURL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
    print("\n作成されたシート:")
    for ws in sh.worksheets():
        print(f"  - {ws.title}")

if __name__ == "__main__":
    rebuild()
