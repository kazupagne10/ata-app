"""
スプレッドシートの全シートを削除して正しいヘッダー構成で再構築するスクリプト。

戦略:
  どんな中途半端な状態からでも正しく再構築できるよう、
  「必要なシートを確実に作り、不要なシートを全て削除する」方式を採用。

  1. 正式名称のシートが既に存在する場合はスキップ（既に正しい状態）
  2. _new / _temp_rebuild などの残骸シートを削除
  3. 不足しているシートを新規作成
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

TRADES_HEADERS = [
    "案件ID", "案件名", "業態", "坪数", "工種",
    "金額（税抜）", "坪単価",
    "全体合計（税抜）", "全体坪単価",
    "施工年月", "備考", "ドライブURL",
]

DETAIL_HEADERS = [
    "案件ID", "案件名", "業者名", "工種", "大分類",
    "項目名", "仕様・規格", "数量", "単位", "単価", "金額", "備考",
]

CONDITION_HEADERS = [
    "案件ID", "案件名",
    "工事時間帯", "立地タイプ", "工事状態", "指定業者範囲",
    "厨房区画", "室外機設置階数", "照明器具支給", "家具・什器含む",
    "業態", "フロア", "グリストラップ", "排気ダクト",
    "工期", "施工エリア", "防災工事", "その他備考",
]

# 正式なシート名と設定のマッピング
SHEET_SPECS = [
    {
        "title": SHEET_MASTER,
        "rows": 500,
        "headers": MASTER_HEADERS,
        "extra_setup": "master",
    },
    {
        "title": SHEET_TRADES,
        "rows": 1000,
        "headers": TRADES_HEADERS,
        "extra_setup": None,
    },
    {
        "title": SHEET_DETAIL,
        "rows": 2000,
        "headers": DETAIL_HEADERS,
        "extra_setup": None,
    },
    {
        "title": SHEET_CONDITIONS,
        "rows": 500,
        "headers": CONDITION_HEADERS,
        "extra_setup": None,
    },
]

OFFICIAL_TITLES = {spec["title"] for spec in SHEET_SPECS}


def rebuild():
    print("Googleスプレッドシートに接続中...")
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    all_ws = sh.worksheets()
    existing_titles = {ws.title: ws for ws in all_ws}
    print(f"現在のシート一覧: {list(existing_titles.keys())}")

    # ── Step1: 正式名称のシートを全てクリア（データ削除）またはスキップ ──
    # 正式名称シートが存在する場合は一旦全データを消去してヘッダーだけ再設定する
    # 存在しない場合は新規作成する
    created_ws = {}

    for spec in SHEET_SPECS:
        title = spec["title"]
        headers = spec["headers"]
        cols = len(headers) + 5

        if title in existing_titles:
            # 既存シートをクリアしてヘッダーを再設定
            ws = existing_titles[title]
            print(f"  既存シートをクリア: {title}")
            ws.clear()
            ws.append_row(headers)
            ws.format(f"A1:{chr(64 + len(headers))}1", {"textFormat": {"bold": True}})
        else:
            # 新規作成
            print(f"  新規作成: {title}")
            ws = sh.add_worksheet(title=title, rows=spec["rows"], cols=cols)
            ws.append_row(headers)
            ws.format(f"A1:{chr(64 + len(headers))}1", {"textFormat": {"bold": True}})

        # 案件マスタ固有の設定
        if spec["extra_setup"] == "master":
            try:
                _set_dropdown_validation(sh, ws.id, col_index=4, values=["新装", "改装"])
                bid_col_idx = headers.index("入札結果")
                _set_dropdown_validation(sh, ws.id, col_index=bid_col_idx, values=BID_RESULT_OPTIONS)
            except Exception as e:
                print(f"    プルダウン設定スキップ: {e}")

        created_ws[title] = ws
        print(f"  → {title} 完了（{len(headers)}列）")

    # ── Step2: 不要なシート（正式名称以外）を全て削除 ──
    print("\n不要なシートを削除中...")
    # 最新のシート一覧を再取得
    all_ws_now = sh.worksheets()
    for ws in all_ws_now:
        if ws.title not in OFFICIAL_TITLES:
            print(f"  削除: {ws.title}")
            try:
                sh.del_worksheet(ws)
            except Exception as e:
                print(f"    削除失敗（スキップ）: {e}")

    print("\n=== 再構築完了 ===")
    print(f"スプレッドシートURL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
    print("\n最終シート一覧:")
    for ws in sh.worksheets():
        print(f"  - {ws.title}")


if __name__ == "__main__":
    rebuild()
