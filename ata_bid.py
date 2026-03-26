"""
Apples to Apples 入札結果分析モジュール
- 入札結果データの読み込み・フィルタリング
- 自社受注/他社受注の坪単価集計
- 散布図用データ生成
"""

import os
import sys

import gspread

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import get_google_credentials

SPREADSHEET_ID = "10uXWjPTuYcMtnvmWt6A9fMWRxvPlVf82vIVpM90u95U"

SHEET_MASTER = "案件マスタ"


def get_spreadsheet() -> gspread.Spreadsheet:
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = get_google_credentials(scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


def load_bid_data(master: list[dict]) -> list[dict]:
    """案件マスタから入札結果が入力されている案件を抽出し分析用データを構築する"""
    rows = []
    for p in master:
        bid_result = str(p.get("入札結果", "")).strip()
        if not bid_result or bid_result == "未定":
            continue

        tsubo = 0.0
        try:
            tsubo = float(p.get("坪数", 0))
        except (ValueError, TypeError):
            pass
        if tsubo <= 0:
            continue

        total = 0
        try:
            total = int(float(p.get("合計金額（税抜）", 0)))
        except (ValueError, TypeError):
            pass

        own_unit_price = round(total / tsubo) if tsubo > 0 and total > 0 else 0

        competitor_amount = 0
        try:
            competitor_amount = int(float(p.get("他社受注金額", 0) or 0))
        except (ValueError, TypeError):
            pass

        competitor_unit_price = round(competitor_amount / tsubo) if tsubo > 0 and competitor_amount > 0 else 0

        rows.append({
            "案件ID": p.get("案件ID", ""),
            "案件名": p.get("案件名", ""),
            "ブランド名": p.get("ブランド名", ""),
            "業態": p.get("業態", ""),
            "工事種別": p.get("工事種別", ""),
            "坪数": tsubo,
            "自社提出金額": total,
            "自社提出坪単価": own_unit_price,
            "入札結果": bid_result,
            "他社受注金額": competitor_amount,
            "他社受注坪単価": competitor_unit_price,
            "差額": own_unit_price - competitor_unit_price if competitor_unit_price > 0 else 0,
        })

    return rows


def filter_bid_data(
    data: list[dict],
    brand: str = "",
    category: str = "",
    koji_type: str = "",
) -> list[dict]:
    """フィルター条件に基づいて入札データを絞り込む"""
    filtered = data
    if brand and brand.strip():
        bl = brand.strip().lower()
        filtered = [r for r in filtered
                    if bl in r.get("ブランド名", "").lower()
                    or bl in r.get("案件名", "").lower()]
    if category:
        filtered = [r for r in filtered if r.get("業態", "") == category]
    if koji_type and koji_type != "すべて":
        if koji_type == "新装のみ":
            filtered = [r for r in filtered if r.get("工事種別", "") == "新装"]
        elif koji_type == "改装のみ":
            filtered = [r for r in filtered if r.get("工事種別", "") == "改装"]
    return filtered


def calc_bid_summary(data: list[dict]) -> dict:
    """入札データの集計サマリーを算出する"""
    own_won = [r for r in data if r["入札結果"] == "自社受注" and r["自社提出坪単価"] > 0]
    competitor_won = [r for r in data if r["入札結果"] == "他社受注" and r["他社受注坪単価"] > 0]

    own_avg = round(sum(r["自社提出坪単価"] for r in own_won) / len(own_won)) if own_won else 0
    competitor_avg = round(sum(r["他社受注坪単価"] for r in competitor_won) / len(competitor_won)) if competitor_won else 0

    # 他社受注案件における自社提出坪単価の平均
    own_in_lost = [r for r in data if r["入札結果"] == "他社受注" and r["自社提出坪単価"] > 0]
    own_avg_in_lost = round(sum(r["自社提出坪単価"] for r in own_in_lost) / len(own_in_lost)) if own_in_lost else 0

    diff = own_avg_in_lost - competitor_avg if own_avg_in_lost > 0 and competitor_avg > 0 else 0

    return {
        "自社受注件数": len(own_won),
        "他社受注件数": len(competitor_won),
        "自社受注平均坪単価": own_avg,
        "他社受注平均坪単価": competitor_avg,
        "敗退時自社平均坪単価": own_avg_in_lost,
        "差額": diff,
    }


def write_bid_result(project_id: str, bid_result: str, competitor_amount: int = 0):
    """案件マスタの入札結果列を更新する"""
    sh = get_spreadsheet()
    ws = sh.worksheet(SHEET_MASTER)
    _data = ws.get_all_values()
    _headers = _data[0] if _data else []
    all_records = [dict(zip(_headers, row)) for row in _data[1:]] if len(_data) > 1 else []

    # ヘッダーから列インデックスを取得
    headers = ws.row_values(1)
    bid_col = None
    comp_amount_col = None
    comp_unit_col = None
    for i, h in enumerate(headers):
        if h == "入札結果":
            bid_col = i + 1
        elif h == "他社受注金額":
            comp_amount_col = i + 1
        elif h == "他社受注坪単価":
            comp_unit_col = i + 1

    if bid_col is None:
        return  # 列が見つからない

    # 案件IDで行を特定
    for row_idx, record in enumerate(all_records):
        if record.get("案件ID") == project_id:
            data_row = row_idx + 2  # ヘッダー行 + 0-indexed
            ws.update_cell(data_row, bid_col, bid_result)
            if comp_amount_col and competitor_amount > 0:
                ws.update_cell(data_row, comp_amount_col, competitor_amount)
                # 坪単価を自動計算
                tsubo = float(record.get("坪数", 0) or 0)
                if comp_unit_col and tsubo > 0:
                    ws.update_cell(data_row, comp_unit_col, round(competitor_amount / tsubo))
            break
