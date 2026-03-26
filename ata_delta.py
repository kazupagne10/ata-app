"""
Apples to Apples 増減表生成スクリプト
- 新規案件と参照案件を工種別に比較
- ベース金額（参照坪単価×新規坪数）vs 新規実額の増減を算出
- Googleスプレッドシートに「増減表」シートを自動生成
"""

import argparse
import os
import sys

import gspread

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import get_google_credentials

# ── 設定 ──
SPREADSHEET_ID = "10uXWjPTuYcMtnvmWt6A9fMWRxvPlVf82vIVpM90u95U"

SHEET_MASTER = "案件マスタ"
SHEET_TRADES = "工種別金額"
SHEET_DELTA = "増減表"

TRADE_CATEGORIES = ["内装", "電気", "給排水衛生", "空調換気", "ガス", "看板"]


# ── スプレッドシート接続 ──

def get_spreadsheet(readonly: bool = False) -> gspread.Spreadsheet:
    scope = "https://www.googleapis.com/auth/spreadsheets"
    if readonly:
        scope += ".readonly"
    creds = get_google_credentials([scope])
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


# ── データ取得 ──

def _safe_get_all_records(ws: gspread.Worksheet) -> list[dict]:
    """空白・重複ヘッダーがあっても安全にレコードを取得する（全列空の行はスキップ）"""
    data = ws.get_all_values()
    if not data:
        return []
    headers = data[0]
    records = []
    for row in data[1:]:
        if not any(cell.strip() for cell in row):
            continue
        records.append(dict(zip(headers, row)))
    return records


def load_master(sh: gspread.Spreadsheet) -> list[dict]:
    ws = sh.worksheet(SHEET_MASTER)
    return _safe_get_all_records(ws)


def load_trades(sh: gspread.Spreadsheet) -> list[dict]:
    ws = sh.worksheet(SHEET_TRADES)
    return _safe_get_all_records(ws)


def find_project(master: list[dict], project_id: str) -> dict | None:
    for p in master:
        if p.get("案件ID") == project_id:
            return p
    return None


def get_trade_amounts(project: dict, trades: list[dict]) -> dict[str, int]:
    """案件の工種別金額を取得（工種別金額シート優先、案件マスタでフォールバック）"""
    pid = project.get("案件ID")
    amounts = {}
    for cat in TRADE_CATEGORIES:
        matched = [t for t in trades
                   if t.get("案件ID") == pid and t.get("工種") == cat]
        if matched:
            amounts[cat] = int(float(matched[0].get("金額（税抜）", 0)))
        else:
            amounts[cat] = int(float(project.get(cat, 0) or 0))
    return amounts


# ── 増減計算 ──

def calc_delta(ref_project: dict, new_project: dict, trades: list[dict]) -> list[dict]:
    """工種別の増減データを計算する"""
    ref_tsubo = float(ref_project.get("坪数", 0))
    new_tsubo = float(new_project.get("坪数", 0))

    if ref_tsubo == 0 or new_tsubo == 0:
        print("エラー: 坪数が0の案件では比較できません。", file=sys.stderr)
        sys.exit(1)

    ref_amounts = get_trade_amounts(ref_project, trades)
    new_amounts = get_trade_amounts(new_project, trades)

    rows = []
    total_ref = 0
    total_base = 0
    total_new = 0

    for cat in TRADE_CATEGORIES:
        ref_amt = ref_amounts[cat]
        new_amt = new_amounts[cat]
        unit_price = ref_amt / ref_tsubo if ref_tsubo else 0
        base_amt = round(unit_price * new_tsubo)
        delta = new_amt - base_amt
        rate = (delta / base_amt * 100) if base_amt != 0 else (100.0 if new_amt > 0 else 0.0)

        total_ref += ref_amt
        total_base += base_amt
        total_new += new_amt

        rows.append({
            "trade": cat,
            "ref_amount": ref_amt,
            "ref_unit_price": round(unit_price),
            "base_amount": base_amt,
            "new_amount": new_amt,
            "new_unit_price": round(new_amt / new_tsubo) if new_tsubo else 0,
            "delta": delta,
            "rate": round(rate, 1),
        })

    # 合計行
    total_delta = total_new - total_base
    total_rate = (total_delta / total_base * 100) if total_base != 0 else 0.0
    rows.append({
        "trade": "合計",
        "ref_amount": total_ref,
        "ref_unit_price": round(total_ref / ref_tsubo) if ref_tsubo else 0,
        "base_amount": total_base,
        "new_amount": total_new,
        "new_unit_price": round(total_new / new_tsubo) if new_tsubo else 0,
        "delta": total_delta,
        "rate": round(total_rate, 1),
    })

    return rows


# ── 表示 ──

def display_delta(ref_project: dict, new_project: dict, delta_rows: list[dict]):
    ref_id = ref_project.get("案件ID", "")
    ref_name = ref_project.get("案件名", "")
    ref_tsubo = ref_project.get("坪数", 0)
    new_id = new_project.get("案件ID", "")
    new_name = new_project.get("案件名", "")
    new_tsubo = new_project.get("坪数", 0)

    print("=" * 90)
    print("  Apples to Apples 増減表")
    print("=" * 90)
    print()
    print(f"  新規案件:   [{new_id}] {new_name}（{new_tsubo} 坪）")
    print(f"  参照案件:   [{ref_id}] {ref_name}（{ref_tsubo} 坪）")
    print()
    print("-" * 90)
    print(f"  {'工種':<10}  {'参照坪単価':>12}  {'ベース金額':>14}  {'新規金額':>14}  {'増減額':>14}  {'増減率':>8}")
    print(f"  {'─'*8}    {'─'*12}  {'─'*14}  {'─'*14}  {'─'*14}  {'─'*8}")

    for r in delta_rows:
        trade = r["trade"]
        if trade == "合計":
            print(f"  {'─'*8}    {'─'*12}  {'─'*14}  {'─'*14}  {'─'*14}  {'─'*8}")

        sign = "+" if r["delta"] > 0 else ""
        rate_sign = "+" if r["rate"] > 0 else ""
        print(f"  {trade:<8}    ¥{r['ref_unit_price']:>10,}/坪"
              f"  ¥{r['base_amount']:>12,}"
              f"  ¥{r['new_amount']:>12,}"
              f"  {sign}¥{r['delta']:>12,}"
              f"  {rate_sign}{r['rate']:>6.1f}%")

    print()
    print("=" * 90)


# ── スプレッドシート書き込み ──

def write_delta_sheet(sh: gspread.Spreadsheet, ref_project: dict, new_project: dict, delta_rows: list[dict]):
    """増減表シートに結果を書き込む"""
    ref_id = ref_project.get("案件ID", "")
    ref_name = ref_project.get("案件名", "")
    ref_tsubo = float(ref_project.get("坪数", 0))
    new_id = new_project.get("案件ID", "")
    new_name = new_project.get("案件名", "")
    new_tsubo = float(new_project.get("坪数", 0))

    existing = {ws.title for ws in sh.worksheets()}

    if SHEET_DELTA not in existing:
        ws = sh.add_worksheet(title=SHEET_DELTA, rows=200, cols=14)
        ws.append_row([
            "新規案件ID", "新規案件名", "新規坪数",
            "参照案件ID", "参照案件名", "参照坪数",
            "工種", "参照坪単価", "ベース金額", "新規金額", "新規坪単価",
            "増減額", "増減率（%）", "判定",
        ])
        ws.format("A1:N1", {"textFormat": {"bold": True}})
        print(f"  シート「{SHEET_DELTA}」を作成しました")
    else:
        ws = sh.worksheet(SHEET_DELTA)

    # データ行を書き込み
    rows_to_write = []
    for r in delta_rows:
        # 判定: 増減率に基づく3段階評価
        rate = r["rate"]
        if abs(rate) <= 5:
            judgement = "◎ 妥当"
        elif abs(rate) <= 15:
            judgement = "○ 軽微" if rate > 0 else "○ 減額"
        elif abs(rate) <= 30:
            judgement = "△ 要確認" if rate > 0 else "△ 大幅減"
        else:
            judgement = "× 要精査" if rate > 0 else "× 大幅減"

        rows_to_write.append([
            new_id, new_name, new_tsubo,
            ref_id, ref_name, ref_tsubo,
            r["trade"],
            r["ref_unit_price"],
            r["base_amount"],
            r["new_amount"],
            r["new_unit_price"],
            r["delta"],
            r["rate"],
            judgement,
        ])

    ws.append_rows(rows_to_write, value_input_option="USER_ENTERED")
    print(f"  → 増減表に {len(rows_to_write)} 行書き込み完了")

    # 条件付き書式: 増減額列の色分け（新規データ範囲に適用）
    _apply_conditional_formatting(ws)

    print()
    print(f"  スプレッドシートURL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")


def _apply_conditional_formatting(ws: gspread.Worksheet):
    """増減額列に条件付き書式（赤=増額、青=減額）をSheets API直接で適用する"""
    sheet_id = ws.id

    # L列 = index 11
    col_index = 11
    range_def = {
        "sheetId": sheet_id,
        "startRowIndex": 1,
        "startColumnIndex": col_index,
        "endColumnIndex": col_index + 1,
    }

    requests = [
        # 増額 → 薄赤背景・赤文字
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [range_def],
                    "booleanRule": {
                        "condition": {
                            "type": "NUMBER_GREATER",
                            "values": [{"userEnteredValue": "0"}],
                        },
                        "format": {
                            "backgroundColor": {"red": 1, "green": 0.85, "blue": 0.85},
                            "textFormat": {"foregroundColor": {"red": 0.8, "green": 0, "blue": 0}},
                        },
                    },
                },
                "index": 0,
            }
        },
        # 減額 → 薄青背景・青文字
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [range_def],
                    "booleanRule": {
                        "condition": {
                            "type": "NUMBER_LESS",
                            "values": [{"userEnteredValue": "0"}],
                        },
                        "format": {
                            "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 1},
                            "textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0.8}},
                        },
                    },
                },
                "index": 1,
            }
        },
    ]

    try:
        ws.spreadsheet.batch_update({"requests": requests})
        print("  → 条件付き書式を適用しました（増額=赤、減額=青）")
    except Exception as e:
        print(f"  ⚠ 条件付き書式の適用をスキップしました: {e}")


# ── メイン ──

def main():
    parser = argparse.ArgumentParser(description="Apples to Apples 増減表生成")
    parser.add_argument("--new", "-n", required=True, help="新規案件ID（例: P-002）")
    parser.add_argument("--ref", "-r", required=True, help="参照案件ID（例: P-001）")
    parser.add_argument("--no-write", action="store_true", help="スプレッドシートに書き込まない（表示のみ）")
    args = parser.parse_args()

    print("スプレッドシートからデータを取得中...")
    sh = get_spreadsheet(readonly=args.no_write)
    master = load_master(sh)
    trades = load_trades(sh)

    if not master:
        print("エラー: 案件マスタにデータがありません。", file=sys.stderr)
        sys.exit(1)

    print(f"  → {len(master)} 件の過去案件を取得しました")
    print()

    # 案件の取得
    new_project = find_project(master, args.new)
    ref_project = find_project(master, args.ref)

    if not new_project:
        print(f"エラー: 新規案件 '{args.new}' が見つかりません。", file=sys.stderr)
        sys.exit(1)
    if not ref_project:
        print(f"エラー: 参照案件 '{args.ref}' が見つかりません。", file=sys.stderr)
        sys.exit(1)

    # 増減計算
    delta_rows = calc_delta(ref_project, new_project, trades)

    # 表示
    display_delta(ref_project, new_project, delta_rows)

    # スプレッドシート書き込み
    if not args.no_write:
        print()
        print("Googleスプレッドシートに書き込み中...")
        sh_write = get_spreadsheet(readonly=False)
        write_delta_sheet(sh_write, ref_project, new_project, delta_rows)


if __name__ == "__main__":
    main()
