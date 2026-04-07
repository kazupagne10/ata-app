"""
Apples to Apples 複数社見積 横並び比較スクリプト
- 同一案件に対する複数社の見積書を工種別に横並び比較
- 各工種で最安値の業者をハイライト
- 比較結果をスプレッドシート「比較表」シートに書き込み
"""

import argparse
import json
import os
import sys
import tempfile

import gspread
import pdfplumber
from openai import OpenAI

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import get_google_credentials, get_openai_client

# ── 設定 ──
SPREADSHEET_ID = "10uXWjPTuYcMtnvmWt6A9fMWRxvPlVf82vIVpM90u95U"

SHEET_COMPARE = "比較表"
TRADE_CATEGORIES = ["内装", "電気", "給排水衛生", "空調換気", "ガス", "看板"]


# ── スプレッドシート接続 ──

def get_spreadsheet(readonly: bool = False) -> gspread.Spreadsheet:
    scope = "https://www.googleapis.com/auth/spreadsheets"
    if readonly:
        scope += ".readonly"
    creds = get_google_credentials([scope])
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


# ── PDF → GPT-4o による見積データ抽出（1社分） ──

def extract_text_from_pdf(pdf_path: str, max_pages: int = 5) -> str:
    texts = []
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[:max_pages]
        for i, page in enumerate(pages):
            text = page.extract_text()
            if text:
                texts.append(f"--- ページ {i + 1} ---\n{text}")
    return "\n\n".join(texts)


def extract_estimate_with_gpt(client: OpenAI, text: str, filename: str) -> dict:
    """見積書PDFから業者名・工種別金額を抽出"""
    trade_list = "、".join(TRADE_CATEGORIES)

    prompt = f"""以下の見積書PDFのテキストを分析し、JSONで返してください。

## 抽出項目
1. **company**: 見積を出した会社名（施工業者名）
2. **total_amount**: 合計金額（税別、数値のみ）
3. **trades**: 工種別金額の内訳
4. **trade_mapping**: 各項目を次の6カテゴリに分類: {trade_list}

## ファイル名（参考情報）
{filename}

## 出力JSON形式（厳密にこの形式で）
```json
{{
  "company": "会社名",
  "total_amount": 数値,
  "trades": [
    {{
      "trade": "元の工種名",
      "mapped_category": "6カテゴリのいずれか（該当なしならnull）",
      "amount": 数値
    }}
  ]
}}
```

## 見積書テキスト
{text}

JSONのみを返してください。"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
    )

    content = response.choices[0].message.content.strip()
    if content.startswith("```"):
        content = content.split("\n", 1)[1]
        content = content.rsplit("```", 1)[0]

    return json.loads(content)


# ── 比較データ構築 ──

def build_comparison(estimates: list[dict]) -> dict:
    """複数社のデータから比較テーブルを構築する"""
    companies = []
    trade_data = {cat: {} for cat in TRADE_CATEGORIES}
    totals = {}

    for est in estimates:
        company = est["company"]
        companies.append(company)
        totals[company] = est["total_amount"]

        # 工種別集計
        cat_amounts = {cat: 0 for cat in TRADE_CATEGORIES}
        for t in est["trades"]:
            mapped = t.get("mapped_category")
            if mapped and mapped in cat_amounts:
                cat_amounts[mapped] += t["amount"]

        for cat in TRADE_CATEGORIES:
            trade_data[cat][company] = cat_amounts[cat]

    # 各工種の最安値を特定
    cheapest = {}
    for cat in TRADE_CATEGORIES:
        amounts = {c: trade_data[cat][c] for c in companies if trade_data[cat][c] > 0}
        if amounts:
            cheapest[cat] = min(amounts, key=amounts.get)
        else:
            cheapest[cat] = None

    # 合計の最安値
    total_amounts = {c: totals[c] for c in companies if totals[c] > 0}
    cheapest["合計"] = min(total_amounts, key=total_amounts.get) if total_amounts else None

    return {
        "companies": companies,
        "trade_data": trade_data,
        "totals": totals,
        "cheapest": cheapest,
    }


# ── スプレッドシート書き込み ──

def write_comparison_sheet(project_name: str, tsubo: float, comparison: dict):
    """比較表シートに結果を書き込む"""
    gc_scope = "https://www.googleapis.com/auth/spreadsheets"
    creds = get_google_credentials([gc_scope])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    existing = {ws.title for ws in sh.worksheets()}

    if SHEET_COMPARE not in existing:
        ws = sh.add_worksheet(title=SHEET_COMPARE, rows=200, cols=20)
    else:
        ws = sh.worksheet(SHEET_COMPARE)

    companies = comparison["companies"]
    trade_data = comparison["trade_data"]
    totals = comparison["totals"]
    cheapest = comparison["cheapest"]

    # 書き込み開始行を取得（既存データの下に追加）
    existing_data = ws.get_all_values()
    start_row = len(existing_data) + 1
    if start_row > 1:
        start_row += 1  # 空行を1行挟む

    rows = []

    # ヘッダー: 案件名
    rows.append([f"案件名: {project_name}", f"坪数: {tsubo}", "", *["" for _ in companies]])

    # カラムヘッダー
    header = ["工種", *companies, "最安値業者", "最安値金額"]
    rows.append(header)

    # 工種別行
    for cat in TRADE_CATEGORIES:
        row = [cat]
        for c in companies:
            row.append(trade_data[cat].get(c, 0))
        cheapest_company = cheapest.get(cat, "")
        cheapest_amount = trade_data[cat].get(cheapest_company, 0) if cheapest_company else ""
        row.append(cheapest_company or "")
        row.append(cheapest_amount)
        rows.append(row)

    # 合計行
    total_row = ["合計"]
    for c in companies:
        total_row.append(totals.get(c, 0))
    cheapest_total_co = cheapest.get("合計", "")
    cheapest_total_amt = totals.get(cheapest_total_co, 0) if cheapest_total_co else ""
    total_row.append(cheapest_total_co or "")
    total_row.append(cheapest_total_amt)
    rows.append(total_row)

    # 坪単価行
    unit_row = ["坪単価"]
    for c in companies:
        amt = totals.get(c, 0)
        unit_row.append(round(amt / tsubo) if tsubo else 0)
    unit_row.append("")
    unit_row.append("")
    rows.append(unit_row)

    # 書き込み
    cell_range = f"A{start_row}"
    ws.update(cell_range, rows, value_input_option="USER_ENTERED")

    # ヘッダー行の書式
    header_row_num = start_row + 1
    num_cols = len(header)
    col_letter = chr(ord("A") + num_cols - 1)
    ws.format(f"A{header_row_num}:{col_letter}{header_row_num}", {"textFormat": {"bold": True}})

    # 合計行の書式
    total_row_num = header_row_num + len(TRADE_CATEGORIES) + 1
    ws.format(f"A{total_row_num}:{col_letter}{total_row_num}", {"textFormat": {"bold": True}})

    # 最安値セルをハイライト（黄色背景）
    highlight_format = {
        "backgroundColor": {"red": 1, "green": 0.95, "blue": 0.6},
        "textFormat": {"bold": True},
    }

    for i, cat in enumerate(TRADE_CATEGORIES):
        cheapest_company = cheapest.get(cat)
        if cheapest_company and cheapest_company in companies:
            col_idx = companies.index(cheapest_company) + 1  # +1 for 工種 column
            col_letter = chr(ord("A") + col_idx)
            row_num = header_row_num + 1 + i
            ws.format(f"{col_letter}{row_num}", highlight_format)

    # 合計の最安値もハイライト
    cheapest_total_co = cheapest.get("合計")
    if cheapest_total_co and cheapest_total_co in companies:
        col_idx = companies.index(cheapest_total_co) + 1
        col_letter = chr(ord("A") + col_idx)
        ws.format(f"{col_letter}{total_row_num}", highlight_format)

    print(f"  → 比較表シートに書き込み完了（{len(companies)} 社比較）")
    print(f"  スプレッドシートURL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")


# ── CLI ──

def main():
    parser = argparse.ArgumentParser(description="AtA 複数社見積 横並び比較")
    parser.add_argument(
        "pdfs", nargs="+", help="見積書PDFファイルのパス（複数指定）",
    )
    parser.add_argument("--project", "-p", required=True, help="案件名")
    parser.add_argument("--tsubo", "-t", type=float, required=True, help="坪数")
    parser.add_argument("--no-write", action="store_true", help="スプレッドシートに書き込まない")
    args = parser.parse_args()

    openai_client = get_openai_client()
    estimates = []

    for pdf_path in args.pdfs:
        pdf_path = os.path.abspath(pdf_path)
        if not os.path.exists(pdf_path):
            print(f"エラー: {pdf_path} が見つかりません。", file=sys.stderr)
            continue

        filename = os.path.basename(pdf_path)
        print(f"  解析中: {filename}")
        text = extract_text_from_pdf(pdf_path, max_pages=5)
        est = extract_estimate_with_gpt(openai_client, text, filename)
        estimates.append(est)
        print(f"    → {est['company']}: {est['total_amount']:,.0f}")

    if len(estimates) < 2:
        print("エラー: 比較には最低2社の見積書が必要です。", file=sys.stderr)
        sys.exit(1)

    # 比較テーブル構築
    comparison = build_comparison(estimates)

    # 表示
    display_comparison(args.project, args.tsubo, comparison)

    # 書き込み
    if not args.no_write:
        print("\nスプレッドシートに書き込み中...")
        write_comparison_sheet(args.project, args.tsubo, comparison)


def display_comparison(project_name: str, tsubo: float, comparison: dict):
    """比較結果をコンソールに表示"""
    companies = comparison["companies"]
    trade_data = comparison["trade_data"]
    totals = comparison["totals"]
    cheapest = comparison["cheapest"]

    print()
    print("=" * 90)
    print(f"  複数社見積比較: {project_name}（{tsubo} 坪）")
    print("=" * 90)

    # ヘッダー
    header = f"  {'工種':<12}"
    for c in companies:
        header += f"  {c:>14}"
    header += f"  {'最安値':>14}"
    print(header)
    print("  " + "─" * (14 + 16 * len(companies) + 16))

    for cat in TRADE_CATEGORIES:
        row = f"  {cat:<10}"
        for c in companies:
            amt = trade_data[cat].get(c, 0)
            marker = " ★" if cheapest.get(cat) == c and amt > 0 else "  "
            row += f"  ¥{amt:>12,}{marker}"
        ch = cheapest.get(cat, "")
        row += f"  {ch or '-':>14}"
        print(row)

    print("  " + "─" * (14 + 16 * len(companies) + 16))

    row = f"  {'合計':<10}"
    for c in companies:
        amt = totals.get(c, 0)
        marker = " ★" if cheapest.get("合計") == c else "  "
        row += f"  ¥{amt:>12,}{marker}"
    row += f"  {cheapest.get('合計', '-'):>14}"
    print(row)

    print()
    print("=" * 90)


if __name__ == "__main__":
    main()
