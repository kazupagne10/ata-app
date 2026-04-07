"""
Apples to Apples PDF データ抽出スクリプト
- 図面PDFから「坪数」を抽出
- 見積書PDFから「合計金額」「工種別金額」を抽出
- GPT-4oを使用してデータを構造化
- Googleスプレッドシートに自動書き込み
"""

import glob
import json
import os
import sys
from datetime import date

import gspread
import pdfplumber
from openai import OpenAI

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import get_google_credentials, get_openai_client

# ── 設定 ──
SPREADSHEET_ID = "10uXWjPTuYcMtnvmWt6A9fMWRxvPlVf82vIVpM90u95U"

SHEET_MASTER = "案件マスタ"
SHEET_TRADES = "工種別金額"
SHEET_DETAIL = "見積明細"

# スプレッドシートの6工種カテゴリ
TRADE_CATEGORIES = ["内装", "電気", "給排水衛生", "空調換気", "ガス", "看板"]

# 工事条件の定義: (key, 日本語名, 選択肢リスト or None=テキスト)
CONDITION_FIELDS = [
    ("work_hours",       "工事時間帯",       ["日中", "夜間", "混在"]),
    ("location_type",    "立地タイプ",       ["路面店", "ロードサイド", "商業施設", "その他"]),
    ("work_state",       "工事状態",         ["スケルトン", "居抜き"]),
    ("designated_trades","指定業者範囲",     ["なし", "電気のみ", "空調のみ", "給排水のみ", "電気＋空調", "複数工種", "その他"]),
    ("kitchen_zone",     "厨房区画",         ["新規", "既存", "なし"]),
    ("outdoor_unit_floor","室外機設置階数",  "OUTDOOR_UNIT_FLOOR_OPTIONS"),  # config参照
    ("lighting_supply",  "照明器具支給",     ["あり", "なし"]),
    ("furniture_included","家具・什器含む",  ["あり", "なし"]),
    ("biz_category",     "業態",             None),  # config.GYOTAI_OPTIONS を動的参照
    ("floor",            "フロア",           ["地下", "1F", "2F以上"]),
    ("grease_trap",      "グリストラップ",   ["新設", "既存流用", "なし"]),
    ("exhaust_duct",     "排気ダクト",       ["新設", "既存流用", "なし"]),
    ("construction_days","工期",             "CONSTRUCTION_DAYS_OPTIONS"),   # config参照
    ("construction_area","施工エリア",       "CONSTRUCTION_AREA_OPTIONS"),   # config参照
    ("fire_prevention",  "防災工事",         ["あり", "なし"]),
    ("remarks",          "その他備考",       None),
]

CONDITION_KEYS = [f[0] for f in CONDITION_FIELDS]


# ── PDF検索・テキスト抽出 ──

def find_pdfs(input_dir: str) -> tuple[str | None, str | None]:
    """フォルダ内のPDFを図面と見積書に分類する"""
    pdf_files = glob.glob(os.path.join(input_dir, "*.pdf"))
    drawing_pdf = None
    estimate_pdf = None

    for path in pdf_files:
        name = os.path.basename(path)
        if "見積" in name:
            estimate_pdf = path
        elif "図" in name or "入札" in name:
            drawing_pdf = path

    return drawing_pdf, estimate_pdf


def extract_text_from_pdf(pdf_path: str, max_pages: int | None = None) -> str:
    """PDFからテキストを抽出する"""
    texts = []
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[:max_pages] if max_pages else pdf.pages
        for i, page in enumerate(pages):
            text = page.extract_text()
            if text:
                texts.append(f"--- ページ {i + 1} ---\n{text}")
    return "\n\n".join(texts)


# ── GPT-4o による抽出 ──

def extract_with_gpt4o(client: OpenAI, drawing_text: str, estimate_text: str) -> dict:
    """GPT-4oを使って坪数・金額データを構造化抽出する"""
    trade_list = "、".join(TRADE_CATEGORIES)

    # 工事条件の選択肢一覧を生成
    from config import GYOTAI_OPTIONS
    conditions_spec = ""
    for key, label, choices in CONDITION_FIELDS:
        # biz_category は config.GYOTAI_OPTIONS を使用
        if key == "biz_category":
            conditions_spec += f'    - {key} ({label}): 選択肢=[{", ".join(GYOTAI_OPTIONS)}]\n'
        elif choices:
            conditions_spec += f'    - {key} ({label}): 選択肢=[{", ".join(choices)}]\n'
        else:
            conditions_spec += f'    - {key} ({label}): テキスト\n'

    prompt = f"""以下の2つのPDFから抽出したテキストを分析し、データをJSON形式で返してください。

## 抽出してほしい項目

1. **project_name**: 工事名称（例: "CRISP SALAD WORKS マルイ海老名 新装工事"）
2. **brand**: ブランド名（例: "CRISP SALAD WORKS"）
3. **category**: 業態。以下の6つから選択: {", ".join(GYOTAI_OPTIONS)}
4. **tsubo**: 図面PDFから計画面積の坪数（数値のみ。例: 21.7）
5. **prefecture**: 都道府県（例: "神奈川県"）
6. **station**: 最寄り駅（例: "海老名"）
7. **total_amount**: 見積書表紙の合計金額（税別、数値のみ。例: 15200000）
8. **trades**: 見積書表紙に記載されている各業者/工種ごとの金額リスト
9. **trade_mapping**: 各業者の工種を以下の6カテゴリのいずれかに分類してください: {trade_list}
   - 「内装・給排水工事」のように複数工種が混在する場合、金額は主たる工種（この例なら「内装」）に分類してください
   - 「ディレクション費」「諸経費」など6カテゴリに当てはまらないものは分類不要です
   - 「サイン工事」は「看板」に分類してください
10. **construction_item_flags**: 以下の工事項目がこの案件に含まれるかをTrue/Falseで判定してください。
    見積書の工種・項目名・金額内訳から判断してください。
    - waterproof (防水工事あり): 防水工事が含まれていればTrue
    - kitchen_hood (厨房フード工事あり): 厨房フード・排気フードの工事が含まれていればTrue
    - signage (看板工事あり): 看板・サイン工事が含まれていればTrue
    - grease_trap (グリストラップあり): グリストラップの新設・設置が含まれていればTrue
    - exterior_sign (外部サイン工事あり): 外部サイン・ファサードサイン工事が含まれていればTrue

11. **outdoor_unit_floor**: 室外機の設置場所。以下から選択: 屋上, 地下, 屋外地上, その他, 未記入
    図面や見積書から室外機の設置場所が読み取れる場合は選択し、不明なら"未記入"。
    「屋外地上」は屋外に地上設置する場合を指す（店内に室外機が設置されることはない）。
12. **construction_days**: 工期。着工〜竣工の期間。以下から選択:
    未記入, 〜15日, 16〜30日, 31〜45日, 46〜60日, 61日〜
13. **construction_area**: 施工エリア。都道府県情報から推定。以下から選択:
    未記入, 一都三県, 北関東, 北海道・東北, 中部, 近畿, 中国・四国, 九州
14. **remarks_extra**: その他備考。夜間工事・居抜き・スケルトン・特殊条件などPDFから読み取れた注意事項（なければ空文字）。

15. **conditions**: 工事条件。PDF内容から読み取れる情報で以下の項目を推定してください。
    各項目に value（値）と confidence（"high"/"medium"/"low"）を返してください。
    - PDFに明確に記載がある → confidence="high"
    - 文脈から推測できる → confidence="medium"
    - 全く手がかりがない → confidence="low", value=""

    条件項目:
{conditions_spec}

## 出力JSON形式（厳密にこの形式で）

```json
{{
  "project_name": "案件名",
  "brand": "ブランド名",
  "category": "業態",
  "tsubo": 数値,
  "prefecture": "都道府県",
  "station": "最寄り駅",
  "total_amount": 数値（税別）,
  "trades": [
    {{
      "company": "会社名",
      "trade": "元の工種名",
      "mapped_category": "6カテゴリのいずれか（該当なしならnull）",
      "amount": 数値
    }}
  ],
  "construction_item_flags": {{
    "waterproof": true/false,
    "kitchen_hood": true/false,
    "signage": true/false,
    "grease_trap": true/false,
    "exterior_sign": true/false
  }},
  "outdoor_unit_floor": "屋上 or 地下 or 屋外地上 or その他 or 未記入",
  "construction_days": "〜15日 or 16〜30日 or 31〜45日 or 46〜60日 or 61日〜 or 未記入",
  "construction_area": "一都三県 or 北関東 or 北海道・東北 or 中部 or 近畿 or 中国・四国 or 九州 or 未記入",
  "remarks_extra": "特記事項（なければ空文字）",
  "conditions": {{
    "work_hours": {{"value": "日中", "confidence": "high"}},
    "location_type": {{"value": "商業施設", "confidence": "medium"}},
    ...
  }}
}}
```

## 図面PDFテキスト（先頭5ページ分）
{drawing_text}

## 見積書PDFテキスト（先頭3ページ分）
{estimate_text}

JSONのみを返してください。説明やマークダウンのコードブロック記法は不要です。"""

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


# ── 表示 ──

def display_results(data: dict):
    """抽出結果をきれいに表示する"""
    tsubo = data["tsubo"]
    total = data["total_amount"]

    print("=" * 60)
    print("  Apples to Apples データ抽出結果")
    print("=" * 60)
    print()
    print(f"  案件名:     {data['project_name']}")
    print(f"  ブランド:   {data['brand']}")
    print(f"  業態:       {data['category']}")
    print(f"  坪数:       {tsubo} 坪")
    print(f"  所在地:     {data['prefecture']} ({data['station']})")
    print(f"  合計金額:   ¥{total:,.0f}（税別）")
    print(f"  全体坪単価: ¥{total / tsubo:,.0f}/坪")
    print()
    print("-" * 60)
    print("  工種別内訳")
    print("-" * 60)
    print(f"  {'会社名':<16} {'工種':<18} {'金額':>14} {'坪単価':>12}")
    print(f"  {'─' * 14}  {'─' * 16}  {'─' * 14} {'─' * 12}")

    trade_total = 0
    for t in data["trades"]:
        amt = t["amount"]
        trade_total += amt
        unit = amt / tsubo
        mapped = t.get("mapped_category") or "-"
        print(f"  {t['company']:<14}  {t['trade']:<14}  ¥{amt:>12,.0f} ¥{unit:>10,.0f}/坪  → {mapped}")

    print(f"  {'─' * 14}  {'─' * 16}  {'─' * 14} {'─' * 12}")
    print(f"  {'合計':<14}  {'':<14}  ¥{trade_total:>12,.0f} ¥{trade_total / tsubo:>10,.0f}/坪")
    print()
    print("=" * 60)


# ── Google スプレッドシート書き込み ──

def get_gspread_client() -> gspread.Client:
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    creds = get_google_credentials(scopes)
    return gspread.authorize(creds)
def _set_dropdown_validation(sh: gspread.Spreadsheet, sheet_id: int, col_index: int, values: list[str]):
    """指定列にプルダウン（データ検証）を設定する"""
    sh.batch_update({"requests": [{
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,
                "startColumnIndex": col_index,
                "endColumnIndex": col_index + 1,
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": v} for v in values],
                },
                "showCustomUi": True,
                "strict": False,
            },
        }
    }]})


# 案件マスタの期待ヘッダー一覧
MASTER_HEADERS = [
    "案件ID", "案件名", "ブランド名", "業態", "工事種別", "坪数",
    "施工年月", "都道府県", "最寄り駅",
    "合計金額（税抜）",
    "内装", "電気", "給排水衛生", "空調換気", "ガス", "看板",
    "客席数", "厨房有無", "ドライブURL",
    "防水工事あり", "厨房フード工事あり", "看板工事あり",
    "グリストラップあり", "外部サイン工事あり",
    "入札結果", "他社受注金額", "他社受注坪単価",
    "室外機設置階数", "工期（日数）", "施工エリア", "その他備考",
]


def ensure_sheets(sh: gspread.Spreadsheet):
    """必要なシートがなければ作成し、ヘッダーを書き込む"""
    existing = {ws.title for ws in sh.worksheets()}

    # ①案件マスタ
    if SHEET_MASTER not in existing:
        ws = sh.add_worksheet(title=SHEET_MASTER, rows=200, cols=len(MASTER_HEADERS) + 5)
        ws.append_row(MASTER_HEADERS)
        ws.format(f"A1:{chr(64 + len(MASTER_HEADERS))}1", {"textFormat": {"bold": True}})
        # 工事種別列（E列）にデータ検証（新装/改装のプルダウン）を設定
        _set_dropdown_validation(sh, ws.id, col_index=4, values=["新装", "改装"])
        # 入札結果列にプルダウン設定
        from config import BID_RESULT_OPTIONS
        bid_col_idx = MASTER_HEADERS.index("入札結果") if "入札結果" in MASTER_HEADERS else 25
        _set_dropdown_validation(sh, ws.id, col_index=bid_col_idx, values=BID_RESULT_OPTIONS)
        print(f"  シート「{SHEET_MASTER}」を作成しました")
    else:
        # 既存シートのヘッダーを確認し、不足列があれば追加する
        ws = sh.worksheet(SHEET_MASTER)
        current_headers = ws.row_values(1)
        missing = [h for h in MASTER_HEADERS if h not in current_headers]
        if missing:
            # 末尾に不足ヘッダーを追加
            next_col = len(current_headers) + 1
            for i, h in enumerate(missing):
                ws.update_cell(1, next_col + i, h)
            print(f"  案件マスタに列を追加しました: {missing}")

    # ②工種別金額
    if SHEET_TRADES not in existing:
        ws = sh.add_worksheet(title=SHEET_TRADES, rows=500, cols=12)
        ws.append_row([
            "案件ID", "案件名", "業態", "坪数", "工種",
            "金額（税抜）", "坪単価",
            "全体合計（税抜）", "全体坪単価",
            "施工年月", "備考", "ドライブURL",
        ])
        ws.format("A1:L1", {"textFormat": {"bold": True}})
        print(f"  シート「{SHEET_TRADES}」を作成しました")

    # ③見積明細
    if SHEET_DETAIL not in existing:
        ws = sh.add_worksheet(title=SHEET_DETAIL, rows=1000, cols=12)
        ws.append_row([
            "案件ID", "案件名", "業者名", "工種", "大分類",
            "項目名", "仕様・規格", "数量", "単位", "単価", "金額", "備考",
        ])
        ws.format("A1:L1", {"textFormat": {"bold": True}})
        print(f"  シート「{SHEET_DETAIL}」を作成しました")


def next_project_id(ws: gspread.Worksheet) -> str:
    """案件マスタの既存行数から次の案件IDを生成する"""
    all_values = ws.col_values(1)  # A列
    # ヘッダー行を除いたデータ行数
    data_rows = [v for v in all_values[1:] if v]
    next_num = len(data_rows) + 1
    return f"P-{next_num:03d}"


def write_to_spreadsheet(data: dict):
    """抽出データをGoogleスプレッドシートに書き込む"""
    print("Googleスプレッドシートに接続中...")
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    ensure_sheets(sh)

    ws_master = sh.worksheet(SHEET_MASTER)
    ws_trades = sh.worksheet(SHEET_TRADES)

    project_id = next_project_id(ws_master)
    # None または 0 の場合に安全に処理
    tsubo = data.get("tsubo") or None
    if tsubo is not None:
        try:
            tsubo = float(tsubo)
        except (TypeError, ValueError):
            tsubo = None
    total = data.get("total_amount") or None
    if total is not None:
        try:
            total = int(total)
        except (TypeError, ValueError):
            total = None
    construction_date = date.today().strftime("%Y-%m")

    # 工種別金額を6カテゴリに集計
    category_amounts = {cat: 0 for cat in TRADE_CATEGORIES}
    for t in data["trades"]:
        mapped = t.get("mapped_category")
        if mapped and mapped in category_amounts:
            category_amounts[mapped] += t["amount"]

    # 工事項目フラグの取得
    from config import CONSTRUCTION_ITEM_FLAGS
    item_flags = data.get("construction_item_flags", {})
    flag_dict = {label: str(item_flags.get(key, False)) for key, label in CONSTRUCTION_ITEM_FLAGS}

    # ── ①案件マスタ に書き込み ──
    # ヘッダー名で列を特定して書き込む（列順のズレを防ぐ）
    master_headers = ws_master.row_values(1)

    master_data = {
        "案件ID": project_id,
        "案件名": data.get("project_name", ""),
        "ブランド名": data.get("brand", ""),
        "業態": data.get("category", ""),
        "工事種別": "",
        "坪数": tsubo or "",
        "施工年月": construction_date,
        "都道府県": data.get("prefecture", ""),
        "最寄り駅": data.get("station", ""),
        "合計金額（税抜）": total or "",
        "内装": category_amounts["内装"],
        "電気": category_amounts["電気"],
        "給排水衛生": category_amounts["給排水衛生"],
        "空調換気": category_amounts["空調換気"],
        "ガス": category_amounts["ガス"],
        "看板": category_amounts["看板"],
        "客席数": "",
        "厨房有無": "",
        "ドライブURL": "",
        **flag_dict,
        "入札結果": "",
        "他社受注金額": "",
        "他社受注坪単価": "",
        "室外機設置階数": data.get("outdoor_unit_floor", ""),
        "工期（日数）": data.get("construction_days", ""),
        "施工エリア": data.get("construction_area", ""),
        "その他備考": data.get("remarks_extra", ""),
    }

    if master_headers:
        # 既存ヘッダーに合わせて行を構築
        master_row = [master_data.get(h, "") for h in master_headers]
    else:
        # ヘッダーがない場合はデフォルト順序
        master_row = list(master_data.values())

    ws_master.append_row(master_row, value_input_option="USER_ENTERED")
    print(f"  → 案件マスタに書き込み完了（{project_id}）")

    # ── ②工種別金額 に書き込み ──
    trade_rows = []
    for t in data["trades"]:
        mapped = t.get("mapped_category")
        if not mapped:
            continue
        amt = t.get("amount")
        if amt is None:
            try:
                amt = int(amt)
            except (TypeError, ValueError):
                amt = 0
        amt = amt or 0
        unit_price = round(amt / tsubo) if tsubo else ""
        total_unit = round(total / tsubo) if tsubo and total else ""
        trade_rows.append([
            project_id,
            data.get("project_name", ""),
            data.get("category", ""),
            tsubo or "",
            mapped,               # 工種
            amt,                   # 金額
            unit_price,            # 嵪単価
            total or "",           # 全体合計
            total_unit,            # 全体嵪単価
            construction_date,
            t.get("company", ""),  # 備考に会社名
            "",                    # ドライブURL
        ])

    if trade_rows:
        ws_trades.append_rows(trade_rows, value_input_option="USER_ENTERED")
        print(f"  → 工種別金額に {len(trade_rows)} 行書き込み完了")

    print()
    print(f"  スプレッドシートURL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")


# ── 工事条件の保存 ──

SHEET_CONDITIONS = "案件サマリー"

def write_conditions_to_spreadsheet(project_id: str, project_name: str, conditions: dict):
    """確定済みの工事条件をスプレッドシートの案件サマリーシートに書き込む"""
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    existing = {ws.title for ws in sh.worksheets()}

    # ヘッダー定義
    headers = ["案件ID", "案件名"] + [f[1] for f in CONDITION_FIELDS]

    if SHEET_CONDITIONS not in existing:
        ws = sh.add_worksheet(title=SHEET_CONDITIONS, rows=200, cols=len(headers) + 2)
        ws.append_row(headers)
        ws.format(f"A1:{chr(64 + len(headers))}1", {"textFormat": {"bold": True}})
    else:
        ws = sh.worksheet(SHEET_CONDITIONS)

    # データ行
    row = [project_id, project_name]
    for key, _, _ in CONDITION_FIELDS:
        row.append(conditions.get(key, ""))

    ws.append_row(row, value_input_option="USER_ENTERED")


# ── 案件マスタ更新 ──

def update_project_fields(project_id: str, fields: dict) -> bool:
    """
    案件マスタの指定案件IDの行に対し、fieldsで指定された列のみを更新する。
    fields = {"工事種別": "新装", "施工エリア": "東京都心（山手線内）", ...}
    空文字の値はスキップする。
    """
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(SPREADSHEET_ID)
        ws = sh.worksheet(SHEET_MASTER)
        headers = ws.row_values(1)

        # 案件IDの列（A列）から対象行を検索
        id_col = ws.col_values(1)
        row_idx = None
        for i, v in enumerate(id_col):
            if v == project_id:
                row_idx = i + 1  # 1-indexed
                break

        if row_idx is None:
            return False

        # 各フィールドの列を特定して更新
        cells_to_update = []
        for col_name, col_val in fields.items():
            if col_val == "" or col_val is None:
                continue
            if col_name in headers:
                col_idx = headers.index(col_name) + 1  # 1-indexed
                cells_to_update.append((row_idx, col_idx, col_val))

        for r, c, v in cells_to_update:
            ws.update_cell(r, c, v)

        return True
    except Exception:
        return False


# ── メイン ──

def main():
    import argparse

    parser = argparse.ArgumentParser(description="Apples to Apples PDF データ抽出")
    parser.add_argument("--drawing", "-d", help="図面PDFのパス")
    parser.add_argument("--estimate", "-e", help="見積書PDFのパス")
    args = parser.parse_args()

    input_dir = os.path.dirname(os.path.abspath(__file__))

    if args.drawing and args.estimate:
        drawing_pdf = os.path.abspath(args.drawing)
        estimate_pdf = os.path.abspath(args.estimate)
    else:
        drawing_pdf, estimate_pdf = find_pdfs(input_dir)

    if not drawing_pdf or not os.path.exists(drawing_pdf):
        print("エラー: 図面PDFが見つかりません。-d オプションで指定してください。", file=sys.stderr)
        sys.exit(1)
    if not estimate_pdf or not os.path.exists(estimate_pdf):
        print("エラー: 見積書PDFが見つかりません。-e オプションで指定してください。", file=sys.stderr)
        sys.exit(1)

    print(f"図面PDF:   {os.path.basename(drawing_pdf)}")
    print(f"見積書PDF: {os.path.basename(estimate_pdf)}")
    print()
    print("PDFからテキストを抽出中...")

    drawing_text = extract_text_from_pdf(drawing_pdf, max_pages=5)
    estimate_text = extract_text_from_pdf(estimate_pdf, max_pages=3)

    print("GPT-4oでデータを解析中...")
    print()

    client = get_openai_client()
    data = extract_with_gpt4o(client, drawing_text, estimate_text)

    display_results(data)

    print()
    write_to_spreadsheet(data)


if __name__ == "__main__":
    main()
