"""
Apples to Apples 類似案件比較・ベース金額算出スクリプト
- スプレッドシートから過去案件を取得
- 業態・ブランドで類似案件を検索
- 坪単価ベースでベース金額を算出
"""

import argparse
import os
import sys

import gspread

# ── 設定 ──
SPREADSHEET_ID = "10uXWjPTuYcMtnvmWt6A9fMWRxvPlVf82vIVpM90u95U"

SHEET_MASTER = "案件マスタ"
SHEET_TRADES = "工種別金額"

TRADE_CATEGORIES = ["内装", "電気", "給排水衛生", "空調換気", "ガス", "看板"]

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import get_google_credentials


# ── スプレッドシート接続 ──

def get_spreadsheet() -> gspread.Spreadsheet:
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = get_google_credentials(scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


# ── データ取得 ──

def _safe_get_all_records(ws: gspread.Worksheet) -> list[dict]:
    """空白・重複ヘッダーがあっても安全にレコードを取得する（全列空の行はスキップ）"""
    data = ws.get_all_values()
    if not data:
        return []
    headers = data[0]
    # ヘッダー行が空の場合は空リストを返す
    if not any(h.strip() for h in headers):
        return []
    records = []
    for row in data[1:]:
        # 行全体が空の場合はスキップ
        if not any(cell.strip() for cell in row):
            continue
        # 行の長さがヘッダーより短い場合はパディング
        padded_row = row + [""] * (len(headers) - len(row))
        record = dict(zip(headers, padded_row))
        records.append(record)
    return records


def load_master(sh: gspread.Spreadsheet) -> list[dict]:
    """案件マスタから全案件を読み込む"""
    ws = sh.worksheet(SHEET_MASTER)
    return _safe_get_all_records(ws)


def load_trades(sh: gspread.Spreadsheet) -> list[dict]:
    """工種別金額から全データを読み込む"""
    ws = sh.worksheet(SHEET_TRADES)
    return _safe_get_all_records(ws)


# ── 類似案件検索 ──

def score_similarity(project: dict, brand: str, category: str) -> int:
    """類似度スコアを計算（高いほど類似）"""
    score = 0
    p_brand = project.get("ブランド名", "")
    p_category = project.get("業態", "")

    # ブランド一致（最も重要）
    if brand and p_brand and brand.lower() in p_brand.lower():
        score += 10
    elif brand and p_brand and p_brand.lower() in brand.lower():
        score += 10

    # 業態一致
    if category and p_category:
        if category == p_category:
            score += 5
        elif category in p_category or p_category in category:
            score += 3

    return score


def find_similar_projects(master: list[dict], brand: str, category: str) -> list[dict]:
    """類似案件を類似度順に返す"""
    scored = []
    for p in master:
        s = score_similarity(p, brand, category)
        scored.append((s, p))

    # スコア降順でソート（同スコアなら元順維持）
    scored.sort(key=lambda x: x[0], reverse=True)
    return [(s, p) for s, p in scored]


# ── ベース金額計算 ──

def calc_base_amount(ref_project: dict, ref_trades: list[dict], new_tsubo: float) -> dict:
    """参照案件の坪単価 × 新規坪数 でベース金額を算出"""
    ref_tsubo = float(ref_project.get("坪数", 0))
    if ref_tsubo == 0:
        return {}

    result = {}
    for cat in TRADE_CATEGORIES:
        # 工種別金額シートからこの案件・工種の金額を取得
        matched = [t for t in ref_trades
                   if t.get("案件ID") == ref_project.get("案件ID") and t.get("工種") == cat]
        if matched:
            ref_amount = int(float(matched[0].get("金額（税抜）", 0)))
        else:
            # 案件マスタのカラムからフォールバック
            ref_amount = int(float(ref_project.get(cat, 0) or 0))

        unit_price = ref_amount / ref_tsubo if ref_tsubo else 0
        base = round(unit_price * new_tsubo)
        result[cat] = {
            "ref_amount": ref_amount,
            "unit_price": round(unit_price),
            "base_amount": base,
        }

    total_ref = int(sum(v["ref_amount"] for v in result.values()))
    total_base = int(sum(v["base_amount"] for v in result.values()))
    result["合計"] = {
        "ref_amount": total_ref,
        "unit_price": round(total_ref / ref_tsubo) if ref_tsubo else 0,
        "base_amount": total_base,
    }

    return result


# ── 工事項目差分による補正 ──

def calc_item_correction(
    ref_project: dict,
    new_flags: dict[str, bool],
    new_tsubo: float,
    ref_trades: list[dict] | None = None,
) -> list[dict]:
    """
    類似案件と新規案件の工事項目フラグ差分から補正額を算出する。

    ref_project: 参照案件の案件マスタレコード（工事項目フラグ列を含む）
    new_flags:   新規案件の工事項目フラグ {key: True/False}
    new_tsubo:   新規案件の坪数
    ref_trades:  工種別金額データ（坪単価の上書き用、任意）

    Returns: [{"key", "label", "ref_has", "new_has", "direction", "amount"}]
      direction: "subtract" | "add" | None
      amount: 補正金額（正の値。direction で加減を判断）
    """
    from config import CONSTRUCTION_ITEM_FLAGS, CONSTRUCTION_ITEM_UNIT_PRICES

    corrections = []
    for key, label in CONSTRUCTION_ITEM_FLAGS:
        ref_val_raw = ref_project.get(label, "")
        ref_has = str(ref_val_raw).strip().lower() in ("true", "1", "yes", "○")
        new_has = bool(new_flags.get(key, False))

        if ref_has == new_has:
            continue  # 差分なし

        unit_price = CONSTRUCTION_ITEM_UNIT_PRICES.get(key, 0)

        # 工種別金額から実際の坪単価を取得できる場合は上書き
        if ref_trades and ref_has:
            ref_tsubo = float(ref_project.get("坪数", 0) or 0)
            # 看板工事・防水工事などは工種名で近似マッチ
            _trade_map = {
                "waterproof": "内装",  # 防水は内装に含まれることが多い
                "kitchen_hood": "空調換気",
                "signage": "看板",
                "grease_trap": "給排水衛生",
                "exterior_sign": "看板",
            }
            mapped_trade = _trade_map.get(key)
            if mapped_trade and ref_tsubo > 0:
                matched = [t for t in ref_trades
                           if t.get("案件ID") == ref_project.get("案件ID")
                           and t.get("工種") == mapped_trade]
                if matched:
                    trade_amount = int(float(matched[0].get("金額（税抜）", 0)))
                    trade_unit = trade_amount / ref_tsubo
                    # フラグに対応する工事の割合を推定（全額ではなく一部）
                    # デフォルト坪単価と実際の坪単価の小さい方を採用
                    unit_price = min(unit_price, round(trade_unit * 0.3))
                    if unit_price < 5000:
                        unit_price = CONSTRUCTION_ITEM_UNIT_PRICES.get(key, 0)

        amount = round(unit_price * new_tsubo)

        if ref_has and not new_has:
            direction = "subtract"
        else:
            direction = "add"

        corrections.append({
            "key": key,
            "label": label,
            "ref_has": ref_has,
            "new_has": new_has,
            "direction": direction,
            "unit_price": unit_price,
            "amount": amount,
        })

    return corrections


# ── 表示 ──

def display_similar_projects(scored_projects: list[tuple[int, dict]]):
    """類似案件一覧を表示"""
    print("-" * 80)
    print("  過去案件一覧（類似度順）")
    print("-" * 80)
    print(f"  {'No':<4} {'案件ID':<8} {'案件名':<30} {'業態':<12} {'坪数':>6} {'合計金額':>14} {'類似度':>6}")
    print(f"  {'─'*3}  {'─'*6}  {'─'*28}  {'─'*10}  {'─'*6} {'─'*14} {'─'*6}")

    for i, (score, p) in enumerate(scored_projects):
        pid = p.get("案件ID", "")
        name = p.get("案件名", "")[:28]
        cat = p.get("業態", "")[:10]
        tsubo = p.get("坪数", "")
        total = p.get("合計金額（税抜）", 0)
        try:
            total_str = f"¥{int(total):>12,}"
        except (ValueError, TypeError):
            total_str = f"{'N/A':>14}"
        star = "★" * min(score // 3, 5) if score > 0 else "-"
        print(f"  {i+1:<4} {pid:<8} {name:<28}  {cat:<10}  {tsubo:>5}  {total_str} {star}")
    print()


def display_base_calculation(ref_project: dict, calc: dict, new_tsubo: float):
    """ベース金額計算結果を表示"""
    ref_tsubo = ref_project.get("坪数", 0)
    ref_name = ref_project.get("案件名", "")
    ref_id = ref_project.get("案件ID", "")

    print("=" * 80)
    print("  Apples to Apples ベース金額算出結果")
    print("=" * 80)
    print()
    print(f"  参照案件:   [{ref_id}] {ref_name}")
    print(f"  参照坪数:   {ref_tsubo} 坪")
    print(f"  新規坪数:   {new_tsubo} 坪")
    print()
    print("-" * 80)
    print(f"  {'工種':<12} {'参照金額':>14} {'坪単価':>12}   {'×':>1} {'新規坪数':>6}  {'＝':>1}  {'ベース金額':>14}")
    print(f"  {'─'*10}  {'─'*14} {'─'*12}   {'─'*1} {'─'*6}  {'─'*1}  {'─'*14}")

    for cat in TRADE_CATEGORIES:
        v = calc.get(cat, {})
        if not v or v["ref_amount"] == 0:
            continue
        print(f"  {cat:<10}  ¥{v['ref_amount']:>12,}  ¥{v['unit_price']:>10,}/坪"
              f"   × {new_tsubo:>5}   =  ¥{v['base_amount']:>12,}")

    v = calc["合計"]
    print(f"  {'─'*10}  {'─'*14} {'─'*12}   {'─'*1} {'─'*6}  {'─'*1}  {'─'*14}")
    print(f"  {'合計':<10}  ¥{v['ref_amount']:>12,}  ¥{v['unit_price']:>10,}/坪"
          f"   × {new_tsubo:>5}   =  ¥{v['base_amount']:>12,}")
    print()

    # 最終金額サマリー
    base_total = v["base_amount"]
    tax = round(base_total * 0.1)
    direction = round(base_total * 0.05)

    print("-" * 80)
    print("  概算サマリー")
    print("-" * 80)
    print(f"  ベース金額（税抜）:       ¥{base_total:>14,}")
    print(f"  消費税（10%）:            ¥{tax:>14,}")
    print(f"  ベース金額（税込）:       ¥{base_total + tax:>14,}")
    print(f"  ディレクション費（5%）:   ¥{direction:>14,}")
    print(f"  請求総額（税込）:         ¥{base_total + tax + round(direction * 1.1):>14,}")
    print()
    print("  ※ 増減項目は含まれていません。個別の追加・削除項目を考慮して最終金額を調整してください。")
    print("=" * 80)


# ── メイン ──

def main():
    parser = argparse.ArgumentParser(description="Apples to Apples 類似案件比較")
    parser.add_argument("--tsubo", "-t", type=float, required=True, help="新規案件の坪数")
    parser.add_argument("--brand", "-b", default="", help="ブランド名")
    parser.add_argument("--category", "-c", default="", help="業態")
    parser.add_argument("--ref", "-r", default="", help="参照する案件ID（指定しない場合は最も類似度の高い案件を自動選択）")
    args = parser.parse_args()

    print("スプレッドシートからデータを取得中...")
    sh = get_spreadsheet()
    master = load_master(sh)
    trades = load_trades(sh)

    if not master:
        print("エラー: 案件マスタにデータがありません。", file=sys.stderr)
        sys.exit(1)

    print(f"  → {len(master)} 件の過去案件を取得しました")
    print()

    # 類似案件の検索・表示
    scored = find_similar_projects(master, args.brand, args.category)
    display_similar_projects(scored)

    # 参照案件の決定
    if args.ref:
        ref = next((p for _, p in scored if p.get("案件ID") == args.ref), None)
        if not ref:
            print(f"エラー: 案件ID '{args.ref}' が見つかりません。", file=sys.stderr)
            sys.exit(1)
    else:
        # 最も類似度の高い案件を自動選択
        top_score, ref = scored[0]
        if top_score == 0:
            print("類似案件が見つかりません。-r オプションで案件IDを直接指定してください。")
            sys.exit(1)
        print(f"  → 最も類似度の高い案件を自動選択: [{ref.get('案件ID')}] {ref.get('案件名')}")
        print()

    # ベース金額計算
    calc = calc_base_amount(ref, trades, args.tsubo)
    display_base_calculation(ref, calc, args.tsubo)


if __name__ == "__main__":
    main()
