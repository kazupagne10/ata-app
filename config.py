"""
Apples to Apples 共通設定
"""

import os

# ── パス ──
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")

# ── Google Spreadsheet ──
SPREADSHEET_ID = "10uXWjPTuYcMtnvmWt6A9fMWRxvPlVf82vIVpM90u95U"

# ── Google Drive フォルダ ID ──
DRIVE_INBOX_FOLDER_ID = "1uql7bBBnplhg8vfyfY7s6WoRY3owPBku"
INBOX_FOLDER_ID = DRIVE_INBOX_FOLDER_ID  # 短縮エイリアス
DRIVE_PROCESSED_FOLDER_NAME = "AtA_processed"

# ── Drive 監視 ──
DRIVE_POLL_INTERVAL = 300  # 秒（5分）
DRIVE_STATE_FILE = os.path.join(BASE_DIR, "drive_state.json")

# ── 工種カテゴリ ──
TRADE_CATEGORIES = ["内装", "電気", "給排水衛生", "空調換気", "ガス", "看板"]

# ── 業態マスタ ──
GYOTAI_OPTIONS = [
    "飲食(防水あり)",
    "飲食(簡易防水)",
    "飲食(防水B工事)",
    "飲食(防水なし)",
    "物販",
    "オフィス",
    "その他",
]

# ── 工事項目フラグ（差分補正用） ──
CONSTRUCTION_ITEM_FLAGS = [
    ("waterproof",    "防水工事あり"),
    ("kitchen_hood",  "厨房フード工事あり"),
    ("signage",       "看板工事あり"),
    ("grease_trap",   "グリストラップあり"),
    ("exterior_sign", "外部サイン工事あり"),
]

# 工事項目ごとのデフォルト坪単価（推計用）
CONSTRUCTION_ITEM_UNIT_PRICES = {
    "waterproof":    40_000,   # 防水工事 円/坪
    "kitchen_hood":  35_000,   # 厨房フード工事 円/坪
    "signage":       25_000,   # 看板工事 円/坪
    "grease_trap":   20_000,   # グリストラップ 円/坪
    "exterior_sign": 30_000,   # 外部サイン工事 円/坪
}

# ── 室外機設置階数 ──
OUTDOOR_UNIT_FLOOR_OPTIONS = ["未記入", "屋上", "地下", "屋外地上", "その他"]

# ── 施工エリア ──
CONSTRUCTION_AREA_OPTIONS = [
    "未記入", "一都三県", "北関東", "北海道・東北",
    "中部", "近畿", "中国・四国", "九州",
]

# ── 工期 ──
CONSTRUCTION_DAYS_OPTIONS = [
    "未記入", "〜15日", "16〜30日", "31〜45日", "46〜60日", "61日〜",
]

# ── 入札結果 ──
BID_RESULT_OPTIONS = ["自社受注", "他社受注", "随意契約", "未定"]


# ── OpenAI ヘルパー ──

def get_openai_client():
    """
    Streamlit Secrets（本番）または環境変数（ローカル）から
    OpenAI クライアントを生成する。
    Secrets の OPENAI_API_KEY はトップレベルまたは
    [gcp_service_account] セクション内のどちらにあっても対応する。
    """
    from openai import OpenAI

    api_key = None

    # Streamlit Cloud 環境
    try:
        import streamlit as st
        # トップレベルにある場合
        if "OPENAI_API_KEY" in st.secrets:
            api_key = st.secrets["OPENAI_API_KEY"]
        # [gcp_service_account] セクション内にある場合
        elif "gcp_service_account" in st.secrets and "OPENAI_API_KEY" in st.secrets["gcp_service_account"]:
            api_key = st.secrets["gcp_service_account"]["OPENAI_API_KEY"]
    except Exception:
        pass

    # ローカル環境: 環境変数から取得
    if not api_key:
        api_key = os.environ.get("OPENAI_API_KEY")

    if not api_key:
        raise ValueError(
            "OpenAI APIキーが見つかりません。\n"
            "Streamlit Cloud の Secrets に OPENAI_API_KEY を設定してください。"
        )

    return OpenAI(api_key=api_key)


# ── Google 認証ヘルパー ──

def get_google_credentials(scopes: list):
    """
    Streamlit Secrets（本番）または credentials.json（ローカル）から
    Google サービスアカウント認証情報を取得する。
    """
    from google.oauth2.service_account import Credentials

    # Streamlit Cloud 環境: st.secrets に gcp_service_account が存在する場合
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            info = dict(st.secrets["gcp_service_account"])
            return Credentials.from_service_account_info(info, scopes=scopes)
    except Exception:
        pass

    # ローカル環境: credentials.json ファイルから読み込む
    if os.path.exists(CREDENTIALS_FILE):
        return Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)

    raise FileNotFoundError(
        "Google 認証情報が見つかりません。\n"
        "Streamlit Cloud の Secrets に [gcp_service_account] を設定するか、"
        "credentials.json をプロジェクトルートに配置してください。"
    )
