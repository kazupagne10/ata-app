"""
Apples to Apples Google Drive 自動監視・取り込みスクリプト
- AtA_inbox フォルダ（ID指定）を監視して新規PDFを検出
- GPT-4oでファイル内容を読み、図面か見積書かを自動判定
- ペアが揃ったら ata_extract で抽出・スプレッドシート書き込み
- 処理済みファイルは AtA_processed フォルダに移動
- 状態ファイル（JSON）で最終チェック日時・未処理数を管理
- バックグラウンドスレッドとして Streamlit 内から起動可能
"""

import json
import logging
import os
import sys
import tempfile
import threading
import time
from datetime import datetime, timezone

import pdfplumber
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from openai import OpenAI

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ata_extract
from config import (
    CREDENTIALS_FILE,
    DRIVE_INBOX_FOLDER_ID,
    DRIVE_POLL_INTERVAL,
    DRIVE_PROCESSED_FOLDER_NAME,
    DRIVE_STATE_FILE,
)

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

# バックグラウンドスレッド管理
_watcher_thread: threading.Thread | None = None
_watcher_lock = threading.Lock()


# ── 状態管理 ──

def load_state() -> dict:
    if os.path.exists(DRIVE_STATE_FILE):
        try:
            with open(DRIVE_STATE_FILE) as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {"last_check": None, "pending_count": 0, "processed": []}


def save_state(state: dict):
    tmp_path = DRIVE_STATE_FILE + ".tmp"
    with open(tmp_path, "w") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, DRIVE_STATE_FILE)


def get_last_check_ago() -> str:
    """最終チェックからの経過時間を人間可読な文字列で返す"""
    state = load_state()
    last_check = state.get("last_check")
    if not last_check:
        return "未実行"
    try:
        dt = datetime.fromisoformat(last_check)
        now = datetime.now(timezone.utc)
        delta = now - dt
        minutes = int(delta.total_seconds() / 60)
        if minutes < 1:
            return "たった今"
        if minutes < 60:
            return f"{minutes}分前"
        hours = minutes // 60
        if hours < 24:
            return f"{hours}時間前"
        days = hours // 24
        return f"{days}日前"
    except (ValueError, TypeError):
        return "-"


# ── Google Drive API ──

def get_drive_service():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def find_or_create_processed_folder(service) -> str:
    """AtA_processed フォルダを検索し、なければ作成する"""
    query = (
        f"name = '{DRIVE_PROCESSED_FOLDER_NAME}'"
        " and mimeType = 'application/vnd.google-apps.folder'"
        " and trashed = false"
    )
    results = service.files().list(
        q=query, spaces="drive", fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = results.get("files", [])

    if files:
        return files[0]["id"]

    metadata = {
        "name": DRIVE_PROCESSED_FOLDER_NAME,
        "mimeType": "application/vnd.google-apps.folder",
    }
    folder = service.files().create(
        body=metadata, fields="id", supportsAllDrives=True,
    ).execute()
    logger.info("フォルダ「%s」を作成: %s", DRIVE_PROCESSED_FOLDER_NAME, folder["id"])
    return folder["id"]


def list_pdfs_in_folder(service, folder_id: str) -> list[dict]:
    """フォルダ直下のPDFファイル一覧を取得（[処理済] を除外）"""
    query = (
        f"'{folder_id}' in parents"
        " and mimeType = 'application/pdf'"
        " and trashed = false"
        " and not name contains '[処理済]'"
    )
    results = service.files().list(
        q=query, spaces="drive",
        fields="files(id, name, createdTime)",
        orderBy="createdTime",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    return results.get("files", [])


def _list_subfolders(service, folder_id: str) -> list[dict]:
    """フォルダ直下のサブフォルダ一覧を取得"""
    query = (
        f"'{folder_id}' in parents"
        " and mimeType = 'application/vnd.google-apps.folder'"
        " and trashed = false"
    )
    results = service.files().list(
        q=query, spaces="drive",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    return results.get("files", [])


def _classify_hint_from_folder_name(folder_name: str) -> str | None:
    """フォルダ名からPDF種別のヒントを返す。該当なしなら None"""
    if "見積" in folder_name:
        return "estimate"
    if "図面" in folder_name or "図" in folder_name or "入札" in folder_name or "平面" in folder_name or "設計" in folder_name:
        return "drawing"
    return None


def list_pdfs_recursive(service, folder_id: str, folder_hint: str | None = None) -> list[dict]:
    """フォルダを再帰的に探索し、全PDFを返す。
    各PDFには 'folder_hint' キー（'drawing'|'estimate'|None）が付与される。
    """
    all_pdfs = []

    # このフォルダ直下のPDF
    pdfs = list_pdfs_in_folder(service, folder_id)
    for pdf in pdfs:
        pdf["folder_hint"] = folder_hint
    all_pdfs.extend(pdfs)

    # サブフォルダを再帰探索
    subfolders = _list_subfolders(service, folder_id)
    for sf in subfolders:
        # サブフォルダ名からヒントを判定（親のヒントより優先）
        sub_hint = _classify_hint_from_folder_name(sf["name"]) or folder_hint
        logger.info("  サブフォルダ探索: %s (hint=%s)", sf["name"], sub_hint or "なし")
        sub_pdfs = list_pdfs_recursive(service, sf["id"], folder_hint=sub_hint)
        all_pdfs.extend(sub_pdfs)

    return all_pdfs


def download_pdf(service, file_id: str, dest_path: str):
    """Google DriveからPDFをダウンロード"""
    request = service.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def move_file(service, file_id: str, dest_folder_id: str):
    """ファイルを別フォルダに移動する。権限不足の場合はリネームでフォールバック。"""
    try:
        file_info = service.files().get(
            fileId=file_id, fields="parents", supportsAllDrives=True,
        ).execute()
        previous_parents = ",".join(file_info.get("parents", []))
        service.files().update(
            fileId=file_id,
            addParents=dest_folder_id,
            removeParents=previous_parents,
            fields="id, parents",
            supportsAllDrives=True,
        ).execute()
    except Exception as e:
        if "insufficientFilePermissions" in str(e) or "403" in str(e):
            logger.warning("移動権限なし → リネームで処理済みマーク: %s", file_id)
            mark_as_processed(service, file_id)
        else:
            raise


def mark_as_processed(service, file_id: str):
    """ファイル名に [処理済] を付与して処理済みとマークする"""
    original = service.files().get(
        fileId=file_id, fields="name", supportsAllDrives=True,
    ).execute()
    original_name = original["name"]
    if original_name.startswith("[処理済]"):
        return  # 既にマーク済み
    new_name = f"[処理済] {original_name}"
    service.files().update(
        fileId=file_id,
        body={"name": new_name},
        fields="id, name",
        supportsAllDrives=True,
    ).execute()
    print(f"[DEBUG]     リネーム完了: {new_name}")


# ── PDF分類 ──

def classify_pdf(client: OpenAI, text: str, filename: str, folder_hint: str | None = None) -> str:
    """PDFの種別を判定する。優先順位:
    1. フォルダ名ヒント（'drawing' or 'estimate'）があれば即採用
    2. ファイル名キーワード
    3. GPT-4oによる内容判定
    返り値: 'drawing' | 'estimate' | 'unknown'
    """
    # (1) フォルダ名ヒント
    if folder_hint in ("drawing", "estimate"):
        logger.info("    フォルダ名から判定: %s → %s", filename, folder_hint)
        return folder_hint

    # (2) ファイル名キーワード
    if "見積" in filename:
        return "estimate"
    if "図" in filename or "入札" in filename or "平面" in filename or "設計" in filename:
        return "drawing"

    # (3) GPT-4oで内容判定
    return _classify_pdf_with_gpt(client, text, filename)


def _classify_pdf_with_gpt(client: OpenAI, text: str, filename: str) -> str:
    """GPT-4oでPDFが図面か見積書かを判定する"""
    prompt = f"""以下のPDFの内容を分析して、このPDFが「図面（設計図・平面図・入札図）」か「見積書」かを判定してください。

## ファイル名
{filename}

## PDFテキスト（先頭3ページ分）
{text[:3000]}

以下のいずれか1語のみを返してください:
- drawing（図面・設計図・平面図・入札図の場合）
- estimate（見積書の場合）
- unknown（判定できない場合）"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
        max_tokens=10,
    )
    result = response.choices[0].message.content.strip().lower()

    if result in ("drawing", "estimate"):
        return result

    return "unknown"


# ── ペアリングと処理 ──

def group_files_by_project(classified: list[dict]) -> dict[str, dict]:
    """分類済みファイルから図面+見積書のペアを作成する"""
    drawings = [f for f in classified if f["type"] == "drawing"]
    estimates = [f for f in classified if f["type"] == "estimate"]

    pairs = {}
    used_estimates = set()

    for d in drawings:
        best_match = None
        best_score = 0
        for e in estimates:
            if e["id"] in used_estimates:
                continue
            score = _name_similarity(d["name"], e["name"])
            if score > best_score:
                best_score = score
                best_match = e

        if best_match:
            used_estimates.add(best_match["id"])
            pair_key = d["name"].split(".")[0]
            pairs[pair_key] = {"drawing": d, "estimate": best_match}

    return pairs


def _name_similarity(name1: str, name2: str) -> int:
    """ファイル名の類似度スコア（共通文字数ベース）"""
    n1 = set(name1.replace(".pdf", "").replace(" ", ""))
    n2 = set(name2.replace(".pdf", "").replace(" ", ""))
    return len(n1 & n2)


def process_pair(service, openai_client: OpenAI, pair: dict, processed_folder_id: str) -> bool:
    """図面+見積書のペアを処理する"""
    drawing_info = pair["drawing"]
    estimate_info = pair["estimate"]

    logger.info("処理開始: 図面=%s, 見積書=%s", drawing_info["name"], estimate_info["name"])

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f_d:
        download_pdf(service, drawing_info["id"], f_d.name)
        drawing_path = f_d.name

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f_e:
        download_pdf(service, estimate_info["id"], f_e.name)
        estimate_path = f_e.name

    try:
        drawing_text = ata_extract.extract_text_from_pdf(drawing_path, max_pages=5)
        estimate_text = ata_extract.extract_text_from_pdf(estimate_path, max_pages=3)

        logger.info("GPT-4oでデータを解析中...")
        data = ata_extract.extract_with_gpt4o(openai_client, drawing_text, estimate_text)

        logger.info("スプレッドシートに書き込み中...")
        ata_extract.write_to_spreadsheet(data)

        move_file(service, drawing_info["id"], processed_folder_id)
        move_file(service, estimate_info["id"], processed_folder_id)
        logger.info("AtA_processed に移動完了")

        return True

    except Exception as e:
        logger.error("処理エラー: %s", e, exc_info=True)
        return False

    finally:
        os.unlink(drawing_path)
        os.unlink(estimate_path)


# ── メインチェック ──

def check_inbox_once(service=None, openai_client=None) -> dict:
    """inboxを1回チェックし、結果を返す"""
    print("[DEBUG] check_inbox_once: 開始")

    if service is None:
        print("[DEBUG]   Drive service を初期化中...")
        service = get_drive_service()
    if openai_client is None:
        print("[DEBUG]   OpenAI client を初期化中...")
        openai_client = OpenAI()

    state = load_state()
    print(f"[DEBUG]   現在のstate: pending_count={state.get('pending_count')}")

    inbox_id = DRIVE_INBOX_FOLDER_ID
    print(f"[DEBUG]   inbox_id: {inbox_id}")

    print("[DEBUG]   processed フォルダを取得中...")
    processed_id = find_or_create_processed_folder(service)
    print(f"[DEBUG]   processed_id: {processed_id}")

    # サブフォルダも含めて再帰的にPDFを探索
    print("[DEBUG]   list_pdfs_recursive 開始...")
    pdf_files = list_pdfs_recursive(service, inbox_id)
    now = datetime.now(timezone.utc).isoformat()
    print(f"[DEBUG]   検出PDF数: {len(pdf_files)}")
    for pf in pdf_files:
        print(f"[DEBUG]     - {pf['name']} (id={pf['id'][:12]}..., hint={pf.get('folder_hint')})")

    # 検出直後に未処理数を書き込む（処理前の件数をUIに即反映）
    state["last_check"] = now
    state["pending_count"] = len(pdf_files)
    save_state(state)
    print(f"[DEBUG]   state保存完了: pending_count={len(pdf_files)}")

    if not pdf_files:
        print("[DEBUG]   PDF無し — 終了")
        return {"checked": now, "found": 0, "processed": 0, "pending": 0}

    # 各PDFを分類
    print("[DEBUG]   分類フェーズ開始...")
    classified = []
    for pf in pdf_files:
        folder_hint = pf.get("folder_hint")

        # フォルダ名やファイル名で判定できる場合はダウンロード不要
        quick_type = classify_pdf(openai_client, "", pf["name"], folder_hint)
        if quick_type != "unknown":
            print(f"[DEBUG]     {pf['name']} → {quick_type} (ヒント判定)")
            classified.append({**pf, "type": quick_type})
            continue

        # GPT-4o判定が必要 → ダウンロードしてテキスト抽出
        print(f"[DEBUG]     {pf['name']} → GPT-4oで判定中...")
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            download_pdf(service, pf["id"], tmp.name)
            tmp_path = tmp.name

        try:
            text = ata_extract.extract_text_from_pdf(tmp_path, max_pages=3)
            pdf_type = classify_pdf(openai_client, text, pf["name"], folder_hint)
        finally:
            os.unlink(tmp_path)

        print(f"[DEBUG]     {pf['name']} → {pdf_type} (GPT-4o)")
        classified.append({**pf, "type": pdf_type})

    # 分類結果サマリー
    drawings = [c for c in classified if c["type"] == "drawing"]
    estimates = [c for c in classified if c["type"] == "estimate"]
    unknowns = [c for c in classified if c["type"] == "unknown"]
    print(f"[DEBUG]   分類結果: drawing={len(drawings)}, estimate={len(estimates)}, unknown={len(unknowns)}")

    # ペアリング
    print("[DEBUG]   ペアリング開始...")
    pairs = group_files_by_project(classified)
    print(f"[DEBUG]   ペア数: {len(pairs)}")
    for key, pair in pairs.items():
        print(f"[DEBUG]     ペア '{key}':")
        print(f"[DEBUG]       図面:   {pair['drawing']['name']}")
        print(f"[DEBUG]       見積書: {pair['estimate']['name']}")

    processed_count = 0
    for key, pair in pairs.items():
        print(f"[DEBUG]   ペア '{key}' を処理中...")
        try:
            success = process_pair(service, openai_client, pair, processed_id)
            print(f"[DEBUG]     結果: {'成功' if success else '失敗'}")
        except Exception as e:
            print(f"[DEBUG]     例外発生: {type(e).__name__}: {e}")
            success = False

        if success:
            processed_count += 1
            state["processed"].append({
                "drawing": pair["drawing"]["name"],
                "estimate": pair["estimate"]["name"],
                "processed_at": now,
            })

    # 処理後の未処理数を再帰的に再カウントして更新
    print("[DEBUG]   処理後の残数を再カウント中...")
    remaining = list_pdfs_recursive(service, inbox_id)
    state["last_check"] = now
    state["pending_count"] = len(remaining)
    save_state(state)

    print(f"[DEBUG] check_inbox_once 完了: found={len(pdf_files)}, processed={processed_count}, pending={len(remaining)}")
    return {
        "checked": now,
        "found": len(pdf_files),
        "processed": processed_count,
        "pending": len(remaining),
    }


# ── バックグラウンドスレッド ──

def _watcher_loop():
    """バックグラウンドで5分ごとにinboxをチェックするループ"""
    logging.basicConfig(
        level=logging.INFO,
        format="[DriveWatcher %(asctime)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    logger.info("バックグラウンド監視を開始（間隔: %d秒）", DRIVE_POLL_INTERVAL)

    service = get_drive_service()
    openai_client = OpenAI()

    while True:
        try:
            result = check_inbox_once(service, openai_client)
            logger.info(
                "チェック完了 — 検出: %d, 処理: %d, 未処理: %d",
                result["found"], result["processed"], result["pending"],
            )
        except Exception as e:
            logger.error("チェック失敗: %s", e, exc_info=True)

        time.sleep(DRIVE_POLL_INTERVAL)


def start_watcher():
    """バックグラウンド監視スレッドを開始する（二重起動防止付き）"""
    global _watcher_thread
    with _watcher_lock:
        if _watcher_thread is not None and _watcher_thread.is_alive():
            return False  # 既に動作中
        _watcher_thread = threading.Thread(
            target=_watcher_loop, daemon=True, name="DriveWatcher"
        )
        _watcher_thread.start()
        return True


def is_watcher_running() -> bool:
    """監視スレッドが動作中かどうか"""
    with _watcher_lock:
        return _watcher_thread is not None and _watcher_thread.is_alive()


# ── CLI ──

def main():
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format="[%(asctime)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    parser = argparse.ArgumentParser(description="AtA Google Drive 自動監視")
    parser.add_argument("--daemon", "-d", action="store_true",
                        help="デーモンモードで5分ごとに監視（フォアグラウンド）")
    parser.add_argument("--once", action="store_true",
                        help="1回だけチェックして終了")
    args = parser.parse_args()

    if args.daemon:
        print("=" * 60)
        print("  AtA Drive Watcher — デーモンモード")
        print(f"  inbox フォルダID: {DRIVE_INBOX_FOLDER_ID}")
        print(f"  監視間隔: {DRIVE_POLL_INTERVAL} 秒")
        print("=" * 60)
        _watcher_loop()
    else:
        print(f"AtA_inbox ({DRIVE_INBOX_FOLDER_ID}) をチェック中...")
        result = check_inbox_once()
        print(f"\n完了:")
        print(f"  検出:   {result['found']} 件")
        print(f"  処理:   {result['processed']} 件")
        print(f"  未処理: {result['pending']} 件")


if __name__ == "__main__":
    main()
