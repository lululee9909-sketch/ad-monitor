import os
import json
import sys
import requests
from datetime import datetime

import gspread
from oauth2client.service_account import ServiceAccountCredentials


def get_env_var(name: str) -> str:
    """安全讀取環境變數，沒有就直接報錯結束。"""
    value = os.environ.get(name)
    if not value:
        print(f"[ERROR] Environment variable '{name}' is not set.", file=sys.stderr)
        sys.exit(1)
    return value


def get_gsheet_client():
    """用環境變數中的 JSON 字串認證 Google Sheets（不讀任何本機檔案）。"""
    credentials_json_str = get_env_var("GCP_CREDENTIALS_JSON")

    try:
        credentials_dict = json.loads(credentials_json_str)
    except json.JSONDecodeError as e:
        print("[ERROR] GCP_CREDENTIALS_JSON is not a valid JSON string.", file=sys.stderr)
        print(e, file=sys.stderr)
        sys.exit(1)

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials = ServiceAccountCredentials.from_json_keyfile_dict(
        credentials_dict, scopes=scope
    )
    client = gspread.authorize(credentials)
    return client


def get_or_create_worksheet(sh, title: str, rows: int = 1000, cols: int = 10):
    """取得指定工作表，若不存在則建立。"""
    try:
        ws = sh.worksheet(title)
        return ws
    except gspread.WorksheetNotFound:
        print(f"[INFO] Worksheet '{title}' not found. Creating a new one...")
        ws = sh.add_worksheet(title=title, rows=str(rows), cols=str(cols))
        return ws


def read_config(config_ws):
    """從 Config 分頁讀取 Keywords (A 欄) 與 Competitors (B 欄)。"""
    # 第一列預設是標題，從第二列開始讀
    keywords_col = config_ws.col_values(1)[1:]
    competitors_col = config_ws.col_values(2)[1:]

    keywords = [k.strip() for k in keywords_col if k.strip()]
    competitors = [c.strip().lower() for c in competitors_col if c.strip()]

    if not keywords:
        print("[WARN] No keywords found in Config! (Column A, starting from A2)")
    if not competitors:
        print("[WARN] No competitors found in Config! (Column B, starting from B2)")

    return keywords, competitors


def classify_ad(link: str, competitors: list) -> str:
    """
    判斷是否為競品：
    link 包含任一 competitor 字串 → "Competitor"，否則 "Other"
    """
    if not link:
        return "Other"

    link_lower = link.lower()
    for comp in competitors:
        if comp in link_lower:
            return "Competitor"
    return "Other"


def fetch_serper_results(api_key: str, keyword: str) -> dict:
    """呼叫 Serper.dev API，取得搜尋結果，並印出 Debug 資訊。"""
    url = "https://google.serper.dev/search"
    headers = {
        "X-API-KEY": api_key,
        "Content-Type": "application/json",
    }
    payload = {
        "q": keyword,
        "gl": "tw",
        "hl": "zh-tw",
    }

    resp = requests.post(url, headers=headers, json=payload, timeout=30)
    resp.raise_for_status()
    result = resp.json()

    # ==== 這一段是 Debug：看看 Serper 實際回什麼 ====
    try:
        print(f"[DEBUG] Keys in response for '{keyword}': {list(result.keys())}")
        snippet = json.dumps(result, ensure_ascii=False)[:3000]
        print(f"[DEBUG] Raw Serper response snippet for '{keyword}':\n{snippet}")
    except Exception as e:
        print(
            f"[DEBUG] Failed to pretty-print response for '{keyword}': {e}",
            file=sys.stderr,
        )
    # =================================================

    return result


def ensure_headers(ws, headers: list):
    """如果工作表第一列是空的，就自動建立標題列。"""
    first_row = ws.row_values(1)
    if not first_row:
        ws.insert_row(headers, 1)


def append_ads_to_data_sheet(
    data_ws,
    keyword: str,
    competitors: list,
    ads: list,
    today_str: str,
):
    """
    將廣告資料寫入 Data 分頁：
    欄位: Date, Keyword, Position, Status, Title, Description, Link
    """
    ensure_headers(
        data_ws,
        ["Date", "Keyword", "Position", "Status", "Title", "Description", "Link"],
    )

    rows = []
    for idx, ad in enumerate(ads, start=1):
        title = ad.get("title") or ""
        description = ad.get("description") or ad.get("snippet") or ""
        link = ad.get("link") or ad.get("url") or ""

        status = classify_ad(link, competitors)

        row = [
            today_str,
            keyword,
            idx,  # Position：就用當次結果的順位
            status,
            title,
            description,
            link,
        ]
        rows.append(row)

    if not rows:
        return

    # 優先用 append_rows，若版本不支援就退回一列一列 append_row
    try:
        data_ws.append_rows(rows, value_input_option="RAW")
    except AttributeError:
        for row in rows:
            data_ws.append_row(row, value_input_option="RAW")


def append_related_searches_to_sheet(
    kw_ideas_ws,
    keyword: str,
    related_searches: list,
    today_str: str,
):
    """
    將相關搜尋寫入 Keyword_Ideas 分頁：
    欄位: Keyword, Date
    這裡的 Keyword 就是 relatedSearches 裡的 query。
    """
    ensure_headers(
        kw_ideas_ws,
        ["Keyword", "Date"],
    )

    rows = []
    for item in rela
