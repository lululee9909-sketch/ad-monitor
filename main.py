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
    keywords_col = config_ws.col_values(1)[1:]  # 跳過標題列
    competitors_col = config_ws.col_values(2)[1:]

    keywords = [k.strip() for k in keywords_col if k.strip()]
    competitors = [c.strip().lower() for c in competitors_col if c.strip()]

    if not keywords:
        print("[WARN] No keywords found in Config! (Column A, starting from A2)")
    if not competitors:
        print("[WARN] No competitors found in Config! (Column B, starting from B2)")

    return keywords, competitors


def classify_link(link: str, competitors: list) -> str:
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


def fetch_serpapi_results(api_key: str, keyword: str) -> dict:
    """
    呼叫 SerpApi 的 Google Search API，取得 SERP，
    用來抓自然結果 (organic_results) 與相關搜尋 (related_searches)。
    """
    url = "https://serpapi.com/search.json"
    params = {
        "api_key": api_key,
        "engine": "google",
        "q": keyword,
        "location": "Taipei, Taiwan",
        # 如果之後覺得語言怪怪的，可再加上 "hl": "zh-TW"
    }

    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()
    result = resp.json()

    # Debug：看一下這次回傳有哪些 key（方便排錯）
    try:
        print(f"[DEBUG] Keys in SerpApi response for '{keyword}': {list(result.keys())}")
    except Exception as e:
        print(f"[DEBUG] Failed to list keys for '{keyword}': {e}", file=sys.stderr)

    return result


def ensure_headers(ws, headers: list):
    """如果工作表第一列是空的，就自動建立標題列。"""
    first_row = ws.row_values(1)
    if not first_row:
        ws.insert_row(headers, 1)


def append_results_to_data_sheet(
    data_ws,
    keyword: str,
    competitors: list,
    organic_results: list,
    today_str: str,
):
    """
    將自然搜尋結果寫入 Data 分頁：
    欄位: Date, Keyword, Position, Status, Title, Description, Link
    """
    ensure_headers(
        data_ws,
        ["Date", "Keyword", "Position", "Status", "Title", "Description", "Link"],
    )

    rows = []
    for idx, item in enumerate(organic_results, start=1):
        # SerpApi 通常有 "position" 或 "rank"，沒有就用 idx
        position = item.get("position") or item.get("rank") or idx
        title = item.get("title") or ""
        description = item.get("snippet") or item.get("description") or ""
        link = item.get("link") or item.get("url") or ""

        status = classify_link(link, competitors)

        row = [
            today_str,
            keyword,
            position,
            status,
            title,
            description,
            link,
        ]
        rows.append(row)

    if not rows:
        return

    try:
        data_ws.append_rows(rows, value_input_option="RAW")
    except AttributeError:
        for row in rows:
            data_ws.append_row(row, value_input_option="RAW")


def append_related_searches_to_sheet(
    kw_ideas_ws,
    seed_keyword: str,
    related_searches: list,
    today_str: str,
):
    """
    將相關搜尋寫入 Keyword_Ideas 分頁：
    欄位: SeedKeyword, RelatedKeyword, Date
    SeedKeyword = 原本搜尋的關鍵字
    RelatedKeyword = SerpApi 回傳的相關搜尋 query
    """
    ensure_headers(
        kw_ideas_ws,
        ["SeedKeyword", "RelatedKeyword", "Date"],
    )

    rows = []
    for item in related_searches:
        query = item.get("query") or item.get("text")
        if not query:
            continue
        rows.append([seed_keyword, query, today_str])

    if not rows:
        return

    try:
        kw_ideas_ws.append_rows(rows, value_input_option="RAW")
    except AttributeError:
        for row in rows:
            kw_ideas_ws.append_row(row, value_input_option="RAW")


def main():
    print("[INFO] Starting Organic SERP Monitor (SerpApi)...")

    # 1. 讀取必要環境變數
    serpapi_key = get_env_var("SERPAPI_API_KEY")
    sheet_id = get_env_var("SHEET_ID")

    # 2. 連線 Google Sheets
    client = get_gsheet_client()
    sh = client.open_by_key(sheet_id)

    # 3. 取得 Config / Data / Keyword_Ideas 分頁
    try:
        config_ws = sh.worksheet("Config")
    except gspread.WorksheetNotFound:
        print(
            "[ERROR] Worksheet 'Config' not found. "
            "Please create it and set A1='Keywords', B1='Competitors'.",
            file=sys.stderr,
        )
        sys.exit(1)

    data_ws = get_or_create_worksheet(sh, "Data")
    kw_ideas_ws = get_or_create_worksheet(sh, "Keyword_Ideas")

    # 4. 讀取設定
    keywords, competitors = read_config(config_ws)

    if not keywords:
        print("[ERROR] No keywords to process. Exiting.")
        sys.exit(0)

    today_str = datetime.utcnow().strftime("%Y-%m-%d")

    # 5. 逐一 Keyword 呼叫 SerpApi，寫入自然結果 + 相關搜尋
    for keyword in keywords:
        print(f"[INFO] Processing keyword (organic): {keyword}")

        try:
            result = fetch_serpapi_results(serpapi_key, keyword)
        except requests.HTTPError as e:
            print(
                f"[ERROR] SerpApi request failed for keyword '{keyword}': {e}",
                file=sys.stderr,
            )
            continue
        except Exception as e:
            print(
                f"[ERROR] Unexpected error while fetching SerpApi results for '{keyword}': {e}",
                file=sys.stderr,
            )
            continue

        organic_results = (
            result.get("organic_results")
            or result.get("organic")
            or []
        )
        related_searches = (
            result.get("related_searches")
            or result.get("relatedSearches")
            or []
        )

        if organic_results:
            append_results_to_data_sheet(
                data_ws,
                keyword,
                competitors,
                organic_results,
                today_str,
            )
            print(
                f"[INFO] Appended {len(organic_results)} organic results "
                f"for keyword '{keyword}'."
            )
        else:
            print(f"[INFO] No organic results found for keyword '{keyword}'.")

        if related_searches:
            append_related_searches_to_sheet(
                kw_ideas_ws,
                keyword,
                related_searches,
                today_str,
            )
            print(
                f"[INFO] Appended {len(related_searches)} related searches "
                f"for keyword '{keyword}'."
            )
        else:
            print(f"[INFO] No related searches found for keyword '{keyword}'.")

    print("[INFO] Organic SERP Monitor script finished.")


if __name__ == "__main__":
    main()
