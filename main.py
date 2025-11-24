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


def classify_lin_
