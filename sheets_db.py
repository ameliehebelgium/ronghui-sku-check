"""
Google Sheets 数据库连接层 —— RongHui 专属

本系统与 EU Risk App、UK SKU Tracker 完全独立：独立的 Google Spreadsheet
文件（不同的 SHEET_ID，在 Streamlit secrets 中配置为 RONGHUI_SHEET_ID），
复用同一个 Google service account 凭据即可，无需重新走 Google Cloud 授权流程。

两个核心 Sheet：
    1. RongHui_SKU_Master  - 主数据库（当前生效的 SKU / 品名 / HS / 首次录入时间）
    2. RongHui_Change_Log  - 变更日志（每次冲突时人工选择"保留"或"更新"的记录）
"""
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
import streamlit as st

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

MASTER_SHEET_NAME = "RongHui_SKU_Master"
LOG_SHEET_NAME = "RongHui_Change_Log"

MASTER_COLUMNS = ["sku", "description", "hs_code", "tax_rate", "first_entry_time", "last_update_time", "sources"]
LOG_COLUMNS = [
    "timestamp", "sku", "old_description", "old_hs_code", "old_tax_rate",
    "new_description", "new_hs_code", "new_tax_rate", "action", "operator",
]


@st.cache_resource
def _get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def _get_spreadsheet():
    client = _get_client()
    return client.open_by_key(st.secrets["RONGHUI_SHEET_ID"])


def _get_or_create_worksheet(sheet_name, columns):
    ss = _get_spreadsheet()
    try:
        ws = ss.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=sheet_name, rows=1000, cols=len(columns) + 2)
        ws.append_row(columns)
    return ws


def load_master_db() -> pd.DataFrame:
    ws = _get_or_create_worksheet(MASTER_SHEET_NAME, MASTER_COLUMNS)
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=MASTER_COLUMNS)
    df = pd.DataFrame(records)
    for col in MASTER_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df["sku"] = df["sku"].astype(str).str.strip()
    df["description"] = df["description"].astype(str).str.strip()
    df["hs_code"] = df["hs_code"].astype(str).str.strip()
    df["tax_rate"] = df["tax_rate"].astype(str).str.strip()
    return df


def append_new_skus(new_rows: list[dict], operator: str = "system"):
    """new_rows: [{sku, description, hs_code, tax_rate, sources}, ...]"""
    if not new_rows:
        return
    ws = _get_or_create_worksheet(MASTER_SHEET_NAME, MASTER_COLUMNS)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows_to_append = []
    for r in new_rows:
        rows_to_append.append([
            r["sku"], r["description"], r["hs_code"], r.get("tax_rate", ""), now, now, r.get("sources", "")
        ])
    ws.append_rows(rows_to_append)
    _append_log_rows([
        {
            "timestamp": now, "sku": r["sku"], "old_description": "", "old_hs_code": "", "old_tax_rate": "",
            "new_description": r["description"], "new_hs_code": r["hs_code"], "new_tax_rate": r.get("tax_rate", ""),
            "action": "NEW_ENTRY", "operator": operator,
        }
        for r in new_rows
    ])


def apply_conflict_decisions(decisions: list[dict], operator: str = "user"):
    """
    decisions: [{sku, action: 'KEEP'|'UPDATE', old_description, old_hs_code, old_tax_rate,
                 new_description, new_hs_code, new_tax_rate}, ...]
    action == 'UPDATE' 时才真正写回数据库，'KEEP' 只记录日志不改数据。
    """
    if not decisions:
        return
    ws = _get_or_create_worksheet(MASTER_SHEET_NAME, MASTER_COLUMNS)
    all_values = ws.get_all_values()
    header = all_values[0] if all_values else MASTER_COLUMNS
    sku_col_idx = header.index("sku") if "sku" in header else 0
    desc_col_idx = header.index("description") if "description" in header else 1
    hs_col_idx = header.index("hs_code") if "hs_code" in header else 2
    tax_col_idx = header.index("tax_rate") if "tax_rate" in header else None
    update_col_idx = header.index("last_update_time") if "last_update_time" in header else 4

    sku_to_row = {row[sku_col_idx].strip(): idx + 2 for idx, row in enumerate(all_values[1:])}

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_rows = []
    for d in decisions:
        sku = d["sku"]
        log_rows.append({
            "timestamp": now, "sku": sku,
            "old_description": d["old_description"], "old_hs_code": d["old_hs_code"],
            "old_tax_rate": d.get("old_tax_rate", ""),
            "new_description": d["new_description"], "new_hs_code": d["new_hs_code"],
            "new_tax_rate": d.get("new_tax_rate", ""),
            "action": d["action"], "operator": operator,
        })
        if d["action"] == "UPDATE" and sku in sku_to_row:
            row_idx = sku_to_row[sku]
            ws.update_cell(row_idx, desc_col_idx + 1, d["new_description"])
            ws.update_cell(row_idx, hs_col_idx + 1, d["new_hs_code"])
            if tax_col_idx is not None:
                ws.update_cell(row_idx, tax_col_idx + 1, d.get("new_tax_rate", ""))
            ws.update_cell(row_idx, update_col_idx + 1, now)

    _append_log_rows(log_rows)


def _append_log_rows(rows: list[dict]):
    ws = _get_or_create_worksheet(LOG_SHEET_NAME, LOG_COLUMNS)
    ws.append_rows([[r.get(c, "") for c in LOG_COLUMNS] for r in rows])


def load_change_log() -> pd.DataFrame:
    ws = _get_or_create_worksheet(LOG_SHEET_NAME, LOG_COLUMNS)
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=LOG_COLUMNS)
    return pd.DataFrame(records)
