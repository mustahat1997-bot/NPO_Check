import os
import sqlite3
import pandas as pd
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Q_Actions.db")
EXCEL_PATH = os.path.join(BASE_DIR, "Updated changes queues Repeaters Data V148  Final).xlsx")
TARGET_SHEET = "All Repeaters & Affiliates"
TABLE_NAME = "repeater_actions"

def normalize_text(text):
    return re.sub(r"\s+", "", str(text)).lower().strip()

def excel_to_dataframe():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"{EXCEL_PATH} not found in repo!")
    df = pd.read_excel(EXCEL_PATH, sheet_name=TARGET_SHEET)
    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        "RepeaterName / Affiliates Name": "name",
        "Site code": "site_code",
        "Q Action": "q_action",
        "Repeater Action": "repeater_action"
    })
    required_columns = ["name", "site_code", "q_action", "repeater_action"]
    df = df[required_columns]
    df["name"] = df["name"].astype(str).str.strip()
    df["site_code"] = df["site_code"].astype(str).str.strip()
    df["q_action"] = df["q_action"].fillna("No Action")
    df["repeater_action"] = df["repeater_action"].fillna("No Action")
    if df.empty:
        raise ValueError("DataFrame is empty")
    return df, os.path.basename(EXCEL_PATH)

def save_to_sqlite(df):
    conn = sqlite3.connect(DB_PATH)
    df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)
    conn.close()

def connect_db():
    return sqlite3.connect(DB_PATH)

def get_actions_from_db(cursor, value):
    normalized_value = normalize_text(value)
    cursor.execute("""
        SELECT q_action, repeater_action
        FROM repeater_actions
        WHERE REPLACE(LOWER(name), ' ', '') = ?
        OR REPLACE(LOWER(site_code), ' ', '') = ?
        LIMIT 1
    """, (normalized_value, normalized_value))
    row = cursor.fetchone()
    return row if row else ("No Action", "No Action")

def get_q_action_by_site_code(cursor, site_code):
    normalized_value = normalize_text(site_code)
    cursor.execute("""
        SELECT q_action
        FROM repeater_actions
        WHERE REPLACE(LOWER(site_code), ' ', '') = ?
        LIMIT 1
    """, (normalized_value,))
    row = cursor.fetchone()
    return row[0] if row else "No Action"

def get_action_type(rule):
    if rule.startswith("(+"):
        return "add"
    if rule.startswith("(-"):
        return "decrease"
    return "add"
