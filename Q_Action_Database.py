import os
import sqlite3
import pandas as pd
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Q_Actions.db")

TARGET_SHEET = "All Repeaters & Affiliates"
TABLE_NAME = "repeater_actions"


def normalize_text(text):
    return re.sub(r"\s+", "", str(text)).lower().strip()


# ✅ إنشاء الجدول إذا ما موجود
def ensure_table_exists():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
            name TEXT,
            site_code TEXT,
            q_action TEXT,
            repeater_action TEXT
        )
    """)

    conn.commit()
    conn.close()


def excel_to_dataframe(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET)

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

    return df


def save_to_sqlite(df):
    ensure_table_exists()
    conn = sqlite3.connect(DB_PATH)
    df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)
    conn.close()


def connect_db():
    ensure_table_exists()
    return sqlite3.connect(DB_PATH)


def get_actions_from_db(cursor, value):
    normalized_value = normalize_text(value)

    cursor.execute(f"""
        SELECT q_action, repeater_action
        FROM {TABLE_NAME}
        WHERE REPLACE(LOWER(name), ' ', '') = ?
        LIMIT 1
    """, (normalized_value,))

    row = cursor.fetchone()

    if row:
        q_action = row[0] if row[0] else "No Action"
        repeater_action = row[1] if row[1] else "No Action"
        return q_action, repeater_action

    return "No Action", "No Action"


def get_q_action_by_site_code(cursor, site_code):
    normalized_value = normalize_text(site_code)

    cursor.execute(f"""
        SELECT q_action
        FROM {TABLE_NAME}
        WHERE REPLACE(LOWER(site_code), ' ', '') = ?
        LIMIT 1
    """, (normalized_value,))

    row = cursor.fetchone()

    if row and row[0]:
        return row[0]

    return "No Action"


# 🔥 الدالة المعدلة لإرجاع Q Action Repeater النهائي
def get_final_q_action_repeater(q_action, r_action, q_action_repeater):
    if (
        q_action == "No Action" and
        r_action == "No Action" and
        q_action_repeater == "No Action"
    ):
        return "Default => (+80%) International & Local CDN 100%"

    return q_action_repeater


def get_action_type(rule):
    if rule.startswith("(+"):
        return "add"
    if rule.startswith("(-"):
        return "decrease"
    return "add"