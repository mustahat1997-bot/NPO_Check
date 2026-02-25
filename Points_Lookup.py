import os
import pandas as pd
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Points_Feb_2026.xlsx")
SHEET_NAME = "Points&Repeaters&Province"

_cached_df = None

def normalize_text(text):
    return re.sub(r"\s+", "", str(text)).lower().strip()

def load_points_excel():
    global _cached_df
    if _cached_df is None:
        if not os.path.exists(EXCEL_PATH):
            raise FileNotFoundError(f"{EXCEL_PATH} not found in repo!")

        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
        df.columns = df.columns.str.strip()
        df["affiliate_normalized"] = df["Affiliate_Name"].apply(normalize_text)
        _cached_df = df

    return _cached_df

def get_repeater_and_province_from_excel(point_name):
    df = load_points_excel()
    search_value = normalize_text(point_name)
    match = df[df["affiliate_normalized"] == search_value]

    if not match.empty:
        row = match.iloc[0]
        return str(row["Repeater Class"]).strip(), str(row["Province"]).strip()

    return "No Repeater", "No Province"