"""
tools/data_profile.py
─────────────────────
Tool: 📊 Data Profile
Thống kê chất lượng dữ liệu theo từng cột.
"""

import streamlit as st
import pandas as pd

from modules.excel_utils import df_to_excel_bytes
from modules.ui_components import (
    show_single_preview, show_stats, excel_download_button, no_file_warning,
)


def _infer_kind(series: pd.Series) -> str:
    if pd.api.types.is_datetime64_any_dtype(series):
        return "date"
    if pd.api.types.is_numeric_dtype(series):
        return "number"
    sample = series.dropna().astype(str).head(20)
    if sample.empty:
        return "empty"
    if sample.str.match(r"^\d{4}-\d{2}-\d{2}").mean() > 0.5:
        return "date-like"
    if sample.str.match(r"^[\d,.\s]+$").mean() > 0.5:
        return "number-like"
    if sample.str.contains("@").mean() > 0.5:
        return "email-like"
    return "text"


def _profile_column(series: pd.Series) -> dict:
    n = len(series)
    null_cnt = int(series.isna().sum())
    non_null = series.dropna()
    distinct = int(non_null.nunique()) if len(non_null) else 0
    dup_cnt = int(non_null.duplicated().sum()) if len(non_null) else 0

    row = {
        "Cột": series.name,
        "Kiểu gợi ý": _infer_kind(series),
        "Tổng dòng": n,
        "Null": null_cnt,
        "% Null": round(100 * null_cnt / n, 1) if n else 0,
        "Distinct": distinct,
        "Trùng lặp": dup_cnt,
        "Ví dụ": "",
        "Min": "",
        "Max": "",
    }

    if len(non_null):
        samples = non_null.astype(str).unique()[:3]
        row["Ví dụ"] = " | ".join(samples)

    if pd.api.types.is_numeric_dtype(non_null):
        row["Min"] = non_null.min()
        row["Max"] = non_null.max()
    elif _infer_kind(series) in ("date", "date-like"):
        parsed = pd.to_datetime(non_null.astype(str), errors="coerce", dayfirst=True)
        if parsed.notna().any():
            row["Min"] = parsed.min().strftime("%Y-%m-%d")
            row["Max"] = parsed.max().strftime("%Y-%m-%d")

    return row


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    options = [f"Sheet 1 ({name1})"]
    if df2 is not None:
        options.append(f"Sheet 2 ({name2})")
    target = st.radio("Profile sheet nào?", options, horizontal=True, key="profile_target")

    df = df1 if "Sheet 1" in target else df2
    show_single_preview(df, name1 if "Sheet 1" in target else name2)
    st.markdown("### 📊 Data Profile")

    if st.button("🚀 Chạy Profile", use_container_width=True):
        _run(df)


def _run(df: pd.DataFrame) -> None:
    rows = [_profile_column(df[c]) for c in df.columns]
    profile_df = pd.DataFrame(rows)

    n_rows, n_cols = len(df), len(df.columns)
    null_cells = int(df.isna().sum().sum())
    dup_rows = int(df.duplicated().sum())

    show_stats([
        {"num": str(n_rows), "lbl": "Dòng", "color": None},
        {"num": str(n_cols), "lbl": "Cột", "color": None},
        {"num": str(null_cells), "lbl": "Ô trống", "color": "#ffa94d"},
        {"num": str(dup_rows), "lbl": "Dòng trùng", "color": "#ff6b6b"},
    ])

    st.dataframe(profile_df, use_container_width=True, height=420)
    excel_download_button(
        df_to_excel_bytes(profile_df, "Data Profile"),
        filename="data_profile.xlsx",
        label="⬇️ Tải báo cáo Profile",
    )
