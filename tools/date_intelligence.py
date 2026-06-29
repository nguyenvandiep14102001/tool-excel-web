"""
tools/date_intelligence.py
──────────────────────────
Tool: 📅 Date Intelligence
Parse cột ngày và thêm year/quarter/month/week/weekday.
"""

import streamlit as st
import pandas as pd
import openpyxl

from modules.excel_utils import (
    style_header, auto_col_width, freeze_header, save_workbook_to_bytes,
)
from modules.ui_components import (
    show_single_preview, show_stats, excel_download_button, no_file_warning,
)


def _parse_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit="D", origin="1899-12-30", errors="coerce")
    texts = series.astype(str)
    mask = texts.str.contains(r"[\/\-.]", regex=True)
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    if mask.any():
        parsed.loc[mask] = pd.to_datetime(texts[mask], errors="coerce", dayfirst=True)
    remain = ~mask & series.notna()
    if remain.any():
        parsed.loc[remain] = pd.to_datetime(texts[remain], errors="coerce", dayfirst=True)
    return parsed


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    options = [f"Sheet 1 ({name1})"]
    if df2 is not None:
        options.append(f"Sheet 2 ({name2})")
    target = st.radio("Sheet nguồn", options, horizontal=True, key="dateintel_target")

    df = df1 if "Sheet 1" in target else df2
    show_single_preview(df, name1 if "Sheet 1" in target else name2)
    st.markdown("### 📅 Date Intelligence")

    cols = list(df.columns)
    date_cols = st.multiselect("Cột ngày cần xử lý", cols, key="dateintel_cols")

    c1, c2, c3 = st.columns(3)
    add_year = c1.checkbox("Year", True, key="di_year")
    add_quarter = c2.checkbox("Quarter", True, key="di_q")
    add_month = c3.checkbox("Month", True, key="di_month")
    c4, c5, c6 = st.columns(3)
    add_week = c4.checkbox("Week", False, key="di_week")
    add_weekday = c5.checkbox("Weekday", True, key="di_wd")
    add_weekend = c6.checkbox("Is weekend", True, key="di_we")
    fmt_out = st.text_input("Format chuẩn hoá ngày", "%Y-%m-%d", key="di_fmt")

    if st.button("🚀 Thêm Cột Ngày", use_container_width=True):
        if not date_cols:
            st.warning("Chọn ít nhất 1 cột ngày.")
            return
        _run(df, date_cols, add_year, add_quarter, add_month, add_week,
             add_weekday, add_weekend, fmt_out)


def _run(df, date_cols, add_year, add_quarter, add_month, add_week,
         add_weekday, add_weekend, fmt_out):
    result = df.copy().astype("object")
    converted = 0

    for col in date_cols:
        parsed = _parse_series(result[col])
        ok = parsed.notna()
        converted += int(ok.sum())

        result[f"{col}_date"] = parsed.dt.strftime(fmt_out).where(ok, result[col])
        if add_year:
            result[f"{col}_year"] = parsed.dt.year.where(ok, "")
        if add_quarter:
            result[f"{col}_quarter"] = parsed.dt.quarter.where(ok, "")
        if add_month:
            result[f"{col}_month"] = parsed.dt.month.where(ok, "")
        if add_week:
            iso_week = parsed.dt.isocalendar().week
            result[f"{col}_week"] = iso_week.where(ok, pd.NA)
        if add_weekday:
            result[f"{col}_weekday"] = parsed.dt.day_name().where(ok, "")
        if add_weekend:
            result[f"{col}_is_weekend"] = parsed.dt.weekday.isin([5, 6]).where(ok, "")

    show_stats([
        {"num": str(len(df)), "lbl": "Dòng", "color": None},
        {"num": str(len(date_cols)), "lbl": "Cột ngày", "color": None},
        {"num": str(converted), "lbl": "Ô parse OK", "color": "#4dbb77"},
    ])

    st.dataframe(result, use_container_width=True, height=380)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Date Intelligence"
    ws.append(list(result.columns))
    style_header(ws)
    for _, row in result.iterrows():
        ws.append(list(row))
    auto_col_width(ws)
    freeze_header(ws)
    excel_download_button(save_workbook_to_bytes(wb), filename="date_intelligence.xlsx")
