"""
tools/clean_blank.py
────────────────────
Tool: 🧽 Dọn Dữ Liệu Trống
Xoá dòng/cột trống hoặc toàn giá trị rỗng.
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


def _is_blank(val) -> bool:
    if pd.isna(val):
        return True
    return str(val).strip() == ""


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    options = [f"Sheet 1 ({name1})"]
    if df2 is not None:
        options.append(f"Sheet 2 ({name2})")
    target = st.radio("Dọn sheet nào?", options, horizontal=True, key="clean_target")

    df = df1 if "Sheet 1" in target else df2
    show_single_preview(df, name1 if "Sheet 1" in target else name2)
    st.markdown("### 🧽 Dọn Dữ Liệu Trống")

    drop_empty_rows = st.checkbox("Xoá dòng trống hoàn toàn", True, key="clean_rows")
    row_mode = st.radio(
        "Hoặc xoá dòng có ô trống",
        ["Không", "Bất kỳ ô trống", "Tất cả ô trống"],
        horizontal=True,
        key="clean_row_mode",
    )
    drop_empty_cols = st.checkbox("Xoá cột trống hoàn toàn", True, key="clean_cols")
    trim_ws = st.checkbox("Trim khoảng trắng toàn bảng", True, key="clean_trim")

    if st.button("🚀 Dọn Dữ Liệu", use_container_width=True):
        _run(df, drop_empty_rows, row_mode, drop_empty_cols, trim_ws)


def _run(df, drop_empty_rows, row_mode, drop_empty_cols, trim_ws):
    result = df.copy().astype("object")
    rows_before, cols_before = result.shape

    if trim_ws:
        for col in result.columns:
            mask = result[col].notna()
            result.loc[mask, col] = result.loc[mask, col].astype(str).str.strip()

    if drop_empty_rows:
        keep = ~result.apply(lambda r: all(_is_blank(v) for v in r), axis=1)
        result = result.loc[keep]

    if row_mode == "Bất kỳ ô trống":
        keep = ~result.apply(lambda r: any(_is_blank(v) for v in r), axis=1)
        result = result.loc[keep]
    elif row_mode == "Tất cả ô trống":
        keep = ~result.apply(lambda r: all(_is_blank(v) for v in r), axis=1)
        result = result.loc[keep]

    if drop_empty_cols:
        blank_cols = [c for c in result.columns if all(_is_blank(v) for v in result[c])]
        result = result.drop(columns=blank_cols)

    result = result.reset_index(drop=True)

    show_stats([
        {"num": f"{rows_before}→{len(result)}", "lbl": "Dòng", "color": None},
        {"num": f"{cols_before}→{result.shape[1]}", "lbl": "Cột", "color": None},
        {"num": str(rows_before - len(result)), "lbl": "Dòng đã xoá", "color": "#ff6b6b"},
        {"num": str(cols_before - result.shape[1]), "lbl": "Cột đã xoá", "color": "#ffa94d"},
    ])

    st.dataframe(result, use_container_width=True, height=380)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cleaned"
    ws.append(list(result.columns))
    style_header(ws)
    for _, row in result.iterrows():
        ws.append(list(row))
    auto_col_width(ws)
    freeze_header(ws)
    excel_download_button(save_workbook_to_bytes(wb), filename="cleaned_data.xlsx")
