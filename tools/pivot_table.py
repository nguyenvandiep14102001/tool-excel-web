"""
tools/pivot_table.py
────────────────────
Tool: 🧮 Pivot Table
Tạo bảng pivot từ dữ liệu dạng bảng.
"""

import streamlit as st
import pandas as pd

from modules.excel_utils import df_to_excel_bytes
from modules.ui_components import (
    show_single_preview, show_stats, excel_download_button, no_file_warning,
)

AGG_MAP = {
    "Sum": "sum",
    "Count": "count",
    "Mean": "mean",
    "Min": "min",
    "Max": "max",
}


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    options = [f"Sheet 1 ({name1})"]
    if df2 is not None:
        options.append(f"Sheet 2 ({name2})")
    target = st.radio("Sheet nguồn", options, horizontal=True, key="pivot_target")

    df = df1 if "Sheet 1" in target else df2
    show_single_preview(df, name1 if "Sheet 1" in target else name2)
    st.markdown("### 🧮 Pivot Table")

    cols = list(df.columns)
    c1, c2, c3 = st.columns(3)
    index_col = c1.selectbox("Rows (index)", cols, key="pivot_index")
    column_col = c2.selectbox("Columns", cols, index=min(1, len(cols) - 1), key="pivot_col")
    value_col = c3.selectbox("Values", cols, index=min(2, len(cols) - 1), key="pivot_val")
    agg_label = st.selectbox("Aggregation", list(AGG_MAP.keys()), key="pivot_agg")

    if st.button("🚀 Tạo Pivot", use_container_width=True):
        _run(df, index_col, column_col, value_col, AGG_MAP[agg_label])


def _run(df: pd.DataFrame, index_col: str, column_col: str, value_col: str, aggfunc: str) -> None:
    work = df.copy()
    if aggfunc != "count":
        work[value_col] = pd.to_numeric(work[value_col], errors="coerce")

    pivot = pd.pivot_table(
        work,
        index=index_col,
        columns=column_col,
        values=value_col,
        aggfunc=aggfunc,
        fill_value=0,
        margins=True,
        margins_name="Tổng",
    )

    show_stats([
        {"num": str(pivot.shape[0]), "lbl": "Số dòng pivot", "color": None},
        {"num": str(pivot.shape[1]), "lbl": "Số cột pivot", "color": None},
        {"num": aggfunc, "lbl": "Aggregation", "color": "#7eb8f7"},
    ])

    st.dataframe(pivot, use_container_width=True, height=380)
    flat = pivot.reset_index()
    excel_download_button(
        df_to_excel_bytes(flat, "Pivot"),
        filename="pivot_result.xlsx",
    )
