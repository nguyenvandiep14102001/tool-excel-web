"""
tools/groupby_agg.py
────────────────────
Tool: 📈 GroupBy / Tổng Hợp
Nhóm dữ liệu và tính sum/count/mean/min/max/distinct.
"""

import streamlit as st
import pandas as pd

from modules.excel_utils import df_to_excel_bytes
from modules.ui_components import (
    show_single_preview, show_stats, excel_download_button, no_file_warning,
)

AGG_OPTIONS = {
    "Đếm (count)": "count",
    "Tổng (sum)": "sum",
    "Trung bình (mean)": "mean",
    "Min": "min",
    "Max": "max",
    "Distinct count": "nunique",
}


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    options = [f"Sheet 1 ({name1})"]
    if df2 is not None:
        options.append(f"Sheet 2 ({name2})")
    target = st.radio("Sheet nguồn", options, horizontal=True, key="gb_target")

    df = df1 if "Sheet 1" in target else df2
    show_single_preview(df, name1 if "Sheet 1" in target else name2)
    st.markdown("### 📈 GroupBy / Tổng Hợp")

    cols = list(df.columns)
    group_cols = st.multiselect("Nhóm theo cột", cols, key="gb_group")
    value_col = st.selectbox("Cột giá trị", cols, key="gb_value")
    agg_label = st.selectbox("Hàm tổng hợp", list(AGG_OPTIONS.keys()), key="gb_agg")

    if st.button("🚀 Tổng Hợp", use_container_width=True):
        if not group_cols:
            st.warning("Chọn ít nhất 1 cột để nhóm.")
            return
        _run(df, group_cols, value_col, AGG_OPTIONS[agg_label])


def _run(df: pd.DataFrame, group_cols: list, value_col: str, aggfunc: str) -> None:
    work = df.copy()
    if aggfunc != "count":
        work[value_col] = pd.to_numeric(work[value_col], errors="coerce")

    grouped = (
        work.groupby(group_cols, dropna=False)[value_col]
        .agg(aggfunc)
        .reset_index()
        .rename(columns={value_col: f"{value_col}_{aggfunc}"})
    )

    show_stats([
        {"num": str(len(df)), "lbl": "Dòng gốc", "color": None},
        {"num": str(len(grouped)), "lbl": "Nhóm", "color": "#4dbb77"},
        {"num": str(len(group_cols)), "lbl": "Cột group", "color": "#7eb8f7"},
    ])

    st.dataframe(grouped, use_container_width=True, height=380)
    excel_download_button(
        df_to_excel_bytes(grouped, "GroupBy"),
        filename="groupby_result.xlsx",
    )
