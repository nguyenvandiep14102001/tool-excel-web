"""
tools/filter_builder.py
───────────────────────
Tool: 🔽 Filter Builder
Lọc dòng theo điều kiện, xuất subset.
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

FILTER_OPS = {
    "Bằng": "eq",
    "Khác": "neq",
    "Chứa": "contains",
    "Không chứa": "not_contains",
    "Lớn hơn": "gt",
    "Nhỏ hơn": "lt",
    "Lớn hơn hoặc bằng": "gte",
    "Nhỏ hơn hoặc bằng": "lte",
    "Trống": "is_null",
    "Không trống": "not_null",
}


def _apply_filter(series: pd.Series, op: str, value: str) -> pd.Series:
    if op == "is_null":
        return series.isna() | (series.astype(str).str.strip() == "")
    if op == "not_null":
        return series.notna() & (series.astype(str).str.strip() != "")
    if op == "contains":
        return series.astype(str).str.contains(value, na=False, case=False, regex=False)
    if op == "not_contains":
        return ~series.astype(str).str.contains(value, na=False, case=False, regex=False)
    if op == "eq":
        return series.astype(str) == str(value)
    if op == "neq":
        return series.astype(str) != str(value)

    num = pd.to_numeric(series, errors="coerce")
    v = float(value)
    if op == "gt":
        return num > v
    if op == "lt":
        return num < v
    if op == "gte":
        return num >= v
    if op == "lte":
        return num <= v
    return pd.Series(True, index=series.index)


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    options = [f"Sheet 1 ({name1})"]
    if df2 is not None:
        options.append(f"Sheet 2 ({name2})")
    target = st.radio("Filter sheet nào?", options, horizontal=True, key="filter_target")

    df = df1 if "Sheet 1" in target else df2
    show_single_preview(df, name1 if "Sheet 1" in target else name2)
    st.markdown("### 🔽 Filter Builder")

    n_filters = st.number_input("Số điều kiện", 1, 8, 1, key="filter_n")
    combine = st.radio("Kết hợp điều kiện", ["AND (tất cả)", "OR (bất kỳ)"], horizontal=True)

    cols = list(df.columns)
    rules = []
    for i in range(int(n_filters)):
        c1, c2, c3 = st.columns([2, 2, 2])
        col = c1.selectbox(f"Cột #{i+1}", cols, key=f"flt_col_{i}")
        op_label = c2.selectbox("Điều kiện", list(FILTER_OPS.keys()), key=f"flt_op_{i}")
        val = c3.text_input("Giá trị", "", key=f"flt_val_{i}")
        rules.append({"col": col, "op": FILTER_OPS[op_label], "value": val})

    if st.button("🚀 Lọc Dữ Liệu", use_container_width=True):
        _run(df, rules, combine.startswith("AND"))


def _run(df: pd.DataFrame, rules: list, use_and: bool) -> None:
    masks = []
    for rule in rules:
        if rule["col"] not in df.columns:
            continue
        masks.append(_apply_filter(df[rule["col"]], rule["op"], rule["value"]))

    if not masks:
        st.warning("Không có điều kiện hợp lệ.")
        return

    final = masks[0]
    for m in masks[1:]:
        final = final & m if use_and else final | m

    filtered = df[final].reset_index(drop=True)
    removed = len(df) - len(filtered)

    show_stats([
        {"num": str(len(df)), "lbl": "Trước lọc", "color": None},
        {"num": str(len(filtered)), "lbl": "Sau lọc", "color": "#4dbb77"},
        {"num": str(removed), "lbl": "Đã loại", "color": "#ff6b6b"},
    ])

    st.dataframe(filtered, use_container_width=True, height=380)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Filtered"
    ws.append(list(filtered.columns))
    style_header(ws)
    for _, row in filtered.iterrows():
        ws.append(list(row))
    auto_col_width(ws)
    freeze_header(ws)
    excel_download_button(save_workbook_to_bytes(wb), filename="filtered_data.xlsx")
