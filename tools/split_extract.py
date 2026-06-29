"""
tools/split_extract.py
──────────────────────
Tool: ✂️ Tách / Trích Cột
Tách cột theo delimiter hoặc regex extract.
"""

import streamlit as st
import pandas as pd
import openpyxl

from modules.excel_utils import (
    style_header, auto_col_width, freeze_header, save_workbook_to_bytes,
)
from modules.ui_components import (
    show_single_preview, excel_download_button, no_file_warning,
)


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    options = [f"Sheet 1 ({name1})"]
    if df2 is not None:
        options.append(f"Sheet 2 ({name2})")
    target = st.radio("Sheet nguồn", options, horizontal=True, key="split_target")

    df = df1 if "Sheet 1" in target else df2
    show_single_preview(df, name1 if "Sheet 1" in target else name2)
    st.markdown("### ✂️ Tách / Trích Cột")

    mode = st.radio("Chế độ", ["Tách theo delimiter", "Trích bằng Regex"], horizontal=True)
    cols = list(df.columns)
    src_col = st.selectbox("Cột nguồn", cols, key="split_src")

    if mode == "Tách theo delimiter":
        delim = st.text_input("Delimiter", ",", key="split_delim")
        n_parts = st.number_input("Số cột tạo ra", 2, 10, 2, key="split_n")
        prefix = st.text_input("Tiền tố tên cột mới", f"{src_col}_", key="split_prefix")
    else:
        pattern = st.text_input("Regex (group capture)", r"(\d+)", key="split_regex")
        new_name = st.text_input("Tên cột kết quả", f"{src_col}_extract", key="split_newname")

    if st.button("🚀 Thực hiện", use_container_width=True):
        if mode == "Tách theo delimiter":
            _run_split(df, src_col, delim, int(n_parts), prefix)
        else:
            _run_extract(df, src_col, pattern, new_name)


def _run_split(df: pd.DataFrame, src_col: str, delim: str, n_parts: int, prefix: str) -> None:
    result = df.copy().astype("object")
    expanded = result[src_col].astype(str).str.split(delim, expand=True, n=n_parts - 1)

    for i in range(n_parts):
        col_name = f"{prefix}{i + 1}"
        if i < expanded.shape[1]:
            result[col_name] = expanded[i]
        else:
            result[col_name] = ""

    st.success(f"✅ Đã tách `{src_col}` thành {n_parts} cột.")
    st.dataframe(result, use_container_width=True, height=380)
    _export(result)


def _run_extract(df: pd.DataFrame, src_col: str, pattern: str, new_name: str) -> None:
    result = df.copy().astype("object")
    extracted = result[src_col].astype(str).str.extract(pattern, expand=False)
    result[new_name] = extracted
    matched = int(extracted.notna().sum())

    st.success(f"✅ Trích được {matched}/{len(df)} dòng từ `{src_col}`.")
    st.dataframe(result, use_container_width=True, height=380)
    _export(result)


def _export(df: pd.DataFrame) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tách Trích Cột"
    ws.append(list(df.columns))
    style_header(ws)
    for _, row in df.iterrows():
        ws.append(list(row))
    auto_col_width(ws)
    freeze_header(ws)
    excel_download_button(save_workbook_to_bytes(wb), filename="split_extract.xlsx")
