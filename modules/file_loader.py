"""
modules/file_loader.py
───────────────────────
Xử lý toàn bộ logic upload file và chọn sheet.
Trả về df1, df2, sheet1_name, sheet2_name cho các tool sử dụng.
"""

import streamlit as st
import pandas as pd
from modules.excel_utils import load_excel_file, read_sheet


def render_file_loader() -> tuple[pd.DataFrame | None, pd.DataFrame | None, str, str]:
    """
    Hiển thị UI upload file và chọn sheet.

    Returns:
        (df1, df2, sheet1_name, sheet2_name)
        df1/df2 = None nếu chưa upload đủ file.
    """
    with st.expander("📁 **Upload File Excel**", expanded=True):
        mode = st.radio(
            "Nguồn dữ liệu",
            ["2 File Excel riêng biệt", "1 File — 2 Sheet"],
            horizontal=True,
            key="upload_mode",
        )

        if mode == "2 File Excel riêng biệt":
            return _load_two_files()
        else:
            return _load_one_file_two_sheets()


# ── Private helpers ───────────────────────────────────────────────────────────

def _load_two_files():
    col1, col2 = st.columns(2)
    with col1:
        f1 = st.file_uploader("📄 File 1 (Sheet 1)", type=["xlsx", "xls"], key="f1")
    with col2:
        f2 = st.file_uploader("📄 File 2 (Sheet 2)", type=["xlsx", "xls"], key="f2")

    if not (f1 and f2):
        return None, None, "", ""

    c1, c2 = st.columns(2)
    with c1:
        h1 = st.number_input("Dòng header File 1 (0 = dòng đầu)", 0, 20, 0, key="h1")
    with c2:
        h2 = st.number_input("Dòng header File 2 (0 = dòng đầu)", 0, 20, 0, key="h2")

    _, sheets1 = load_excel_file(f1)
    _, sheets2 = load_excel_file(f2)

    c1, c2 = st.columns(2)
    with c1:
        s1 = st.selectbox("Sheet trong File 1", sheets1, key="s1")
    with c2:
        s2 = st.selectbox("Sheet trong File 2", sheets2, key="s2")

    df1 = read_sheet(f1, s1, int(h1))
    df2 = read_sheet(f2, s2, int(h2))
    return df1, df2, f"File1::{s1}", f"File2::{s2}"


def _load_one_file_two_sheets():
    col1, _ = st.columns(2)
    with col1:
        f1 = st.file_uploader("📄 File Excel (chứa 2 sheet)", type=["xlsx", "xls"], key="f1b")

    if not f1:
        return None, None, "", ""

    h_row = st.number_input("Dòng header (0 = dòng đầu)", 0, 20, 0, key="hb")
    _, sheets = load_excel_file(f1)

    if len(sheets) < 2:
        st.warning("⚠️ File chỉ có 1 sheet. Cần ít nhất 2 sheet.")
        return None, None, "", ""

    c1, c2 = st.columns(2)
    with c1:
        s1 = st.selectbox("Sheet 1", sheets, index=0, key="s1b")
    with c2:
        s2 = st.selectbox("Sheet 2", sheets, index=1, key="s2b")

    df1 = read_sheet(f1, s1, int(h_row))
    df2 = read_sheet(f1, s2, int(h_row))
    return df1, df2, s1, s2
