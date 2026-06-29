"""
tools/validate_rules.py
───────────────────────
Tool: ✅ Validate Rules
Kiểm tra rule nghiệp vụ trên dữ liệu, xuất dòng lỗi.
"""

import re

import streamlit as st
import pandas as pd
import openpyxl

from modules.excel_utils import (
    RED_FILL, style_header, auto_col_width, freeze_header, save_workbook_to_bytes,
)
from modules.ui_components import (
    show_single_preview, show_stats, excel_download_button, no_file_warning,
)

RULE_TYPES = {
    "Không được trống": "not_null",
    "Lớn hơn": "gt",
    "Nhỏ hơn": "lt",
    "Bằng": "eq",
    "Khác": "neq",
    "Chứa text": "contains",
    "Regex khớp": "regex",
    "Email hợp lệ": "email",
}


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    options = [f"Sheet 1 ({name1})"]
    if df2 is not None:
        options.append(f"Sheet 2 ({name2})")
    target = st.radio("Validate sheet nào?", options, horizontal=True, key="val_target")

    df = df1 if "Sheet 1" in target else df2
    show_single_preview(df, name1 if "Sheet 1" in target else name2)
    st.markdown("### ✅ Validate Rules")

    n_rules = st.number_input("Số rule", 1, 10, 1, key="val_n_rules")
    rules = []
    cols = list(df.columns)

    for i in range(int(n_rules)):
        st.markdown(f"**Rule #{i + 1}**")
        c1, c2, c3 = st.columns([2, 2, 2])
        col = c1.selectbox("Cột", cols, key=f"val_col_{i}")
        rule_label = c2.selectbox("Rule", list(RULE_TYPES.keys()), key=f"val_rule_{i}")
        value = c3.text_input("Giá trị / pattern", "", key=f"val_val_{i}")
        rules.append({"col": col, "type": RULE_TYPES[rule_label], "value": value})

    if st.button("🚀 Chạy Validate", use_container_width=True):
        _run(df, rules)


def _check_rule(series: pd.Series, rule: dict) -> pd.Series:
    """Trả về True = PASS, False = FAIL."""
    s = series
    rtype = rule["type"]
    val = rule.get("value", "")

    if rtype == "not_null":
        return s.notna() & (s.astype(str).str.strip() != "")

    if rtype == "email":
        pat = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
        return s.astype(str).str.match(pat, na=False)

    if rtype == "contains":
        return s.astype(str).str.contains(re.escape(val), na=False, regex=True)

    if rtype == "regex":
        try:
            return s.astype(str).str.match(val, na=False)
        except re.error:
            return pd.Series(False, index=s.index)

    # numeric comparisons
    num = pd.to_numeric(s, errors="coerce")
    if rtype == "gt":
        return num > float(val)
    if rtype == "lt":
        return num < float(val)
    if rtype == "eq":
        return s.astype(str) == str(val)
    if rtype == "neq":
        return s.astype(str) != str(val)

    return pd.Series(True, index=s.index)


def _run(df: pd.DataFrame, rules: list) -> None:
    result = df.copy().astype("object")
    fail_masks = []

    for i, rule in enumerate(rules):
        col = rule["col"]
        if col not in result.columns:
            continue
        ok = _check_rule(result[col], rule)
        fail_col = f"_FAIL_{i+1}_{col}"
        result[fail_col] = (~ok).map({True: "❌", False: ""})
        fail_masks.append(~ok)

    if not fail_masks:
        st.warning("Không có rule hợp lệ.")
        return

    any_fail = fail_masks[0]
    for m in fail_masks[1:]:
        any_fail = any_fail | m

    n_fail = int(any_fail.sum())
    n_pass = len(df) - n_fail

    show_stats([
        {"num": str(len(df)), "lbl": "Tổng dòng", "color": None},
        {"num": str(n_pass), "lbl": "Dòng pass", "color": "#4dbb77"},
        {"num": str(n_fail), "lbl": "Dòng fail", "color": "#ff6b6b"},
    ])

    tab1, tab2 = st.tabs(["❌ Dòng lỗi", "📋 Toàn bộ + cột FAIL"])
    with tab1:
        if n_fail:
            st.dataframe(result[any_fail], use_container_width=True, height=380)
        else:
            st.success("🎉 Tất cả dòng đều pass!")
    with tab2:
        st.dataframe(result, use_container_width=True, height=380)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validate"
    ws.append(list(result.columns))
    style_header(ws)
    for idx, row in result.iterrows():
        ws.append(list(row))
        if any_fail.loc[idx]:
            for cell in ws[ws.max_row]:
                cell.fill = RED_FILL
    auto_col_width(ws)
    freeze_header(ws)
    excel_download_button(save_workbook_to_bytes(wb), filename="validate_result.xlsx")
