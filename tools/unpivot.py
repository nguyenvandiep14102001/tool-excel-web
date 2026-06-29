"""
tools/unpivot.py
────────────────
Tool: 📐 Unpivot (Melt)
Chuyển bảng wide → long cho BI / phân tích.
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


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    options = [f"Sheet 1 ({name1})"]
    if df2 is not None:
        options.append(f"Sheet 2 ({name2})")
    target = st.radio("Unpivot sheet nào?", options, horizontal=True, key="unpivot_target")

    df = df1 if "Sheet 1" in target else df2
    show_single_preview(df, name1 if "Sheet 1" in target else name2)
    st.markdown("### 📐 Unpivot (Melt)")

    cols = list(df.columns)
    id_cols = st.multiselect("Cột giữ nguyên (id_vars)", cols, key="unpivot_id")
    remaining = [c for c in cols if c not in id_cols]
    value_cols = st.multiselect(
        "Cột melt (value_vars) — để trống = tất cả cột còn lại",
        remaining,
        key="unpivot_val",
    )
    var_name = st.text_input("Tên cột attribute", "attribute", key="unpivot_var")
    value_name = st.text_input("Tên cột value", "value", key="unpivot_value")

    if st.button("🚀 Unpivot", use_container_width=True):
        _run(df, id_cols, value_cols or remaining, var_name, value_name)


def _run(df, id_cols, value_cols, var_name, value_name):
    if not value_cols:
        st.warning("Chọn ít nhất 1 cột để melt.")
        return

    melted = pd.melt(
        df,
        id_vars=id_cols if id_cols else None,
        value_vars=value_cols,
        var_name=var_name,
        value_name=value_name,
    )

    show_stats([
        {"num": f"{df.shape[0]}×{df.shape[1]}", "lbl": "Trước", "color": None},
        {"num": f"{melted.shape[0]}×{melted.shape[1]}", "lbl": "Sau", "color": "#4dbb77"},
        {"num": str(len(value_cols)), "lbl": "Cột melt", "color": "#7eb8f7"},
    ])

    st.dataframe(melted, use_container_width=True, height=380)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Unpivot"
    ws.append(list(melted.columns))
    style_header(ws)
    for _, row in melted.iterrows():
        ws.append(list(row))
    auto_col_width(ws)
    freeze_header(ws)
    excel_download_button(save_workbook_to_bytes(wb), filename="unpivot_result.xlsx")
