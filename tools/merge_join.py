"""
tools/merge_join.py
────────────────────
Tool: 🔀 Merge / JOIN
Ghép 2 sheet theo kiểu LEFT / INNER / RIGHT / OUTER JOIN.
"""

import streamlit as st
import pandas as pd

from modules.excel_utils import (
    normalize_key, digit_key_width,
    style_header, auto_col_width, freeze_header,
    save_workbook_to_bytes,
)
from modules.ui_components import (
    show_data_preview, show_stats,
    excel_download_button, no_file_warning,
)
import openpyxl


JOIN_MAP = {
    "LEFT JOIN":  "left",
    "INNER JOIN": "inner",
    "RIGHT JOIN": "right",
    "OUTER JOIN": "outer",
}


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    show_data_preview(df1, df2, name1, name2)
    st.markdown("### ⚙️ Cấu hình JOIN")

    cols1, cols2 = list(df1.columns), list(df2.columns)

    c1, c2 = st.columns(2)
    key1 = c1.multiselect(f"Cột JOIN của {name1} (key)", cols1, key="jk1")
    key2 = c2.multiselect(f"Cột JOIN của {name2} (key)", cols2, key="jk2")

    join_type = st.select_slider(
        "Kiểu JOIN",
        options=list(JOIN_MAP.keys()),
        value="LEFT JOIN",
        key="join_type",
    )

    _show_join_diagram(join_type)

    c1, c2 = st.columns(2)
    suffix1 = c1.text_input("Suffix cột trùng tên (Sheet1)", "_s1", key="sfx1")
    suffix2 = c2.text_input("Suffix cột trùng tên (Sheet2)", "_s2", key="sfx2")

    # Validation
    if key1 and key2 and len(key1) != len(key2):
        st.warning("⚠️ Số cột key của 2 sheet phải bằng nhau.")
        return

    if key1 and key2 and st.button("🚀 Chạy JOIN", use_container_width=True):
        _run(df1, df2, key1, key2, join_type, suffix1, suffix2, name1, name2)
    elif not (key1 and key2):
        st.info("ℹ️ Hãy chọn cột key cho cả 2 sheet.")


# ── Core logic ────────────────────────────────────────────────────────────────

def _run(df1, df2, key1, key2, join_type, suffix1, suffix2, name1, name2):
    left, right = df1.copy(), df2.copy()
    left_on, right_on = [], []
    drop_cols = []
    for i, (c1, c2) in enumerate(zip(key1, key2)):
        width = digit_key_width(df1[c1], df2[c2])
        lc, rc = f"__jk_l{i}", f"__jk_r{i}"
        left[lc] = df1[c1].map(lambda v, w=width: normalize_key(v, True, True, w))
        right[rc] = df2[c2].map(lambda v, w=width: normalize_key(v, True, True, w))
        left_on.append(lc)
        right_on.append(rc)
        drop_cols.extend([lc, rc])

    try:
        merged = pd.merge(
            left, right,
            left_on=left_on, right_on=right_on,
            how=JOIN_MAP[join_type],
            suffixes=(suffix1, suffix2),
        )
        merged = merged.drop(columns=[c for c in drop_cols if c in merged.columns])
    except Exception as e:
        st.error(f"❌ Lỗi khi JOIN: {e}")
        return

    show_stats([
        {"num": str(len(df1)),      "lbl": f"{name1} dòng",  "color": None},
        {"num": str(len(df2)),      "lbl": f"{name2} dòng",  "color": None},
        {"num": str(len(merged)),   "lbl": "Kết quả dòng",   "color": "#4dbb77"},
        {"num": str(len(merged.columns)), "lbl": "Số cột",   "color": "#7eb8f7"},
    ])

    st.dataframe(merged, use_container_width=True, height=380)

    # ── Excel export
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "JOIN Result"
    ws.append(list(merged.columns))
    style_header(ws)
    for _, row in merged.iterrows():
        ws.append(list(row))
    auto_col_width(ws)
    freeze_header(ws)

    excel_download_button(save_workbook_to_bytes(wb), filename="join_result.xlsx")


def _show_join_diagram(join_type: str) -> None:
    """Hiển thị mô tả ngắn về kiểu JOIN đang chọn."""
    descriptions = {
        "LEFT JOIN":  "📌 **LEFT JOIN** — Giữ tất cả dòng Sheet1, Sheet2 không khớp → NULL",
        "INNER JOIN": "📌 **INNER JOIN** — Chỉ giữ dòng khớp ở cả 2 sheet",
        "RIGHT JOIN": "📌 **RIGHT JOIN** — Giữ tất cả dòng Sheet2, Sheet1 không khớp → NULL",
        "OUTER JOIN": "📌 **OUTER JOIN** — Giữ tất cả dòng cả 2 sheet, không khớp → NULL",
    }
    st.markdown(descriptions.get(join_type, ""))
