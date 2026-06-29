"""
tools/dedup.py
───────────────
Tool: 🧹 Dedup
Phát hiện và loại bỏ dòng trùng lặp.
"""

import streamlit as st
import pandas as pd
import openpyxl

from modules.excel_utils import (
    normalize,
    RED_FILL,
    style_header, auto_col_width, freeze_header,
    save_workbook_to_bytes,
)
from modules.ui_components import (
    show_single_preview, normalize_options,
    show_stats, excel_download_button, no_file_warning,
)


KEEP_MAP = {
    "Dòng đầu tiên":            "first",
    "Dòng cuối cùng":           "last",
    "Xoá tất cả dòng trùng":    False,
}


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    show_single_preview(df1, name1)
    st.markdown("### ⚙️ Cấu hình Dedup")

    cols1 = list(df1.columns)
    dedup_cols = st.multiselect(
        "Cột xác định trùng lặp (để trống = dùng tất cả cột)",
        cols1, key="dcols",
    )

    keep = st.radio(
        "Giữ lại dòng nào khi trùng?",
        list(KEEP_MAP.keys()),
        horizontal=True,
        key="keep_policy",
    )

    strip, icase = normalize_options("dedup")

    # Chọn sheet để dedup (có thể dedup sheet2)
    target = st.radio(
        "Dedup sheet nào?",
        [f"Sheet 1 ({name1})", f"Sheet 2 ({name2})"],
        horizontal=True,
        key="dedup_target",
    )

    if st.button("🚀 Chạy Dedup", use_container_width=True):
        df_target = df1 if "Sheet 1" in target else df2
        _run(df_target, dedup_cols, keep, strip, icase)


# ── Core logic ────────────────────────────────────────────────────────────────

def _run(df: pd.DataFrame, dedup_cols: list, keep: str, strip: bool, icase: bool):
    all_cols   = list(df.columns)
    check_cols = dedup_cols if dedup_cols else all_cols

    # Tạo bản sao normalize để check trùng
    temp = df.copy()
    for c in check_cols:
        temp[f"__norm_{c}"] = temp[c].apply(lambda x: normalize(x, strip, icase))

    norm_cols = [f"__norm_{c}" for c in check_cols]
    dup_mask  = temp.duplicated(subset=norm_cols, keep=KEEP_MAP[keep])

    df_clean = df[~dup_mask].reset_index(drop=True)
    df_dups  = df[dup_mask].reset_index(drop=True)

    n_dup   = int(dup_mask.sum())
    n_clean = len(df_clean)

    show_stats([
        {"num": str(len(df)),   "lbl": "Trước Dedup",   "color": None},
        {"num": str(n_dup),     "lbl": "Dòng trùng",    "color": "#ff6b6b"},
        {"num": str(n_clean),   "lbl": "Sau Dedup",     "color": "#4dbb77"},
        {"num": f"{round(100*n_dup/len(df))}%" if len(df) else "0%",
                                "lbl": "Tỉ lệ trùng",   "color": "#ffa94d"},
    ])

    tab1, tab2 = st.tabs(["✅ Dữ liệu sau Dedup", "🗑️ Dòng đã xoá / trùng"])
    with tab1:
        st.dataframe(df_clean, use_container_width=True, height=350)
    with tab2:
        if len(df_dups):
            st.dataframe(df_dups, use_container_width=True, height=350)
        else:
            st.success("🎉 Không có dòng trùng nào!")

    if not n_dup:
        return

    # ── Excel export (2 sheets)
    wb = openpyxl.Workbook()

    # Sheet đã dedup
    ws1 = wb.active
    ws1.title = "Đã Dedup"
    ws1.append(list(df_clean.columns))
    style_header(ws1)
    for _, row in df_clean.iterrows():
        ws1.append(list(row))
    auto_col_width(ws1)
    freeze_header(ws1)

    # Sheet dòng trùng (tô đỏ)
    ws2 = wb.create_sheet("Dòng Trùng")
    ws2.append(list(df_dups.columns))
    style_header(ws2)
    for _, row in df_dups.iterrows():
        ws2.append(list(row))
        for cell in ws2[ws2.max_row]:
            cell.fill = RED_FILL
    auto_col_width(ws2)
    freeze_header(ws2)

    excel_download_button(save_workbook_to_bytes(wb), filename="dedup_result.xlsx")
