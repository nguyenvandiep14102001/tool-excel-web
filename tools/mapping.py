"""
tools/mapping.py
─────────────────
Tool: 🔗 Mapping / VLOOKUP
Tra cứu giá trị từ Sheet2 và gắn vào Sheet1 theo key column.
"""

import streamlit as st
import pandas as pd
import openpyxl

from modules.excel_utils import (
    normalize_key, digit_key_width,
    GREEN_FILL, RED_FILL,
    style_header, auto_col_width, freeze_header,
    save_workbook_to_bytes,
)
from modules.ui_components import (
    show_data_preview, normalize_options,
    show_stats, excel_download_button, no_file_warning,
)


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    show_data_preview(df1, df2, name1, name2)
    st.markdown("### ⚙️ Cấu hình Mapping / VLOOKUP")

    strip, icase = normalize_options("map")

    cols1, cols2 = list(df1.columns), list(df2.columns)

    st.markdown("#### Cấu hình lookup key")
    c1, c2 = st.columns(2)
    lookup_col1 = c1.selectbox(f"Cột KEY trong {name1}", cols1, key="lc1")
    lookup_col2 = c2.selectbox(f"Cột KEY trong {name2}", cols2, key="lc2")

    st.markdown("#### Cột giá trị cần lấy từ Sheet2")
    return_cols = st.multiselect("Chọn cột(s) cần lấy giá trị", cols2, key="rc")

    if not return_cols:
        st.info("ℹ️ Hãy chọn ít nhất 1 cột giá trị cần lấy từ Sheet2.")
        return

    st.markdown("#### Đặt tên cột đích trong Sheet1")
    dest_names: dict[str, str] = {}
    dest_cols = st.columns(min(len(return_cols), 3))
    for i, rc in enumerate(return_cols):
        default = rc if rc not in cols1 else f"{rc}_mapped"
        dest_names[rc] = dest_cols[i % 3].text_input(
            f"Tên đích cho '{rc}'", default, key=f"dn_{rc}"
        )

    handle_dup = st.selectbox(
        "Nếu Sheet2 có nhiều dòng khớp, lấy giá trị nào?",
        ["Đầu tiên", "Cuối cùng", "Tất cả (nối bằng dấu phẩy)"],
        key="dup_handle",
    )

    if st.button("🚀 Chạy Mapping", use_container_width=True):
        _run(df1, df2, lookup_col1, lookup_col2,
             return_cols, dest_names, handle_dup, strip, icase)


# ── Core logic ────────────────────────────────────────────────────────────────

def _run(df1, df2, key1, key2, return_cols, dest_names, handle_dup, strip, icase):
    # Build lookup dict: norm_key → list of {col: val}
    key_width = digit_key_width(df1[key1], df2[key2])
    lookup: dict[str, list[dict]] = {}
    for _, row in df2.iterrows():
        k = normalize_key(row[key2], strip, icase, key_width)
        vals = {rc: row[rc] for rc in return_cols}
        lookup.setdefault(k, []).append(vals)

    result_df = df1.copy()

    for rc in return_cols:
        dest = dest_names[rc]
        resolved = []
        for v in df1[key1]:
            k = normalize_key(v, strip, icase, key_width)
            if k in lookup:
                entries = lookup[k]
                if handle_dup == "Đầu tiên":
                    resolved.append(entries[0][rc])
                elif handle_dup == "Cuối cùng":
                    resolved.append(entries[-1][rc])
                else:
                    resolved.append(", ".join(str(e[rc]) for e in entries))
            else:
                resolved.append(None)
        result_df[dest] = resolved

    # ── Stats
    first_dest = dest_names[return_cols[0]]
    n_match   = int(result_df[first_dest].notna().sum())
    n_nomatch = len(result_df) - n_match
    show_stats([
        {"num": str(len(result_df)), "lbl": "Tổng dòng",  "color": None},
        {"num": str(n_match),        "lbl": "Khớp",        "color": "#4dbb77"},
        {"num": str(n_nomatch),      "lbl": "Không khớp", "color": "#ff6b6b"},
    ])

    # ── Preview
    st.dataframe(result_df, use_container_width=True, height=380)

    # ── Excel export
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "Mapped"
    ws.append(list(result_df.columns))
    style_header(ws)

    new_col_indices = [result_df.columns.get_loc(dest_names[rc]) + 1 for rc in return_cols]

    for _, row in result_df.iterrows():
        ws.append(list(row))
        excel_r = ws.max_row
        for ci in new_col_indices:
            cell = ws.cell(excel_r, ci)
            cell.fill = GREEN_FILL if cell.value is not None else RED_FILL

    auto_col_width(ws)
    freeze_header(ws)

    excel_download_button(save_workbook_to_bytes(wb), filename="mapping_result.xlsx")
