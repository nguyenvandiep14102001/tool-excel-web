"""
tools/normalize_data.py
────────────────────────
Tool: 🔄 Chuẩn Hoá Dữ Liệu
Trim, upper/lower/title, xoá ký tự đặc biệt, chuẩn hoá số/ngày, fill NA.
"""

import streamlit as st
import pandas as pd
import re

from modules.excel_utils import (
    style_header, auto_col_width, freeze_header,
    save_workbook_to_bytes,
)
from modules.ui_components import (
    show_single_preview, show_stats,
    excel_download_button, no_file_warning,
)
import openpyxl


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    # Cho chọn sheet muốn chuẩn hoá
    target_label = st.radio(
        "Chuẩn hoá sheet nào?",
        [f"Sheet 1 ({name1})", f"Sheet 2 ({name2})"],
        horizontal=True,
        key="norm_target",
    )
    df = df1 if "Sheet 1" in target_label else df2
    show_single_preview(df, name1 if "Sheet 1" in target_label else name2)

    st.markdown("### ⚙️ Chọn thao tác chuẩn hoá")

    cols = list(df.columns)
    target_cols = st.multiselect(
        "Áp dụng cho cột (để trống = tất cả cột)",
        cols, key="norm_cols",
    )

    # ── Options grid
    st.markdown("#### Tuỳ chọn xử lý")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**📝 Text**")
        do_trim    = st.checkbox("✂️ Trim khoảng trắng",     True,  key="do_trim")
        do_upper   = st.checkbox("🔠 UPPERCASE",                     key="do_upper")
        do_lower   = st.checkbox("🔡 lowercase",                     key="do_lower")
        do_title   = st.checkbox("📝 Title Case",                    key="do_title")
        do_remove  = st.checkbox("🚫 Xoá ký tự đặc biệt",          key="do_remove")
        keep_chars = st.text_input(
            "Ký tự được giữ lại (ngoài a-z 0-9)",
            value=" .,@-_", key="keep_chars",
        ) if do_remove else ""

    with col2:
        st.markdown("**🔢 Số & Ngày**")
        do_number  = st.checkbox("🔢 Chuẩn hoá số (xoá dấu phẩy/khoảng trắng)", key="do_num")
        do_date    = st.checkbox("📅 Chuẩn hoá ngày",                            key="do_date")
        date_fmt_in  = st.text_input("Format ngày gốc",   "%d/%m/%Y", key="dfmt_in")  if do_date else ""
        date_fmt_out = st.text_input("Format ngày đầu ra", "%Y-%m-%d", key="dfmt_out") if do_date else ""

    with col3:
        st.markdown("**🩹 Giá trị trống**")
        do_fill_na = st.checkbox("🩹 Điền giá trị trống",            key="do_fill")
        fill_val   = st.text_input("Giá trị điền vào ô trống", "N/A", key="fill_val") if do_fill_na else None
        do_drop_na = st.checkbox("🗑️ Xoá dòng có ô trống",           key="do_dropna")
        na_cols    = st.multiselect("Chỉ kiểm tra các cột:", cols,   key="na_cols")   if do_drop_na else []
        st.markdown("**🧹 Khác**")
        do_dedup   = st.checkbox("🧹 Dedup nhanh",                    key="do_dedup")
        do_strip_nl = st.checkbox("↩️ Xoá xuống dòng trong ô",       key="do_nl")

    if st.button("🚀 Chuẩn Hoá Ngay", use_container_width=True):
        _run(df, target_cols, {
            "trim": do_trim, "upper": do_upper, "lower": do_lower,
            "title": do_title, "remove": do_remove, "keep_chars": keep_chars,
            "number": do_number,
            "date": do_date, "date_fmt_in": date_fmt_in, "date_fmt_out": date_fmt_out,
            "fill_na": do_fill_na, "fill_val": fill_val,
            "drop_na": do_drop_na, "na_cols": na_cols,
            "dedup": do_dedup, "strip_nl": do_strip_nl,
        })


# ── Core logic ────────────────────────────────────────────────────────────────

def _run(df: pd.DataFrame, target_cols: list, opts: dict):
    result = df.copy()
    apply_to = target_cols if target_cols else list(df.columns)

    rows_before = len(result)

    for col in apply_to:
        s = result[col].copy()

        if opts["strip_nl"]:
            s = s.astype(str).str.replace(r"[\r\n]+", " ", regex=True)

        if opts["trim"]:
            s = s.astype(str).str.strip()

        if opts["upper"]:
            s = s.astype(str).str.upper()
        elif opts["lower"]:
            s = s.astype(str).str.lower()
        elif opts["title"]:
            s = s.astype(str).str.title()

        if opts["remove"]:
            safe = re.escape(opts.get("keep_chars", " .,@-_"))
            pattern = rf"[^\w{safe}]"
            s = s.astype(str).str.replace(pattern, "", regex=True)

        if opts["number"]:
            cleaned = s.astype(str).str.replace(",", "").str.replace(" ", "")
            numeric  = pd.to_numeric(cleaned, errors="coerce")
            # Chỉ replace những ô thực sự là số
            s = numeric.where(numeric.notna(), s)

        if opts["date"] and opts["date_fmt_in"]:
            try:
                parsed = pd.to_datetime(s.astype(str), format=opts["date_fmt_in"], errors="coerce")
                fmt_out = opts["date_fmt_out"] or "%Y-%m-%d"
                s = parsed.dt.strftime(fmt_out).where(parsed.notna(), s)
            except Exception:
                pass

        if opts["fill_na"] and opts["fill_val"] is not None:
            s = s.replace({"nan": opts["fill_val"], "None": opts["fill_val"],
                            "NaT": opts["fill_val"], "": opts["fill_val"]})
            s = s.fillna(opts["fill_val"])

        result[col] = s

    # Toàn bảng
    if opts["drop_na"]:
        check = opts["na_cols"] if opts["na_cols"] else None
        result = result.dropna(subset=check).reset_index(drop=True)

    if opts["dedup"]:
        result = result.drop_duplicates().reset_index(drop=True)

    rows_after  = len(result)
    rows_removed = rows_before - rows_after

    show_stats([
        {"num": str(rows_before),  "lbl": "Dòng ban đầu",   "color": None},
        {"num": str(rows_after),   "lbl": "Dòng sau xử lý", "color": "#4dbb77"},
        {"num": str(rows_removed), "lbl": "Dòng bị xoá",    "color": "#ff6b6b" if rows_removed else None},
    ])

    st.success(f"✅ Chuẩn hoá xong! {rows_after} dòng · {len(result.columns)} cột")
    st.dataframe(result, use_container_width=True, height=380)

    # ── Excel export
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "Đã Chuẩn Hoá"
    ws.append(list(result.columns))
    style_header(ws)
    for _, row in result.iterrows():
        ws.append(list(row))
    auto_col_width(ws)
    freeze_header(ws)

    excel_download_button(save_workbook_to_bytes(wb), filename="chuan_hoa.xlsx")
