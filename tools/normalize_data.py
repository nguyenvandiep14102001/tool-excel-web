"""
tools/normalize_data.py
────────────────────────
Tool: 🔄 Chuẩn Hoá Dữ Liệu
Trim, chữ hoa/thường, xoá ký tự đặc biệt, chuẩn hoá số/ngày, fill NA, dedup.
"""

import re

import openpyxl
import pandas as pd
import streamlit as st

from modules.excel_utils import (
    auto_col_width,
    freeze_header,
    save_workbook_to_bytes,
    style_header,
)
from modules.ui_components import (
    excel_download_button,
    no_file_warning,
    show_single_preview,
    show_stats,
)

MISSING_TEXT_VALUES = {"", "nan", "none", "nat"}


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None, name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    target_label = st.radio(
        "Chuẩn hoá sheet nào?",
        [f"Sheet 1 ({name1})", f"Sheet 2 ({name2})"],
        horizontal=True,
        key="norm_target",
    )

    df = df1 if "Sheet 1" in target_label else df2
    if df is None:
        st.warning("⚠️ Cần có cả Sheet 1 và Sheet 2 để chọn.")
        return

    show_single_preview(df, name1 if "Sheet 1" in target_label else name2)
    st.markdown("### ⚙️ Chọn thao tác chuẩn hoá")

    cols = list(df.columns)
    target_cols = st.multiselect("Áp dụng cho cột (để trống = tất cả cột)", cols, key="norm_cols")

    st.markdown("#### Tuỳ chọn xử lý")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**📝 Text**")
        do_trim = st.checkbox("✂️ Trim khoảng trắng", True, key="do_trim")
        do_upper = st.checkbox("🔠 UPPERCASE", key="do_upper")
        do_lower = st.checkbox("🔡 lowercase", key="do_lower")
        do_title = st.checkbox("📝 Title Case", key="do_title")
        do_remove = st.checkbox("🚫 Xoá ký tự đặc biệt", key="do_remove")
        keep_chars = (
            st.text_input("Ký tự giữ lại (ngoài chữ/số)", value=" .,@-_/", key="keep_chars") if do_remove else ""
        )
        do_strip_nl = st.checkbox("↩️ Xoá xuống dòng trong ô", key="do_nl")

    with col2:
        st.markdown("**🔢 Số & Ngày**")
        do_number = st.checkbox("🔢 Chuẩn hoá số (xoá dấu phẩy/khoảng trắng)", key="do_num")
        do_date = st.checkbox("📅 Chuẩn hoá ngày", key="do_date")
        date_auto = st.checkbox("Tự nhận diện format ngày", value=True, key="date_auto") if do_date else True
        date_fmt_in = (
            st.text_input("Format ngày gốc (khi tắt tự nhận diện)", "%d/%m/%Y", key="dfmt_in")
            if do_date and not date_auto
            else ""
        )
        date_fmt_out = (
            st.selectbox(
                "Format ngày đầu ra",
                ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"],
                index=0,
                key="dfmt_out",
            )
            if do_date
            else ""
        )

    with col3:
        st.markdown("**🩹 Giá trị trống**")
        do_fill_na = st.checkbox("🩹 Điền giá trị trống", key="do_fill")
        fill_val = st.text_input("Giá trị điền vào ô trống", "N/A", key="fill_val") if do_fill_na else None
        do_drop_na = st.checkbox("🗑️ Xoá dòng có ô trống", key="do_dropna")
        na_cols = st.multiselect("Chỉ kiểm tra các cột:", cols, key="na_cols") if do_drop_na else []

        st.markdown("**🧹 Dedup**")
        do_dedup = st.checkbox("🧹 Dedup nhanh", key="do_dedup")
        dedup_cols = st.multiselect("Dedup theo cột (để trống = toàn dòng)", cols, key="dedup_cols") if do_dedup else []
        dedup_keep = (
            st.selectbox("Giữ dòng", ["first", "last"], index=0, key="dedup_keep") if do_dedup else "first"
        )

    if st.button("🚀 Chuẩn Hoá Ngay", use_container_width=True):
        _run(
            df,
            target_cols,
            {
                "trim": do_trim,
                "upper": do_upper,
                "lower": do_lower,
                "title": do_title,
                "remove": do_remove,
                "keep_chars": keep_chars,
                "number": do_number,
                "date": do_date,
                "date_auto": date_auto,
                "date_fmt_in": date_fmt_in,
                "date_fmt_out": date_fmt_out,
                "fill_na": do_fill_na,
                "fill_val": fill_val,
                "drop_na": do_drop_na,
                "na_cols": na_cols,
                "dedup": do_dedup,
                "dedup_cols": dedup_cols,
                "dedup_keep": dedup_keep,
                "strip_nl": do_strip_nl,
            },
        )


def _strip_newlines(series: pd.Series) -> pd.Series:
    out = series.copy()
    mask = out.notna()
    text = out.loc[mask].astype(str)
    out.loc[mask] = text.str.replace(r"[\r\n]+", " ", regex=True)
    return out


def _trim_values(series: pd.Series) -> pd.Series:
    out = series.copy()
    mask = out.notna()
    text = out.loc[mask].astype(str)
    out.loc[mask] = text.str.strip()
    return out


def _apply_text_case_and_remove(series: pd.Series, opts: dict) -> pd.Series:
    out = series.copy()
    mask = out.notna()
    text = out.loc[mask].astype(str)

    if opts["upper"]:
        text = text.str.upper()
    elif opts["lower"]:
        text = text.str.lower()
    elif opts["title"]:
        text = text.str.title()

    if opts["remove"]:
        safe = re.escape(opts.get("keep_chars", " .,@-_/"))
        pattern = rf"[^\w{safe}]"
        text = text.str.replace(pattern, "", regex=True)

    out.loc[mask] = text
    return out


def _apply_text_ops(series: pd.Series, opts: dict) -> pd.Series:
    out = series.copy()
    if opts["strip_nl"]:
        out = _strip_newlines(out)
    if opts["trim"]:
        out = _trim_values(out)
    return _apply_text_case_and_remove(out, opts)


def _is_excel_serial(val) -> bool:
    if isinstance(val, bool) or pd.isna(val):
        return False
    if isinstance(val, (int, float)):
        return 1 < float(val) < 1_000_000
    return False


def _parse_dates_series(series: pd.Series, opts: dict) -> tuple[pd.Series, int]:
    """
  Parse nhiều kiểu ngày: datetime Excel, chuỗi, serial number.
  Trả về (series đã format, số ô chuẩn hoá thành công).
    """
    fmt_out = (opts.get("date_fmt_out") or "%Y-%m-%d").strip()
    out = series.astype("object").copy()
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    if pd.api.types.is_datetime64_any_dtype(series):
        parsed = pd.to_datetime(series, errors="coerce")
    else:
        values = series.astype("object")

        ts_mask = values.map(lambda x: isinstance(x, pd.Timestamp))
        if ts_mask.any():
            parsed.loc[ts_mask] = pd.to_datetime(values.loc[ts_mask], errors="coerce")

        serial_mask = values.map(_is_excel_serial)
        if serial_mask.any():
            nums = pd.to_numeric(values.loc[serial_mask], errors="coerce")
            parsed.loc[serial_mask] = pd.to_datetime(
                nums, unit="D", origin="1899-12-30", errors="coerce"
            )

        remain = ~(ts_mask | serial_mask) & values.notna()
        if remain.any():
            texts = values.loc[remain].map(
                lambda x: str(x).strip() if not isinstance(x, str) else x.strip()
            )
            texts = texts[~texts.str.lower().isin(MISSING_TEXT_VALUES)]
            if not texts.empty:
                # Auto mode: chỉ parse chuỗi có dấu phân tách ngày để tránh nhầm số thuần (vd "1234")
                # thành năm/thời gian và gây lỗi OutOfBounds.
                if opts.get("date_auto", True):
                    candidate_mask = texts.str.contains(r"[\/\-.]", regex=True)
                    candidates = texts[candidate_mask]
                else:
                    candidates = texts

                if not candidates.empty:
                    fmt_in = (opts.get("date_fmt_in") or "").strip()

                    def _safe_parse_one(val: str):
                        try:
                            if opts.get("date_auto", True) or not fmt_in:
                                return pd.to_datetime(val, errors="raise", dayfirst=True)
                            return pd.to_datetime(val, format=fmt_in, errors="raise")
                        except Exception:
                            return pd.NaT

                    str_parsed = candidates.map(_safe_parse_one)
                    str_parsed = pd.to_datetime(str_parsed, errors="coerce")
                    parsed.loc[str_parsed.index] = str_parsed

    ok = parsed.notna()
    converted = int(ok.sum())
    if converted:
        formatted = parsed.dt.strftime(fmt_out)
        out.loc[ok] = formatted.loc[ok]
    return out, converted


def _normalize_number(series: pd.Series) -> pd.Series:
    # Không dùng gán loc trên string[pyarrow] để tránh TypeError.
    values = series.astype("object")

    def _convert_one(val):
        if pd.isna(val):
            return val
        s = str(val).replace(",", "").replace(" ", "").replace("\u00A0", "")
        num = pd.to_numeric(s, errors="coerce")
        return num if pd.notna(num) else val

    return values.map(_convert_one).astype("object")


def _fill_missing(series: pd.Series, fill_val: str) -> pd.Series:
    out = series.copy()
    mask = out.notna()
    text = out.loc[mask].astype(str).str.strip().str.lower()
    out.loc[mask] = out.loc[mask].where(~text.isin(MISSING_TEXT_VALUES), fill_val)
    out = out.fillna(fill_val)
    return out


def _run(df: pd.DataFrame, target_cols: list, opts: dict) -> None:
    # Ép về object để các thao tác trộn kiểu (str/number/date) không bị kẹt dtype string[pyarrow].
    result = df.copy().astype("object")
    apply_to = target_cols if target_cols else list(df.columns)
    rows_before = len(result)
    date_converted = 0

    for col in apply_to:
        s = result[col].copy()

        # Trim / xuống dòng trước, nhưng chuẩn hoá ngày trước các thao tác phá format ngày
        if opts["strip_nl"]:
            s = _strip_newlines(s)
        if opts["trim"]:
            s = _trim_values(s)
        if opts["date"]:
            s, n = _parse_dates_series(s, opts)
            date_converted += n
        if opts["number"]:
            s = _normalize_number(s)
        if opts["upper"] or opts["lower"] or opts["title"] or opts["remove"]:
            s = _apply_text_case_and_remove(s, opts)
        if opts["fill_na"] and opts["fill_val"] is not None:
            s = _fill_missing(s, opts["fill_val"])
        result[col] = s

    if opts["drop_na"]:
        check_cols = opts["na_cols"] if opts["na_cols"] else list(result.columns)
        cleaned_for_na = result[check_cols].replace(r"^\s*$", pd.NA, regex=True)
        keep_mask = cleaned_for_na.notna().all(axis=1)
        result = result.loc[keep_mask].reset_index(drop=True)

    if opts["dedup"]:
        dedup_subset = opts["dedup_cols"] if opts["dedup_cols"] else None
        result = result.drop_duplicates(subset=dedup_subset, keep=opts.get("dedup_keep", "first")).reset_index(drop=True)

    rows_after = len(result)
    rows_removed = rows_before - rows_after

    show_stats(
        [
            {"num": str(rows_before), "lbl": "Dòng ban đầu", "color": None},
            {"num": str(rows_after), "lbl": "Dòng sau xử lý", "color": "#4dbb77"},
            {"num": str(rows_removed), "lbl": "Dòng bị xoá", "color": "#ff6b6b" if rows_removed else None},
            {"num": str(date_converted), "lbl": "Ô ngày đã chuẩn hoá", "color": "#7eb8f7" if date_converted else None},
        ]
    )

    msg = f"✅ Chuẩn hoá xong! {rows_after} dòng · {len(result.columns)} cột"
    if opts["date"]:
        msg += f" · {date_converted} ô ngày đã đổi format"
    st.success(msg)
    st.dataframe(result, use_container_width=True, height=380)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Đã Chuẩn Hoá"
    ws.append(list(result.columns))
    style_header(ws)
    for _, row in result.iterrows():
        ws.append(list(row))
    auto_col_width(ws)
    freeze_header(ws)
    excel_download_button(save_workbook_to_bytes(wb), filename="chuan_hoa.xlsx")
