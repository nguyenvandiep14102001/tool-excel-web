"""
tools/compare_columns.py
─────────────────────────
Tool: 🔍 So Sánh Cột
So sánh từng cặp cột giữa 2 sheet (ví dụ: A↔C, B↔D).
Cặp #1 tìm dòng tương ứng trong Sheet2; tất cả cặp phải khớp → xanh.
"""

import streamlit as st
import pandas as pd

from modules.excel_utils import (
    normalize_key, digit_key_width,
    GREEN_FILL, RED_FILL, BLUE_FILL,
    STYLE_MATCH, STYLE_DIFF, STYLE_SHEET2,
    build_key_index,
    save_workbook_to_bytes, create_workbook_with_df,
)
from modules.ui_components import (
    show_data_preview, normalize_options,
    key_pair_selector, column_pair_selector, show_stats,
    excel_download_button, no_file_warning,
)

TOOL_VERSION = "2.4"
SESSION_KEY = "col_cmp_v2_result"


def _short(val, max_len: int = 80) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    text = str(val).replace("\n", " ").strip()
    return text[:max_len] + "…" if len(text) > max_len else text


def _mismatch_note(df1, df2, row_i, s2_row, pairs, col_matches) -> str:
    parts = []
    for ca, cb in pairs:
        if col_matches[ca][row_i]:
            continue
        v1 = _short(df1.iloc[row_i][ca])
        v2 = _short(df2.iloc[s2_row][cb]) if s2_row is not None else ""
        parts.append(f"{ca} ↔ {cb}: '{v1}' ≠ '{v2}'")
    return " | ".join(parts)


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    show_data_preview(df1, df2, name1, name2)
    st.markdown("### ⚙️ Cấu hình so sánh cột")
    st.caption(f"Logic phiên bản **{TOOL_VERSION}** — cặp #1 tra cứu dòng Sheet2, tất cả cặp phải khớp.")

    strip, icase = normalize_options("col_cmp")

    if df2 is None:
        st.warning("⚠️ Cần có Sheet 2 để so sánh cột.")
        return

    st.markdown("#### 🔑 Cột khóa (KEY)")
    st.caption("Dùng để tìm dòng tương ứng giữa Sheet1 và Sheet2 trước khi so sánh.")
    key_col_a, key_col_b = key_pair_selector(
        cols1=list(df1.columns),
        cols2=list(df2.columns),
        label1="KEY Sheet1",
        label2="KEY Sheet2",
        key_prefix="col_cmp_v2_key",
    )

    st.markdown("#### Các cặp cột cần so sánh")
    default_n = max(1, min(2, len(df1.columns), len(df2.columns)))
    pairs = column_pair_selector(
        cols1=list(df1.columns),
        cols2=list(df2.columns),
        label1="Cột Sheet1",
        label2="Cột Sheet2",
        key_prefix="col_cmp_v2",
        default_n=default_n,
    )

    pair_desc = " | ".join(
        f"**{ca}** ↔ **{cb}**" for ca, cb in pairs
    )
    st.info(
        f"📌 KEY: **{key_col_a}** ↔ **{key_col_b}** | "
        f"So sánh: {pair_desc} — dòng xanh khi **tất cả cặp** khớp."
    )

    if st.button("🚀 Chạy So Sánh Cột", use_container_width=True):
        try:
            with st.spinner("Đang so sánh — file lớn có thể mất vài chục giây..."):
                st.session_state[SESSION_KEY] = _compute(
                    df1, df2, key_col_a, key_col_b, pairs, strip, icase,
                )
        except Exception as exc:
            st.session_state.pop(SESSION_KEY, None)
            st.error(f"Không chạy được so sánh: {exc}")
            st.exception(exc)
            return

    cached = st.session_state.get(SESSION_KEY)
    if cached and cached.get("version") == TOOL_VERSION:
        _display(cached)
    elif cached:
        st.warning("Kết quả cũ đã hết hạn. Bấm **Chạy So Sánh Cột** lại.")


# ── Core logic ────────────────────────────────────────────────────────────────

def _compute(df1, df2, key_col_a, key_col_b, pairs, strip, icase) -> dict:
    key_width = digit_key_width(df1[key_col_a], df2[key_col_b])
    key_index = build_key_index(df2, key_col_b, strip, icase, width=key_width)
    pair_w = {(ca, cb): digit_key_width(df1[ca], df2[cb]) for ca, cb in pairs}

    n_rows = len(df1)
    compare_cols = [ca for ca, _ in pairs]
    col_matches = {col_a: [False] * n_rows for col_a, _ in pairs}
    s2_rows: list[int | None] = []
    match_flags: list[bool] = []
    notes: list[str] = []

    def _eq(row_i, s2_row, ca, cb) -> bool:
        w = pair_w[(ca, cb)]
        return normalize_key(df1.iloc[row_i][ca], strip, icase, w) == normalize_key(
            df2.iloc[s2_row][cb], strip, icase, w
        )

    for row_i in range(n_rows):
        key_val = normalize_key(df1.iloc[row_i][key_col_a], strip, icase, key_width)
        candidates = key_index.get(key_val, [])

        if not candidates:
            s2_rows.append(None)
            match_flags.append(False)
            key_show = _short(df1.iloc[row_i][key_col_a])
            notes.append(
                f"Không tìm thấy KEY [{key_col_a}={key_show}] trong Sheet2 [{key_col_b}]"
            )
            continue

        matched_s2 = None
        row_ok = False
        for s2_row in candidates:
            if all(_eq(row_i, s2_row, ca, cb) for ca, cb in pairs):
                row_ok = True
                matched_s2 = s2_row
                for ca, _ in pairs:
                    col_matches[ca][row_i] = True
                break

        if not row_ok:
            matched_s2 = candidates[0]
            for ca, cb in pairs:
                col_matches[ca][row_i] = _eq(row_i, matched_s2, ca, cb)

        s2_rows.append(matched_s2)
        match_flags.append(row_ok)
        notes.append(
            "" if row_ok else _mismatch_note(
                df1, df2, row_i, matched_s2, pairs, col_matches
            )
        )

    result_df = df1.copy()
    for _col_a, col_b in pairs:
        result_df[f"{col_b} (Sheet2)"] = [
            df2.iloc[s2][col_b] if s2 is not None else ""
            for s2 in s2_rows
        ]

    result_df = result_df.reset_index(drop=True)
    match_flags_s = pd.Series(match_flags)
    result_df["_RESULT"] = match_flags_s.map(
        {True: "✅ Khớp", False: "❌ Không khớp"}
    )
    result_df["_NOTE"] = notes

    return {
        "version": TOOL_VERSION,
        "key": (key_col_a, key_col_b),
        "pairs": pairs,
        "result_df": result_df,
        "col_matches": col_matches,
        "compare_cols": compare_cols,
        "n_rows": n_rows,
        "n_match": int(match_flags_s.sum()),
        "n_nomatch": len(match_flags_s) - int(match_flags_s.sum()),
    }


def _display(data: dict) -> None:
    result_df = data["result_df"]
    col_matches = data["col_matches"]
    compare_cols = data["compare_cols"]
    n_rows = data["n_rows"]

    show_stats([
        {"num": str(n_rows),              "lbl": "Tổng dòng",       "color": None},
        {"num": str(data["n_match"]),     "lbl": "Dòng khớp",       "color": "#4dbb77"},
        {"num": str(data["n_nomatch"]),   "lbl": "Dòng không khớp", "color": "#ff6b6b"},
    ])

    display_cols = list(result_df.columns)
    export_df = result_df[display_cols].reset_index(drop=True)
    preview_n = min(200, n_rows)
    preview = export_df.head(preview_n)

    def _highlight(row):
        idx = int(row.name)
        styles = []
        for col in row.index:
            if col == "_RESULT":
                is_ok = str(row[col]).startswith("✅")
                styles.append(STYLE_MATCH if is_ok else STYLE_DIFF)
            elif col in compare_cols:
                ok = col_matches[col][idx] if idx < n_rows else False
                styles.append(STYLE_MATCH if ok else STYLE_DIFF)
            elif col == "_NOTE":
                styles.append(STYLE_DIFF if str(row[col]).strip() else "")
            elif col.endswith(" (Sheet2)"):
                styles.append(STYLE_SHEET2)
            else:
                styles.append("")
        return styles

    try:
        st.dataframe(
            preview.style.apply(_highlight, axis=1),
            use_container_width=True,
            height=380,
        )
    except Exception:
        st.dataframe(preview, use_container_width=True, height=380)

    if n_rows > preview_n:
        st.caption(f"Đang xem {preview_n}/{n_rows} dòng đầu. Tải Excel để xem toàn bộ.")

    with st.spinner("Đang tạo file Excel..."):
        wb, ws = create_workbook_with_df(export_df, "Kết Quả So Sánh Cột")
        result_col = list(export_df.columns).index("_RESULT") if "_RESULT" in export_df.columns else None
        for row_i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=0):
            val = export_df.iat[row_i, result_col] if result_col is not None else ""
            row_fill = GREEN_FILL if str(val).startswith("✅") else RED_FILL
            for ci, cell in enumerate(row, start=1):
                col_name = export_df.columns[ci - 1]
                if col_name in compare_cols:
                    cell.fill = GREEN_FILL if col_matches[col_name][row_i] else RED_FILL
                elif col_name == "_NOTE":
                    if str(cell.value or "").strip():
                        cell.fill = RED_FILL
                elif col_name.endswith(" (Sheet2)"):
                    cell.fill = BLUE_FILL
                else:
                    cell.fill = row_fill

    excel_download_button(
        save_workbook_to_bytes(wb),
        filename="so_sanh_cot.xlsx",
    )
