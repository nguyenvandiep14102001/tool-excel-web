"""
tools/compare_rows.py
──────────────────────
Tool: 📏 So Sánh Dòng
Tra cứu dòng Sheet2 theo cột KEY, rồi so sánh từng ô.
Ô giống → xanh | Ô khác → đỏ | Không tìm thấy KEY → đỏ toàn dòng.
"""

import streamlit as st
import pandas as pd
import openpyxl

from modules.excel_utils import (
    normalize_key, digit_key_width,
    GREEN_FILL, RED_FILL, BLUE_FILL,
    STYLE_MATCH, STYLE_DIFF, STYLE_WARN, STYLE_SHEET2,
    build_key_index,
    style_header, auto_col_width, freeze_header,
    save_workbook_to_bytes,
)
from modules.ui_components import (
    show_data_preview, normalize_options,
    key_pair_selector, column_pair_selector, show_stats,
    excel_download_button, no_file_warning,
)

TOOL_VERSION = "2.3"
SESSION_KEY = "row_cmp_v2_result"


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    show_data_preview(df1, df2, name1, name2)
    st.markdown("### ⚙️ Cấu hình so sánh dòng")
    st.caption(
        f"Logic phiên bản **{TOOL_VERSION}** — chọn KEY trước, "
        "tìm dòng Sheet2 tương ứng, so sánh từng ô."
    )

    if df2 is None:
        st.warning("⚠️ Cần có Sheet 2 để so sánh dòng.")
        return

    st.markdown("#### 🔑 Cột khóa (KEY)")
    st.caption("Chọn cột dùng để ghép dòng giữa 2 sheet (ví dụ: Cột 1 ↔ Cột 3).")
    key_col_a, key_col_b = key_pair_selector(
        cols1=list(df1.columns),
        cols2=list(df2.columns),
        label1="KEY Sheet1",
        label2="KEY Sheet2",
        key_prefix="row_cmp_v2_key",
    )

    strip, icase = normalize_options("row_cmp")

    st.markdown("#### Các cột cần so sánh")
    default_n = max(1, min(2, len(df1.columns), len(df2.columns)))
    pairs = column_pair_selector(
        cols1=list(df1.columns),
        cols2=list(df2.columns),
        label1="Cột Sheet1",
        label2="Cột Sheet2",
        key_prefix="row_cmp_v2",
        default_n=default_n,
    )

    pair_desc = " | ".join(f"**{ca}** ↔ **{cb}**" for ca, cb in pairs)
    st.info(
        f"📌 KEY: **{key_col_a}** ↔ **{key_col_b}** | So sánh: {pair_desc}"
    )

    if st.button("🚀 Chạy So Sánh Dòng", use_container_width=True):
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
        st.warning("Kết quả cũ đã hết hạn. Bấm **Chạy So Sánh Dòng** lại.")


def _short(val, max_len: int = 80) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    text = str(val).replace("\n", " ").strip()
    return text[:max_len] + "…" if len(text) > max_len else text


# ── Core logic ────────────────────────────────────────────────────────────────

def _pick_best_row(df1, df2, row_i, candidates, pairs, strip, icase, pair_w) -> int:
    """Chọn dòng Sheet2 khớp nhiều cột nhất (khi KEY trùng nhiều dòng)."""
    best = candidates[0]
    best_score = -1
    for s2_row in candidates:
        score = sum(
            normalize_key(df1.iloc[row_i][ca], strip, icase, pair_w[(ca, cb)])
            == normalize_key(df2.iloc[s2_row][cb], strip, icase, pair_w[(ca, cb)])
            for ca, cb in pairs
        )
        if score > best_score:
            best_score = score
            best = s2_row
    return best


def _compute(df1, df2, key_col_a, key_col_b, pairs, strip, icase) -> dict:
    key_width = digit_key_width(df1[key_col_a], df2[key_col_b])
    key_index = build_key_index(df2, key_col_b, strip, icase, width=key_width)
    pair_w = {(ca, cb): digit_key_width(df1[ca], df2[cb]) for ca, cb in pairs}
    compare_cols = [ca for ca, _ in pairs]
    n_rows = len(df1)

    cell_match: dict[str, list[bool]] = {ca: [False] * n_rows for ca, _ in pairs}
    diff_per_row: list[list[str]] = []
    s2_rows: list[int | None] = []
    notes: list[str] = []
    found_count = not_found_count = same_cells = diff_cells = 0

    for row_i in range(n_rows):
        key_val = normalize_key(df1.iloc[row_i][key_col_a], strip, icase, key_width)
        candidates = key_index.get(key_val, [])

        if not candidates:
            s2_rows.append(None)
            not_found_count += 1
            row_diffs = list(compare_cols)
            diff_per_row.append(row_diffs)
            diff_cells += len(pairs)
            key_show = _short(df1.iloc[row_i][key_col_a])
            notes.append(
                f"Không tìm thấy KEY [{key_col_a}={key_show}] trong Sheet2 [{key_col_b}]"
            )
            continue

        found_count += 1
        s2_row = _pick_best_row(df1, df2, row_i, candidates, pairs, strip, icase, pair_w)
        s2_rows.append(s2_row)
        row_diffs = []
        note_parts = []

        for ca, cb in pairs:
            w = pair_w[(ca, cb)]
            ok = (
                normalize_key(df1.iloc[row_i][ca], strip, icase, w)
                == normalize_key(df2.iloc[s2_row][cb], strip, icase, w)
            )
            cell_match[ca][row_i] = ok
            if ok:
                same_cells += 1
            else:
                diff_cells += 1
                row_diffs.append(ca)
                v1 = _short(df1.iloc[row_i][ca])
                v2 = _short(df2.iloc[s2_row][cb])
                note_parts.append(f"{ca} ↔ {cb}: '{v1}' ≠ '{v2}'")

        diff_per_row.append(row_diffs)
        notes.append(" | ".join(note_parts))

    result_df = df1.copy()
    for _ca, col_b in pairs:
        result_df[f"{col_b} (Sheet2)"] = [
            df2.iloc[s2][col_b] if s2 is not None else ""
            for s2 in s2_rows
        ]

    statuses = []
    for row_i in range(n_rows):
        if s2_rows[row_i] is None:
            statuses.append("❌ Không tìm thấy KEY")
        elif diff_per_row[row_i]:
            statuses.append("⚠️ Có ô khác")
        else:
            statuses.append("✅ Khớp toàn dòng")
    result_df["_STATUS"] = statuses
    result_df["_NOTE"] = notes
    result_df = result_df.reset_index(drop=True)

    total_cells = same_cells + diff_cells
    pct = round(100 * same_cells / total_cells) if total_cells else 0

    return {
        "version": TOOL_VERSION,
        "pairs": pairs,
        "compare_cols": compare_cols,
        "cell_match": cell_match,
        "diff_per_row": diff_per_row,
        "result_df": result_df,
        "n_rows": n_rows,
        "found_count": found_count,
        "not_found_count": not_found_count,
        "same_cells": same_cells,
        "diff_cells": diff_cells,
        "pct": pct,
    }


def _display(data: dict) -> None:
    result_df = data["result_df"]
    compare_cols = data["compare_cols"]
    cell_match = data["cell_match"]
    n_rows = data["n_rows"]

    show_stats([
        {"num": str(n_rows),                    "lbl": "Dòng Sheet1",      "color": None},
        {"num": str(data["found_count"]),       "lbl": "Tìm thấy KEY",     "color": "#7eb8f7"},
        {"num": str(data["not_found_count"]),   "lbl": "Không có KEY",     "color": "#ff6b6b"},
        {"num": str(data["same_cells"]),        "lbl": "Ô giống",          "color": "#4dbb77"},
        {"num": str(data["diff_cells"]),        "lbl": "Ô khác",           "color": "#ff6b6b"},
        {"num": f"{data['pct']}%",              "lbl": "Tỉ lệ ô khớp",     "color": "#7eb8f7"},
    ])

    preview_n = min(200, n_rows)
    preview = result_df.head(preview_n)

    def _highlight(row):
        idx = int(row.name)
        styles = []
        for col in row.index:
            if col == "_STATUS":
                txt = str(row[col])
                if txt.startswith("✅"):
                    styles.append(STYLE_MATCH)
                elif txt.startswith("⚠️"):
                    styles.append(STYLE_WARN)
                else:
                    styles.append(STYLE_DIFF)
            elif col == "_NOTE":
                styles.append(STYLE_DIFF if str(row[col]).strip() else "")
            elif col in compare_cols:
                ok = cell_match[col][idx] if idx < n_rows else False
                styles.append(STYLE_MATCH if ok else STYLE_DIFF)
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
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "So Sánh Dòng"
        ws.append(list(result_df.columns))
        style_header(ws)

        for row_i in range(n_rows):
            ws.append(list(result_df.iloc[row_i]))
            excel_r = ws.max_row
            for ci, col in enumerate(result_df.columns, start=1):
                cell = ws.cell(excel_r, ci)
                if col in compare_cols:
                    cell.fill = GREEN_FILL if cell_match[col][row_i] else RED_FILL
                elif col.endswith(" (Sheet2)"):
                    cell.fill = BLUE_FILL
                elif col == "_NOTE":
                    if str(cell.value or "").strip():
                        cell.fill = RED_FILL
                elif col == "_STATUS":
                    status = result_df.iloc[row_i]["_STATUS"]
                    if str(status).startswith("✅"):
                        cell.fill = GREEN_FILL
                    else:
                        cell.fill = RED_FILL

        auto_col_width(ws)
        freeze_header(ws)

    excel_download_button(save_workbook_to_bytes(wb), filename="so_sanh_dong.xlsx")
