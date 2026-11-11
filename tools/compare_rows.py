"""
tools/compare_rows.py
──────────────────────
Tool: 📏 So Sánh Dòng
Kiểm tra từng dòng File1 có xuất hiện (toàn dòng) trong File2 hay không.
Không chọn từng cột — so sánh toàn bộ giá trị dòng.
- Khớp hết → ✅
- Gần khớp (sai 1 vài cột) → ⚠️ ghi chú cột khác
- Không tìm thấy dòng tương ứng → ❌ không có toàn bộ dữ liệu các cột
Sheet output thêm: dòng có trong File2 mà không có trong File1.
"""

from collections import defaultdict

import streamlit as st
import pandas as pd
import openpyxl

from modules.excel_utils import (
    normalize_key, digit_key_width,
    GREEN_FILL, RED_FILL, YELLOW_FILL,
    STYLE_MATCH, STYLE_DIFF, STYLE_WARN,
    style_header, auto_col_width, freeze_header,
    save_workbook_to_bytes, append_df_to_workbook,
)
from modules.ui_components import (
    show_data_preview, normalize_options, show_stats,
    excel_download_button, no_file_warning,
)

TOOL_VERSION = "3.2"
SESSION_KEY = "row_cmp_v3_result"


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return

    show_data_preview(df1, df2, name1, name2)
    st.markdown("### ⚙️ Cấu hình so sánh dòng")
    st.caption(
        f"Logic phiên bản **{TOOL_VERSION}** — so sánh toàn dòng File1 vs File2; "
        "nếu gần khớp sẽ ghi chú cột khác, nếu không có dòng tương ứng thì báo "
        "không có toàn bộ dữ liệu các cột."
    )

    if df2 is None:
        st.warning("⚠️ Cần có Sheet 2 để so sánh dòng.")
        return

    strip, icase = normalize_options("row_cmp")

    cols1, cols2, align_mode = _resolve_compare_columns(df1, df2)
    if not cols1:
        st.error("Không có cột để so sánh giữa 2 sheet.")
        return

    if align_mode == "name":
        st.info(
            f"📌 So sánh toàn dòng theo **tên cột chung** "
            f"({len(cols1)} cột): {', '.join(f'**{c}**' for c in cols1)}"
        )
    else:
        st.info(
            f"📌 Hai sheet khác tên cột — so sánh theo **thứ tự cột** "
            f"({len(cols1)} cột đầu)."
        )

    if st.button("🚀 Chạy So Sánh Dòng", use_container_width=True):
        try:
            with st.spinner("Đang so sánh — file lớn có thể mất vài chục giây..."):
                st.session_state[SESSION_KEY] = _compute(
                    df1, df2, cols1, cols2, strip, icase,
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


def _short(val, max_len: int = 60) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    text = str(val).replace("\n", " ").strip()
    return text[:max_len] + "…" if len(text) > max_len else text


def _resolve_compare_columns(
    df1: pd.DataFrame, df2: pd.DataFrame
) -> tuple[list[str], list[str], str]:
    """
    Ưu tiên cột trùng tên (theo thứ tự File1).
    Nếu không trùng tên → so sánh theo vị trí cột.
    """
    set2 = set(df2.columns)
    common = [c for c in df1.columns if c in set2]
    if common:
        return common, common, "name"
    n = min(len(df1.columns), len(df2.columns))
    return list(df1.columns)[:n], list(df2.columns)[:n], "pos"


def _row_signature(df: pd.DataFrame, row_i: int, cols: list[str],
                   strip: bool, icase: bool, widths: list[int]) -> tuple:
    return tuple(
        normalize_key(df.iloc[row_i][col], strip, icase, widths[j])
        for j, col in enumerate(cols)
    )


def _build_inv_index(sigs: list[tuple]) -> dict[tuple[int, str], set[int]]:
    inv: dict[tuple[int, str], set[int]] = defaultdict(set)
    for i, sig in enumerate(sigs):
        for j, val in enumerate(sig):
            inv[(j, val)].add(i)
    return inv


def _best_candidate(sig1: tuple, sigs_f2: list[tuple],
                    inv: dict[tuple[int, str], set[int]]) -> tuple[int | None, int]:
    """Tìm dòng File2 khớp nhiều cột nhất với sig1. Trả về (index, score)."""
    n_cols = len(sig1)
    candidates: set[int] = set()
    for j, val in enumerate(sig1):
        candidates |= inv.get((j, val), set())

    if not candidates:
        return None, 0

    best_i = None
    best_score = -1
    for i in candidates:
        sig2 = sigs_f2[i]
        score = sum(a == b for a, b in zip(sig1, sig2))
        if score > best_score:
            best_score = score
            best_i = i
            if score == n_cols:
                break
    return best_i, max(best_score, 0)


def _diff_note(df1, df2, row_i, s2_row, cols1, cols2, sig1, sig2) -> str:
    parts = []
    diff_names = []
    for j, (c1, c2) in enumerate(zip(cols1, cols2)):
        if sig1[j] == sig2[j]:
            continue
        label = c1 if c1 == c2 else f"{c1}↔{c2}"
        diff_names.append(label)
        v1 = _short(df1.iloc[row_i][c1])
        v2 = _short(df2.iloc[s2_row][c2])
        parts.append(f"{label}: '{v1}' ≠ '{v2}'")
    head = f"Khác cột: {', '.join(diff_names)}"
    detail = " | ".join(parts)
    return f"{head} — {detail}" if detail else head


def _compute(df1, df2, cols1, cols2, strip, icase) -> dict:
    widths = [
        digit_key_width(df1[c1], df2[c2])
        for c1, c2 in zip(cols1, cols2)
    ]
    n_cols = len(cols1)
    n_rows = len(df1)

    sigs_f2 = [
        _row_signature(df2, i, cols2, strip, icase, widths)
        for i in range(len(df2))
    ]
    inv = _build_inv_index(sigs_f2)

    statuses: list[str] = []
    notes: list[str] = []
    kinds: list[str] = []  # match | partial | missing
    cell_ok: dict[str, list[bool | None]] = {c: [None] * n_rows for c in cols1}
    found_count = partial_count = missing_count = 0

    for row_i in range(n_rows):
        sig1 = _row_signature(df1, row_i, cols1, strip, icase, widths)
        best_i, score = _best_candidate(sig1, sigs_f2, inv)

        if best_i is not None and score == n_cols:
            found_count += 1
            kinds.append("match")
            statuses.append("✅ Khớp toàn dòng")
            notes.append("")
            for c in cols1:
                cell_ok[c][row_i] = True
            continue

        if best_i is not None and score > 0:
            partial_count += 1
            kinds.append("partial")
            sig2 = sigs_f2[best_i]
            diff_cols = [
                cols1[j] for j in range(n_cols) if sig1[j] != sig2[j]
            ]
            statuses.append(f"⚠️ Khác {len(diff_cols)} cột")
            notes.append(
                _diff_note(df1, df2, row_i, best_i, cols1, cols2, sig1, sig2)
            )
            for j, c in enumerate(cols1):
                cell_ok[c][row_i] = sig1[j] == sig2[j]
            continue

        missing_count += 1
        kinds.append("missing")
        statuses.append("❌ Không có toàn bộ trong File2")
        notes.append("Không có toàn bộ dữ liệu các cột trong File2")
        for c in cols1:
            cell_ok[c][row_i] = False

    result_df = df1.copy().reset_index(drop=True)
    result_df["_STATUS"] = statuses
    result_df["_NOTE"] = notes

    sigs_f1 = {
        _row_signature(df1, i, cols1, strip, icase, widths)
        for i in range(len(df1))
    }
    only_mask = [sig not in sigs_f1 for sig in sigs_f2]
    only_file2_df = df2.loc[only_mask].copy().reset_index(drop=True)
    if not only_file2_df.empty:
        only_file2_df.insert(
            0, "_NOTE", "Có trong File2 — không có dòng tương ứng ở File1"
        )

    return {
        "version": TOOL_VERSION,
        "cols1": cols1,
        "cols2": cols2,
        "result_df": result_df,
        "only_file2_df": only_file2_df,
        "kinds": kinds,
        "cell_ok": cell_ok,
        "n_rows": n_rows,
        "found_count": found_count,
        "partial_count": partial_count,
        "missing_count": missing_count,
        "n_only_file2": len(only_file2_df),
    }


def _display(data: dict) -> None:
    result_df = data["result_df"]
    only_file2_df = data["only_file2_df"]
    kinds = data["kinds"]
    cell_ok = data["cell_ok"]
    cols1 = data["cols1"]
    n_rows = data["n_rows"]

    show_stats([
        {"num": str(n_rows),                  "lbl": "Dòng File1",        "color": None},
        {"num": str(data["found_count"]),     "lbl": "Khớp toàn dòng",    "color": "#4dbb77"},
        {"num": str(data["partial_count"]),   "lbl": "Khác một số cột",   "color": "#eab308"},
        {"num": str(data["missing_count"]),   "lbl": "Không có toàn bộ",  "color": "#ff6b6b"},
        {"num": str(data["n_only_file2"]),    "lbl": "Chỉ có File2",      "color": "#7eb8f7"},
    ])

    preview_n = min(200, n_rows)
    preview = result_df.head(preview_n)

    def _highlight(row):
        idx = int(row.name)
        kind = kinds[idx] if idx < n_rows else "missing"
        styles = []
        for col in row.index:
            if col == "_STATUS":
                if kind == "match":
                    styles.append(STYLE_MATCH)
                elif kind == "partial":
                    styles.append(STYLE_WARN)
                else:
                    styles.append(STYLE_DIFF)
            elif col == "_NOTE":
                styles.append(STYLE_DIFF if str(row[col]).strip() else "")
            elif col in cols1:
                ok = cell_ok[col][idx]
                if ok is True:
                    styles.append(STYLE_MATCH)
                elif ok is False:
                    styles.append(STYLE_DIFF)
                else:
                    styles.append("")
            else:
                if kind == "match":
                    styles.append(STYLE_MATCH)
                elif kind == "partial":
                    styles.append("")
                else:
                    styles.append(STYLE_DIFF)
        return styles

    st.markdown("#### Kết quả theo File1 (gốc)")
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

    st.markdown("#### Dòng chỉ có ở File2 (không có trong File1)")
    if only_file2_df.empty:
        st.success("Không có dòng nào chỉ xuất hiện ở File2.")
    else:
        st.dataframe(
            only_file2_df.head(200),
            use_container_width=True,
            height=280,
        )
        if len(only_file2_df) > 200:
            st.caption(
                f"Đang xem 200/{len(only_file2_df)} dòng đầu. Tải Excel để xem toàn bộ."
            )

    with st.spinner("Đang tạo file Excel..."):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "So Sánh Dòng"
        ws.append(list(result_df.columns))
        style_header(ws)

        for row_i in range(n_rows):
            ws.append(list(result_df.iloc[row_i]))
            excel_r = ws.max_row
            kind = kinds[row_i]
            for ci, col in enumerate(result_df.columns, start=1):
                cell = ws.cell(excel_r, ci)
                if col == "_NOTE":
                    if str(cell.value or "").strip():
                        cell.fill = RED_FILL if kind == "missing" else YELLOW_FILL
                elif col == "_STATUS":
                    if kind == "match":
                        cell.fill = GREEN_FILL
                    elif kind == "partial":
                        cell.fill = YELLOW_FILL
                    else:
                        cell.fill = RED_FILL
                elif col in cols1:
                    ok = cell_ok[col][row_i]
                    if ok is True:
                        cell.fill = GREEN_FILL
                    elif ok is False:
                        cell.fill = RED_FILL
                else:
                    if kind == "match":
                        cell.fill = GREEN_FILL
                    elif kind == "missing":
                        cell.fill = RED_FILL

        auto_col_width(ws)
        freeze_header(ws)
        append_df_to_workbook(wb, only_file2_df, "Chỉ có File2")

    excel_download_button(save_workbook_to_bytes(wb), filename="so_sanh_dong.xlsx")
