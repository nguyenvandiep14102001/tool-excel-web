"""
tools/fuzzy_match.py
────────────────────
Tool: 🔎 Fuzzy Match
Ghép dữ liệu khi giá trị gần giống nhau (không khớp tuyệt đối).
"""

from difflib import SequenceMatcher

import openpyxl
import pandas as pd
import streamlit as st

from modules.excel_utils import (
    normalize,
    GREEN_FILL,
    RED_FILL,
    auto_col_width,
    freeze_header,
    save_workbook_to_bytes,
    style_header,
)
from modules.ui_components import (
    excel_download_button,
    no_file_warning,
    normalize_options,
    show_data_preview,
    show_stats,
)


def _similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def render(df1: pd.DataFrame | None, df2: pd.DataFrame | None,
           name1: str, name2: str) -> None:
    if df1 is None:
        no_file_warning()
        return
    if df2 is None:
        st.warning("⚠️ Fuzzy Match cần cả Sheet 1 và Sheet 2.")
        return

    show_data_preview(df1, df2, name1, name2)
    st.markdown("### 🔎 Fuzzy Match")

    strip, icase = normalize_options("fuzzy")
    cols1, cols2 = list(df1.columns), list(df2.columns)

    c1, c2 = st.columns(2)
    key1 = c1.selectbox(f"KEY {name1}", cols1, key="fuzzy_k1")
    key2 = c2.selectbox(f"KEY {name2}", cols2, key="fuzzy_k2")

    return_cols = st.multiselect("Cột lấy từ Sheet2", cols2, key="fuzzy_ret")
    threshold = st.slider("Ngưỡng khớp (0–1)", 0.5, 1.0, 0.8, 0.05, key="fuzzy_thr")

    if len(df1) > 2000:
        st.warning(f"⚠️ Sheet1 có {len(df1)} dòng — fuzzy match có thể chậm.")

    if st.button("🚀 Chạy Fuzzy Match", use_container_width=True):
        if not return_cols:
            st.warning("Chọn ít nhất 1 cột cần lấy từ Sheet2.")
            return
        _run(df1, df2, key1, key2, return_cols, threshold, strip, icase)


def _run(df1, df2, key1, key2, return_cols, threshold, strip, icase):
    lookup = df2.copy()
    lookup["_norm"] = lookup[key2].apply(lambda x: normalize(x, strip, icase))

    choices = lookup["_norm"].tolist()
    choice_rows = lookup.index.tolist()

    result = df1.copy().astype("object")
    scores: list[float | None] = []
    matched = 0

    for _, row in df1.iterrows():
        q = normalize(row[key1], strip, icase)
        best_score = 0.0
        best_i = None
        for i, c in enumerate(choices):
            if not c and not q:
                s = 1.0
            else:
                s = _similarity(q, c)
            if s > best_score:
                best_score = s
                best_i = i

        if best_i is not None and best_score >= threshold:
            src = lookup.loc[choice_rows[best_i]]
            for rc in return_cols:
                col_name = f"{rc}_fuzzy"
                result.loc[row.name, col_name] = src[rc]
            scores.append(round(best_score, 3))
            matched += 1
        else:
            for rc in return_cols:
                result.loc[row.name, f"{rc}_fuzzy"] = ""
            scores.append(round(best_score, 3) if best_score else None)

    result["_match_score"] = scores

    show_stats([
        {"num": str(len(df1)), "lbl": "Dòng Sheet1", "color": None},
        {"num": str(matched), "lbl": "Khớp fuzzy", "color": "#4dbb77"},
        {"num": str(len(df1) - matched), "lbl": "Không khớp", "color": "#ff6b6b"},
        {"num": str(threshold), "lbl": "Ngưỡng", "color": "#7eb8f7"},
    ])

    st.dataframe(result, use_container_width=True, height=380)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fuzzy Match"
    ws.append(list(result.columns))
    style_header(ws)
    score_idx = list(result.columns).index("_match_score")
    for _, row in result.iterrows():
        ws.append(list(row))
        score = row["_match_score"]
        excel_r = ws.max_row
        fill = GREEN_FILL if score and score >= threshold else RED_FILL
        for ci in range(1, len(result.columns) + 1):
            if ci == score_idx + 1 or result.columns[ci - 1].endswith("_fuzzy"):
                ws.cell(excel_r, ci).fill = fill
    auto_col_width(ws)
    freeze_header(ws)
    excel_download_button(save_workbook_to_bytes(wb), filename="fuzzy_match.xlsx")
