"""
app.py  ←  Entry point duy nhất
────────────────────────────────
Chạy:  streamlit run app.py

Cấu trúc project:
    app.py                  ← bạn đang xem (chỉ routing)
    assets/
        styles.css          ← toàn bộ CSS
        templates.py        ← HTML snippet / component
    modules/
        excel_utils.py      ← helper Excel / Pandas
        ui_components.py    ← widget Streamlit tái sử dụng
        file_loader.py      ← logic upload & chọn sheet
    tools/
        compare_columns.py  ← 🔍 So Sánh Cột
        compare_rows.py     ← 📏 So Sánh Dòng
        mapping.py          ← 🔗 Mapping / VLOOKUP
        merge_join.py       ← 🔀 Merge / JOIN
        dedup.py            ← 🧹 Dedup
        normalize_data.py   ← 🔄 Chuẩn Hoá Dữ Liệu
"""

import streamlit as st

# ── Page config (phải gọi TRƯỚC mọi st.* khác) ───────────────────────────────
st.set_page_config(
    page_title="Excel Analyst Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Internal imports ──────────────────────────────────────────────────────────
from modules.ui_components import load_css
from modules.file_loader   import render_file_loader
from assets.templates      import (
    sidebar_logo, sidebar_footer, sidebar_divider,
    app_header,
)

# ── Tool modules ──────────────────────────────────────────────────────────────
from tools import (
    compare_columns,
    compare_rows,
    mapping,
    merge_join,
    dedup,
    normalize_data,
)

# ═══════════════════════════════════════════════════════════
# TOOL REGISTRY  ← thêm tool mới vào đây là đủ
# ═══════════════════════════════════════════════════════════
TOOLS = {
    "🔍 So Sánh Cột":        {
        "module":   compare_columns,
        "subtitle": "So sánh từng cặp cột theo dòng tương ứng — tô màu kết quả",
    },
    "📏 So Sánh Dòng":       {
        "module":   compare_rows,
        "subtitle": "Tra cứu dòng theo KEY, so sánh từng ô — khác ở đâu tô đỏ",
    },
    "🔗 Mapping / VLOOKUP":  {
        "module":   mapping,
        "subtitle": "Tra cứu giá trị từ Sheet2 và gắn vào Sheet1",
    },
    "🔀 Merge / JOIN":        {
        "module":   merge_join,
        "subtitle": "Ghép 2 sheet theo kiểu LEFT / INNER / RIGHT / OUTER",
    },
    "🧹 Dedup":               {
        "module":   dedup,
        "subtitle": "Phát hiện và loại bỏ dòng trùng lặp",
    },
    "🔄 Chuẩn Hoá Dữ Liệu": {
        "module":   normalize_data,
        "subtitle": "Trim, upper/lower, định dạng ngày, số, fill NA…",
    },
}

# ═══════════════════════════════════════════════════════════
# CSS
# ═══════════════════════════════════════════════════════════
load_css()

# ═══════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(sidebar_logo(), unsafe_allow_html=True)

    tool_name = st.radio(
        "🔧 Chọn công cụ",
        list(TOOLS.keys()),
        label_visibility="visible",
        key="selected_tool",
    )

    st.markdown(sidebar_divider(), unsafe_allow_html=True)

    if st.button("🗑️ Clear cache", use_container_width=True, key="btn_clear_cache"):
        from modules.cache_utils import clear_pycache, clear_tool_session_state

        pycache = clear_pycache()
        n_state = clear_tool_session_state(st.session_state)
        st.cache_data.clear()
        st.cache_resource.clear()
        st.session_state["_cache_clear_msg"] = (
            f"Đã xóa {pycache['pycache_dirs']} thư mục __pycache__ "
            f"({pycache['pycache_files']} file) và {n_state} mục session."
        )
        st.rerun()

    if msg := st.session_state.pop("_cache_clear_msg", None):
        st.success(msg)

    st.markdown(sidebar_divider(), unsafe_allow_html=True)
    st.markdown(sidebar_footer(), unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# HEADER
# ═══════════════════════════════════════════════════════════
tool_cfg = TOOLS[tool_name]
st.markdown(
    app_header(tool_name, tool_cfg["subtitle"]),
    unsafe_allow_html=True,
)

# ═══════════════════════════════════════════════════════════
# FILE UPLOAD  (shared, always visible)
# ═══════════════════════════════════════════════════════════
df1, df2, sheet1_name, sheet2_name = render_file_loader()

# ═══════════════════════════════════════════════════════════
# DISPATCH to selected tool
# ═══════════════════════════════════════════════════════════
st.markdown("---")
tool_cfg["module"].render(df1, df2, sheet1_name, sheet2_name)
