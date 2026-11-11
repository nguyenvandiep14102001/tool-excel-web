"""
modules/ui_components.py
─────────────────────────
Các Streamlit widget / component dùng lại trong nhiều tool.
"""

import streamlit as st
import pandas as pd
from assets.templates import sheet_info, stat_row, section_divider


# ── CSS loader ────────────────────────────────────────────────────────────────
def load_css() -> None:
    """Đọc file CSS và inject vào app."""
    with open("assets/styles.css", "r", encoding="utf-8") as f:
        css = f.read()
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


# ── Sheet preview (dùng ở đầu mỗi tool) ──────────────────────────────────────
def show_data_preview(df1: pd.DataFrame, df2: pd.DataFrame,
                      name1: str = "Sheet 1", name2: str = "Sheet 2",
                      n_rows: int = 5) -> None:
    """Hiển thị preview 2 sheet song song."""
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**📋 {name1}**")
        st.dataframe(df1.head(n_rows), use_container_width=True, height=200)
        st.markdown(sheet_info(name1, len(df1), len(df1.columns)), unsafe_allow_html=True)
    with c2:
        st.markdown(f"**📋 {name2}**")
        st.dataframe(df2.head(n_rows), use_container_width=True, height=200)
        st.markdown(sheet_info(name2, len(df2), len(df2.columns)), unsafe_allow_html=True)
    st.markdown(section_divider(), unsafe_allow_html=True)


def show_single_preview(df: pd.DataFrame, name: str = "Sheet 1", n_rows: int = 8) -> None:
    """Preview 1 sheet."""
    st.markdown(f"**📋 {name}**")
    st.dataframe(df.head(n_rows), use_container_width=True)
    st.markdown(sheet_info(name, len(df), len(df.columns)), unsafe_allow_html=True)
    st.markdown(section_divider(), unsafe_allow_html=True)


# ── Normalize options ─────────────────────────────────────────────────────────
def normalize_options(key_prefix: str = "norm") -> tuple[bool, bool]:
    """
    Widget cho phép người dùng chọn chuẩn hoá khi so sánh.
    Trả về (strip_whitespace, ignore_case).
    """
    with st.expander("⚙️ Tuỳ chọn chuẩn hoá khi so sánh", expanded=False):
        c1, c2 = st.columns(2)
        strip = c1.checkbox("Bỏ khoảng trắng thừa (trim)", value=True,
                            key=f"{key_prefix}_strip")
        icase = c2.checkbox("Không phân biệt hoa/thường", value=True,
                            key=f"{key_prefix}_case")
    return strip, icase


# ── Column pair selector ──────────────────────────────────────────────────────
def column_pair_selector(cols1: list, cols2: list,
                          label1: str = "Cột Sheet1",
                          label2: str = "Cột Sheet2",
                          key_prefix: str = "pair",
                          n_pairs: int | None = None,
                          default_n: int = 1) -> list[tuple[str, str]]:
    """
    Cho người dùng chọn N cặp cột (col_sheet1, col_sheet2).
    Trả về list of (col1, col2).
    """
    if n_pairs is None:
        n_pairs = st.number_input(
            "Số cặp cột cần so sánh", 1, 15, default_n,
            key=f"{key_prefix}_n",
        )
    pairs = []
    for i in range(int(n_pairs)):
        c1, c2 = st.columns(2)
        ca = c1.selectbox(f"{label1} #{i+1}", cols1,
                          index=min(i, len(cols1)-1), key=f"{key_prefix}_a_{i}")
        cb = c2.selectbox(f"{label2} #{i+1}", cols2,
                          index=min(i, len(cols2)-1), key=f"{key_prefix}_b_{i}")
        pairs.append((ca, cb))
    return pairs


def key_pair_selector(cols1: list, cols2: list,
                      label1: str = "KEY Sheet1",
                      label2: str = "KEY Sheet2",
                      key_prefix: str = "key") -> tuple[str, str]:
    """Chọn một cặp cột làm khóa tra cứu dòng giữa 2 sheet."""
    c1, c2 = st.columns(2)
    ka = c1.selectbox(label1, cols1, index=0, key=f"{key_prefix}_a")
    kb = c2.selectbox(label2, cols2, index=0, key=f"{key_prefix}_b")
    return ka, kb


# ── Stat display ──────────────────────────────────────────────────────────────
def show_stats(stats: list[dict]) -> None:
    """
    Hiển thị stat boxes.
    stats = [{"num": "100", "lbl": "Tổng", "color": None}, ...]
    """
    st.markdown(stat_row(stats), unsafe_allow_html=True)


# ── Download button wrapper ───────────────────────────────────────────────────
def excel_download_button(data: bytes, filename: str,
                           label: str = "⬇️ Tải file Excel kết quả") -> None:
    st.download_button(
        label=label,
        data=data,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# ── Info / warning banners ────────────────────────────────────────────────────
def no_file_warning() -> None:
    st.info("👆 Hãy upload file Excel ở phần **Upload File Excel** bên trên để bắt đầu.")


def tool_header(title: str, subtitle: str) -> None:
    from assets.templates import app_header
    st.markdown(app_header(title, subtitle), unsafe_allow_html=True)


def _build_guide_dialog():
    if not hasattr(st, "dialog"):
        return None
    try:
        decorator = st.dialog("📖 Hướng dẫn sử dụng", width="large")
    except TypeError:
        decorator = st.dialog("📖 Hướng dẫn sử dụng")

    @decorator
    def _dialog(tool_name: str, body: str) -> None:
        st.markdown(f"### {tool_name}")
        st.markdown(body)
        st.caption("Đóng cửa sổ này (✕) để quay lại làm việc.")

    return _dialog


_TOOL_GUIDE_DIALOG = _build_guide_dialog()


def render_tool_header_with_guide(title: str, subtitle: str) -> None:
    """Banner màu + nút Hướng dẫn nằm trong cùng một ô gradient."""
    from assets.templates import app_header
    from assets.tool_guides import get_tool_guide

    left, right = st.columns([5.5, 1.05], gap="small")
    with left:
        st.markdown(app_header(title, subtitle), unsafe_allow_html=True)
    with right:
        st.markdown("<div class='help-btn-wrap'></div>", unsafe_allow_html=True)
        if st.button(
            "❓ Hướng dẫn",
            key="btn_tool_guide",
            use_container_width=True,
            type="primary",
            help="Xem hướng dẫn sử dụng công cụ này",
        ):
            body = get_tool_guide(title)
            if _TOOL_GUIDE_DIALOG is not None:
                _TOOL_GUIDE_DIALOG(title, body)
            else:
                with st.expander(f"📖 Hướng dẫn — {title}", expanded=True):
                    st.markdown(body)
