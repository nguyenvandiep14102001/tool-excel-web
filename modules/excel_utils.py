"""
modules/excel_utils.py
──────────────────────
Các hàm tiện ích Excel/Pandas dùng chung cho tất cả tool.
"""

from io import BytesIO
import math
import re

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

_MISSING = {"", "nan", "none", "nat", "<na>", "natype"}
_INT_DOT_ZERO = re.compile(r"^([+-]?\d+)\.0+$")
_SCI_NUM = re.compile(r"^[+-]?\d+(?:\.\d+)?[eE][+-]?\d+$")
_COL_CCCD = re.compile(
    r"cmnd|cccd|cancuoc|căn\s*cước|identity|id[\s_]*card|s[oố0]\s*cmnd|s[oố0]\s*cccd",
    re.I,
)
_COL_PHONE = re.compile(
    r"sdt|phone|mobile|tel|điện\s*thoại|dien\s*thoai|s[oố0]\s*điện|s[oố0]\s*dien",
    re.I,
)
_FMT_ZERO_PAD = re.compile(r"^0+$")


def cell_to_str(val) -> str:
    """
    Đưa ô Excel về chuỗi, giữ số 0 đầu nếu ô vốn là text.
    Sửa artifact khi Excel/pandas đọc CCCD/SĐT như số: 901.0, 9.01E+08.
    """
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(val, bool):
        return str(val)

    if isinstance(val, int):
        return str(val)

    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return ""
        if val.is_integer() and abs(val) < 1e16:
            return str(int(val))
        return str(val)

    s = str(val).strip()
    if s.lower() in _MISSING:
        return ""

    m = _INT_DOT_ZERO.fullmatch(s)
    if m:
        return m.group(1)

    if _SCI_NUM.fullmatch(s):
        try:
            num = float(s)
            if num.is_integer() and abs(num) < 1e16:
                return str(int(num))
        except ValueError:
            pass
    return s


def _pad_digit_str(val, width: int) -> str:
    text = cell_to_str(val)
    if text.isdigit() and 0 < len(text) < width:
        return text.zfill(width)
    return text


def _guess_id_width(series: pd.Series) -> int | None:
    """Đoán độ dài chuẩn: SĐT 10 số, CCCD 12 số, CMND cũ 9 số."""
    lengths = [len(cell_to_str(v)) for v in series if cell_to_str(v).isdigit()]
    if len(lengths) < 3:
        return None
    from collections import Counter
    mode_len, mode_n = Counter(lengths).most_common(1)[0]
    if mode_n / len(lengths) < 0.45:
        return None
    max_len = max(lengths)
    if mode_len == 9 and max_len <= 10:
        return 10
    if mode_len in (10, 11) and max_len <= 12:
        return 12
    if mode_len == 8 and max_len <= 9:
        return 9
    return None


def restore_vn_id_zeros(df: pd.DataFrame) -> pd.DataFrame:
    """
    Bù số 0 đầu cho cột CCCD/CMND (12 số) và SĐT (10 số).
    Excel lưu Number thì 0 đã mất — bù theo tên cột hoặc độ dài phổ biến.
    """
    out = df.copy()
    for col in out.columns:
        name = str(col)
        width = None
        if _COL_CCCD.search(name):
            width = 12
        elif _COL_PHONE.search(name):
            width = 10
        else:
            width = _guess_id_width(out[col])
        if width:
            out[col] = out[col].map(lambda v, w=width: _pad_digit_str(v, w))
    return out


def _excel_cell_str(cell) -> str:
    """Đọc 1 ô openpyxl: giữ text; số thì theo number_format (000…) nếu có."""
    val = cell.value
    if val is None:
        return ""
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, str):
        return cell_to_str(val)
    if isinstance(val, (int, float)):
        fmt = str(cell.number_format or "").strip().split(";")[0].replace('"', "")
        int_fmt = fmt.split(".")[0]
        if isinstance(val, float) and val.is_integer():
            val = int(val)
        if isinstance(val, int) and _FMT_ZERO_PAD.fullmatch(int_fmt):
            return str(val).zfill(len(int_fmt))
        return cell_to_str(val)
    return cell_to_str(val)


def _unique_headers(headers: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out = []
    for i, h in enumerate(headers):
        name = h.strip() if h else f"Unnamed_{i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        out.append(name)
    return out


def digit_key_width(*series_list: pd.Series) -> int:
    """Độ dài KEY dạng số dài nhất — dùng để bù 0 đầu (SĐT 10 số, CCCD 12 số)."""
    width = 0
    for series in series_list:
        if series is None:
            continue
        for val in series:
            text = cell_to_str(val)
            if text.isdigit():
                width = max(width, len(text))
    return width


def normalize_key(val, strip_ws: bool = True, ignore_case: bool = True,
                  width: int = 0) -> str:
    """Chuẩn hoá KEY: giống normalize, thêm zfill nếu là dãy số (CCCD/SĐT)."""
    text = normalize(val, strip_ws, ignore_case)
    if width and text.isdigit():
        return text.zfill(width)
    return text


# ── Fill colors ──────────────────────────────────────────────────────────────
GREEN_FILL  = PatternFill("solid", fgColor="C6EFCE")
RED_FILL    = PatternFill("solid", fgColor="FFC7CE")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
BLUE_FILL   = PatternFill("solid", fgColor="BDD7EE")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)

# Pandas Styler highlight (readable on light & dark UI)
STYLE_MATCH  = "background-color:#dcfce7;color:#166534"
STYLE_DIFF   = "background-color:#fee2e2;color:#991b1b"
STYLE_WARN   = "background-color:#fef3c7;color:#92400e"
STYLE_SHEET2 = "background-color:#dbeafe;color:#1e40af"


# ── Data normalization ────────────────────────────────────────────────────────
def normalize(val, strip_ws: bool = True, ignore_case: bool = True) -> str:
    """Chuẩn hoá giá trị để so sánh — giữ số 0 đầu của CCCD/SĐT dạng text."""
    s = cell_to_str(val)
    if not strip_ws:
        try:
            s = "" if val is None or pd.isna(val) else str(val)
        except (TypeError, ValueError):
            s = str(val)
    return s.lower() if ignore_case else s


# ── Excel file loading ────────────────────────────────────────────────────────
def load_excel_file(uploaded_file) -> tuple[pd.ExcelFile, list[str]]:
    """Đọc file Excel, trả về ExcelFile object và danh sách sheet."""
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    xl = pd.ExcelFile(uploaded_file)
    return xl, xl.sheet_names


def read_sheet(uploaded_file, sheet_name: str, header_row: int) -> pd.DataFrame:
    """
    Đọc sheet bằng openpyxl (ô text giữ nguyên số 0).
    Sau đó bù 0 cho cột CMND/CCCD (12) và SĐT (10).
    """
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    raw = uploaded_file.read() if hasattr(uploaded_file, "read") else uploaded_file
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    try:
        df = _read_sheet_openpyxl(raw, sheet_name, header_row)
    except Exception:
        df = _read_sheet_pandas(uploaded_file, sheet_name, header_row)
    return restore_vn_id_zeros(df)


def _read_sheet_openpyxl(raw, sheet_name: str, header_row: int) -> pd.DataFrame:
    bio = raw if isinstance(raw, BytesIO) else BytesIO(raw)
    wb = openpyxl.load_workbook(bio, data_only=True, read_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Không có sheet {sheet_name}")
    ws = wb[sheet_name]
    rows = [[_excel_cell_str(c) for c in row] for row in ws.iter_rows()]
    wb.close()
    if not rows:
        return pd.DataFrame()
    hdr = min(max(int(header_row), 0), len(rows) - 1)
    headers = _unique_headers(rows[hdr])
    body = rows[hdr + 1:]
    if not body:
        return pd.DataFrame(columns=headers)
    width = len(headers)
    body = [(r + [""] * width)[:width] for r in body]
    return pd.DataFrame(body, columns=headers)


def _read_sheet_pandas(uploaded_file, sheet_name: str, header_row: int) -> pd.DataFrame:
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    df = pd.read_excel(
        uploaded_file,
        sheet_name=sheet_name,
        header=header_row,
        dtype=str,
        keep_default_na=False,
    )
    mapped = df.map(cell_to_str) if hasattr(df, "map") else df.applymap(cell_to_str)
    return mapped


# ── Excel styling ─────────────────────────────────────────────────────────────
def style_header(ws) -> None:
    """Tô màu và bold dòng header của worksheet."""
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


def auto_col_width(ws, max_width: int = 50) -> None:
    """Tự động căn độ rộng cột theo nội dung."""
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, max_width)


def freeze_header(ws) -> None:
    """Freeze dòng header để scroll dễ hơn."""
    ws.freeze_panes = "A2"


# ── Export helpers ────────────────────────────────────────────────────────────
def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """Xuất DataFrame ra bytes của file Excel (không màu)."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        style_header(ws)
        auto_col_width(ws)
        freeze_header(ws)
    return buf.getvalue()


def create_workbook_with_df(df: pd.DataFrame, sheet_name: str = "Kết Quả") -> tuple:
    """
    Tạo workbook + worksheet từ DataFrame.
    Trả về (wb, ws) để caller tô màu thêm trước khi save.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(list(df.columns))
    style_header(ws)
    for _, row in df.iterrows():
        ws.append(list(row))
    auto_col_width(ws)
    freeze_header(ws)
    return wb, ws


def save_workbook_to_bytes(wb) -> bytes:
    """Lưu workbook ra bytes."""
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── DataFrame helpers ─────────────────────────────────────────────────────────
def safe_convert_col(series: pd.Series) -> pd.Series:
    """Chuyển cột về string an toàn, xử lý NaN."""
    return series.fillna("").astype(str)


def col_normalized_set(df: pd.DataFrame, col: str,
                       strip_ws: bool = True, ignore_case: bool = True) -> set:
    """Trả về set giá trị đã normalize của một cột."""
    return {normalize(v, strip_ws, ignore_case) for v in df[col].dropna()}


def build_key_index(df: pd.DataFrame, key_col: str,
                    strip_ws: bool = True, ignore_case: bool = True,
                    width: int = 0) -> dict[str, list[int]]:
    """Map giá trị cột khóa → danh sách chỉ số dòng (bù 0 đầu nếu width > 0)."""
    index: dict[str, list[int]] = {}
    for i, val in enumerate(df[key_col]):
        k = normalize_key(val, strip_ws, ignore_case, width)
        index.setdefault(k, []).append(i)
    return index
