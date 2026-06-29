"""
modules/excel_utils.py
──────────────────────
Các hàm tiện ích Excel/Pandas dùng chung cho tất cả tool.
"""

from io import BytesIO
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


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
    """Chuẩn hoá giá trị để so sánh."""
    s = str(val).strip() if strip_ws else str(val)
    return s.lower() if ignore_case else s


# ── Excel file loading ────────────────────────────────────────────────────────
def load_excel_file(uploaded_file) -> tuple[pd.ExcelFile, list[str]]:
    """Đọc file Excel, trả về ExcelFile object và danh sách sheet."""
    xl = pd.ExcelFile(uploaded_file)
    return xl, xl.sheet_names


def read_sheet(uploaded_file, sheet_name: str, header_row: int) -> pd.DataFrame:
    """Đọc một sheet cụ thể từ file Excel."""
    return pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)


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
                    strip_ws: bool = True, ignore_case: bool = True) -> dict[str, list[int]]:
    """Map giá trị cột khóa → danh sách chỉ số dòng."""
    index: dict[str, list[int]] = {}
    for i, val in enumerate(df[key_col]):
        k = normalize(val, strip_ws, ignore_case)
        index.setdefault(k, []).append(i)
    return index
