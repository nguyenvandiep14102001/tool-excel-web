"""
assets/tool_guides.py
─────────────────────
Nội dung hướng dẫn sử dụng từng công cụ (hiện trong popup).
Key phải khớp tên tool trong app.py → TOOLS.
"""

TOOL_GUIDES: dict[str, str] = {
    "🔍 So Sánh Cột": """
**Mục đích:** So sánh nhiều **cặp cột** giữa File1 và File2 theo **KEY**.  
Dòng xanh khi **tất cả cặp** khớp.

**Cách dùng**
1. Upload File1 và File2 (hoặc 1 file — 2 sheet)
2. (Tuỳ chọn) Mở **Tuỳ chọn chuẩn hoá**: trim / không phân biệt hoa thường
3. Chọn **KEY Sheet1 ↔ KEY Sheet2** (cột định danh)
4. Chọn số **cặp cột** cần so sánh và từng cặp
5. Bấm **🚀 Chạy So Sánh Cột**

**Kết quả**
- `_RESULT`: ✅ Khớp / ❌ Không khớp
- Cột so sánh: xanh = khớp, đỏ = lệch
- Sheet Excel **`Chỉ có File2`**: dòng có KEY ở File2 nhưng không có ở File1

**Lưu ý**
- Số thập phân như `6.20` và `6.2` được coi là **khớp**
- File xuất: `so_sanh_cot.xlsx`
""",

    "📏 So Sánh Dòng": """
**Mục đích:** Kiểm tra từng **dòng File1** có xuất hiện trong File2 không (so sánh **toàn dòng**, không chọn từng cột).

**Cách dùng**
1. Upload File1 và File2
2. (Tuỳ chọn) Bật chuẩn hoá trim / không phân biệt hoa thường
3. Bấm **🚀 Chạy So Sánh Dòng**

**Kết quả / ghi chú**
| Status | Ý nghĩa |
|--------|---------|
| ✅ Khớp toàn dòng | Tất cả cột khớp với một dòng File2 |
| ⚠️ Khác N cột | Có dòng gần giống — `_NOTE` liệt kê cột lệch |
| ❌ Không có toàn bộ trong File2 | Không tìm thấy dòng tương ứng |

- Sheet Excel **`Chỉ có File2`**: dòng có ở File2 nhưng không có ở File1
- Ưu tiên so theo **tên cột chung**; nếu khác tên thì so theo **thứ tự cột**

**Lưu ý:** `6.20` ≈ `6.2` được coi là khớp. File xuất: `so_sanh_dong.xlsx`
""",

    "🔗 Mapping / VLOOKUP": """
**Mục đích:** Tra cứu giá trị từ Sheet2 và **gắn thêm cột** vào Sheet1 (giống VLOOKUP).

**Cách dùng**
1. Upload 2 sheet
2. Chọn **KEY** Sheet1 và Sheet2
3. Chọn **cột giá trị** cần lấy từ Sheet2 (có thể nhiều cột)
4. Đặt **tên cột đích** trên Sheet1
5. Chọn xử lý khi Sheet2 **trùng KEY**: Đầu tiên / Cuối cùng / Nối bằng dấu phẩy
6. Bấm **🚀 Chạy Mapping**

**Kết quả:** Sheet1 + cột mới; có giá trị → xanh, không tìm thấy KEY → đỏ.  
File xuất: `mapping_result.xlsx`
""",

    "🔀 Merge / JOIN": """
**Mục đích:** Ghép 2 sheet theo kiểu SQL JOIN.

**Cách dùng**
1. Upload 2 sheet
2. Chọn **cột JOIN** (có thể nhiều cột; số cột key 2 bên phải bằng nhau)
3. Chọn kiểu: **LEFT / INNER / RIGHT / OUTER**
4. Đặt suffix cho cột trùng tên (`_s1`, `_s2`)
5. Bấm **🚀 Chạy JOIN**

| Kiểu | Giữ dòng |
|------|----------|
| LEFT | Tất cả File1 |
| INNER | Chỉ dòng khớp cả 2 |
| RIGHT | Tất cả File2 |
| OUTER | Tất cả cả 2 bên |

File xuất: `join_result.xlsx`
""",

    "🧹 Dedup": """
**Mục đích:** Phát hiện và loại bỏ **dòng trùng**.

**Cách dùng**
1. Upload file
2. Chọn **cột xác định trùng** (để trống = so sánh cả dòng)
3. Chọn giữ: **Đầu tiên / Cuối cùng / Xoá tất cả dòng trùng**
4. Chọn sheet cần dedup
5. Bấm **🚀 Chạy Dedup**

**Kết quả:** 2 sheet — dữ liệu sau dedup + các dòng trùng đã loại.  
File xuất: `dedup_result.xlsx`
""",

    "🔄 Chuẩn Hoá Dữ Liệu": """
**Mục đích:** Làm sạch dữ liệu trước khi so sánh / import.

**Cách dùng**
1. Chọn Sheet 1 hoặc Sheet 2
2. Chọn cột áp dụng (để trống = tất cả)
3. Bật thao tác cần thiết:
   - **Text:** trim, hoa/thường, xoá ký tự đặc biệt, xoá xuống dòng
   - **Số & Ngày:** chuẩn hoá số, chuẩn hoá ngày + format đầu ra
   - **Trống & Dedup:** fill NA, xoá dòng trống, dedup nhanh
4. Bấm **🚀 Chuẩn Hoá Ngay**

**Gợi ý quy trình:** Chuẩn Hoá → So Sánh Cột/Dòng.  
File xuất: `chuan_hoa.xlsx`
""",

    "📊 Data Profile": """
**Mục đích:** Báo cáo chất lượng dữ liệu theo từng cột.

**Cách dùng**
1. Chọn sheet
2. Bấm **🚀 Chạy Profile**

**Thông tin mỗi cột:** kiểu gợi ý, null, % null, distinct, trùng, ví dụ, min/max.  
File xuất: `data_profile.xlsx`
""",

    "✅ Validate Rules": """
**Mục đích:** Kiểm tra **rule nghiệp vụ**, liệt kê dòng vi phạm.

**Cách dùng**
1. Chọn sheet
2. Đặt số rule (1–10)
3. Mỗi rule: cột + loại rule + giá trị/pattern  
   (Không trống, >, <, Bằng/Khác, Chứa, Regex, Email…)
4. Bấm **🚀 Chạy Validate**

**Kết quả:** tab dòng lỗi + cột `_FAIL_*`; dòng lỗi tô đỏ trong Excel.  
File xuất: `validate_result.xlsx`
""",

    "✂️ Tách / Trích Cột": """
**Mục đích:** Tách 1 cột thành nhiều cột, hoặc trích bằng regex.

**Chế độ 1 — Delimiter:** chọn cột, dấu tách (vd `,`), số cột tạo ra.  
**Chế độ 2 — Regex:** pattern có group, ví dụ `(\\d+)$`.

Bấm **🚀 Thực hiện** → File xuất: `split_extract.xlsx`
""",

    "📈 GroupBy": """
**Mục đích:** Nhóm dữ liệu và tổng hợp (sum / count / mean / min / max…).

**Cách dùng**
1. Chọn sheet
2. Chọn cột **nhóm theo**
3. Chọn cột giá trị + hàm tổng hợp
4. Bấm **🚀 Tổng Hợp**

File xuất: `groupby_result.xlsx`
""",

    "🧮 Pivot Table": """
**Mục đích:** Tạo bảng pivot nhanh.

**Cách dùng**
1. Chọn sheet
2. Chọn **Rows**, **Columns**, **Values**
3. Chọn aggregation: Sum / Count / Mean / Min / Max
4. Bấm **🚀 Tạo Pivot**

Có dòng/cột **Tổng** (margins). File xuất: `pivot_result.xlsx`
""",

    "🔎 Fuzzy Match": """
**Mục đích:** Ghép khi giá trị **gần giống** (lỗi chính tả, thiếu dấu) — không cần khớp 100%.

**Cách dùng**
1. Upload 2 sheet
2. Chọn KEY 2 sheet + cột lấy từ Sheet2
3. Kéo **ngưỡng khớp** (0.5–1.0, mặc định 0.8)
4. Bấm **🚀 Chạy Fuzzy Match**

**Kết quả:** cột `*_fuzzy` + `_match_score`.  
> Sheet1 > 2000 dòng có thể chậm. File xuất: `fuzzy_match.xlsx`
""",

    "🔽 Filter Builder": """
**Mục đích:** Lọc dòng theo nhiều điều kiện AND/OR.

**Cách dùng**
1. Chọn sheet
2. Đặt số điều kiện (1–8): cột + toán tử + giá trị
3. Chọn **AND** (tất cả) hoặc **OR** (bất kỳ)
4. Bấm **🚀 Lọc Dữ Liệu**

**Toán tử:** Bằng, Khác, Chứa, >, <, Trống…  
File xuất: `filtered_data.xlsx`
""",

    "📐 Unpivot": """
**Mục đích:** Chuyển bảng **wide → long** (melt) cho BI / Power BI.

**Cách dùng**
1. Chọn sheet
2. **id_vars** — cột giữ nguyên
3. **value_vars** — cột cần melt (để trống = các cột còn lại)
4. Đặt tên cột attribute / value
5. Bấm **🚀 Unpivot**

File xuất: `unpivot_result.xlsx`
""",

    "📅 Date Intelligence": """
**Mục đích:** Parse cột ngày và thêm cột phụ cho báo cáo.

**Cách dùng**
1. Chọn sheet + cột ngày
2. Bật cột cần tạo: Year, Quarter, Month, Week, Weekday, Is weekend
3. Đặt format chuẩn hoá ngày
4. Bấm **🚀 Thêm Cột Ngày**

Hỗ trợ chuỗi ngày, serial Excel, datetime.  
File xuất: `date_intelligence.xlsx`
""",

    "🧽 Dọn Trống": """
**Mục đích:** Dọn dòng/cột trống và khoảng trắng thừa.

**Tuỳ chọn**
- Xoá dòng trống hoàn toàn
- Xoá dòng có ô trống
- Xoá cột trống hoàn toàn
- Trim toàn bảng

Bấm **🚀 Dọn Dữ Liệu** → File xuất: `cleaned_data.xlsx`
""",
}


def get_tool_guide(tool_name: str) -> str:
    return TOOL_GUIDES.get(
        tool_name,
        "Chưa có hướng dẫn chi tiết cho công cụ này. "
        "Upload file → cấu hình các tuỳ chọn → bấm nút Chạy trên màn hình.",
    )
