# 📊 Excel Analyst Tool — Data Tools Suite

Bộ công cụ phân tích Excel dành cho Data Analyst: so sánh, ghép, chuẩn hoá, lọc, pivot và kiểm tra chất lượng dữ liệu — giao diện web Streamlit, xuất kết quả ra Excel có màu.

---

## 🚀 Cài đặt & chạy

### Yêu cầu

- Python 3.10+
- Windows / macOS / Linux

### Các bước

```bash
# 1. Clone hoặc mở thư mục project
cd excel_tool

# 2. Tạo virtual environment (khuyến nghị)
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate     # macOS / Linux

# 3. Cài thư viện
pip install -r requirements.txt

# 4. Chạy app
streamlit run app.py
```

Trình duyệt mở tại **http://localhost:8501** (tự mở nếu `headless = false` trong `.streamlit/config.toml`).

### Dừng app

Trong terminal: `Ctrl + C`.

---

## 👥 Cho người dùng bình thường (chỉ cần mở app)

### So sánh các cách phát hành

| Cách | Gửi gì cho user? | User làm gì? | Cần source trên máy user? |
|------|------------------|--------------|---------------------------|
| **⭐ Server team (khuyến nghị)** | Chỉ **1 link** hoặc file `.url` | Mở Chrome/Edge, dán link | **Không** |
| Portable zip + venv | 1 file zip (~300MB) | Giải nén → `Mo Excel Analyst.bat` | Có folder app (user không sửa code) |
| Portable zip không venv | Zip nhỏ hơn | `Cai dat lan dau.bat` 1 lần + `Mo Excel Analyst.bat` | Có folder + cần Python |
| File .exe (PyInstaller) | 1 file Setup.exe | Cài / chạy exe | Ẩn source, bundle ~200MB+ |

**Trả lời câu hỏi thường gặp:** Với `.bat` trên máy từng người — **đúng**, bạn phải gửi folder app (hoặc zip). User **không cần đọc/sửa code**, chỉ double-click `.bat`.

**Muốn user không nhận gì cả ngoài link** → dùng **Cách Server** bên dưới.

---

### ⭐ Cách 1 — Một máy chạy server, team chỉ mở trình duyệt (tốt nhất)

**Trên máy chủ** (1 máy mạnh / máy IT, luôn bật khi team dùng):

1. `Cai dat lan dau.bat` (một lần)
2. Double-click **`Chay Server cho Team.bat`**
3. Copy link hiện ra, ví dụ `http://192.168.1.50:8501`
4. Gửi link cho đồng nghiệp **hoặc** chạy:
   ```powershell
   powershell -File scripts\tao_link_trinh_duyet.ps1
   ```
   → tạo `Excel Analyst.url` trên Desktop, gửi file `.url` cho user

**User:** mở link — **không cần** folder source, Python, hay file `.bat`.

> Máy chủ và user phải cùng mạng LAN (hoặc VPN). Firewall Windows có thể hỏi — chọn **Allow**.

---

### Cách 2 — Gửi gói cài đặt (mỗi người 1 máy)

**Bạn (dev/IT) đóng gói một lần:**

```text
scripts\dong_goi_phat_hanh.bat
```

→ Tạo folder `release\Excel Analyst Tool\` — zip và gửi.

- Chọn **copy venv** khi được hỏi → user **không cần cài Python**
- Không copy venv → user chạy `Cai dat lan dau.bat` một lần

Kèm file **`HUONG DAN NGUOI DUNG.md`** trong gói.

---

### Cách 3 — Double-click trên máy đã có project (dev / nội bộ)

Người không biết lập trình **không cần gõ lệnh** — chỉ double-click file `.bat` trên Windows.

#### Lần đầu (IT / người cài đặt — 1 lần duy nhất)

1. Cài **Python 3.10+** từ [python.org](https://www.python.org/downloads/) — tick **Add Python to PATH**
2. Double-click **`Cai dat lan dau.bat`** trong thư mục `excel_tool`
3. Đợi cài xong → trên Desktop sẽ có shortcut **Excel Analyst** (nếu tạo được)

#### Mỗi lần dùng (người dùng cuối)

Double-click **`Mo Excel Analyst.bat`** (hoặc shortcut trên Desktop)

→ Trình duyệt tự mở app. **Không đóng cửa sổ đen** — đó là server đang chạy.

| File | Ai dùng | Khi nào |
|------|---------|---------|
| `Chay Server cho Team.bat` | IT — **1 máy** cho cả team | User chỉ cần link trình duyệt |
| `Cai dat lan dau.bat` | IT / admin | **1 lần** trên mỗi máy (cách 2, 3) |
| `Mo Excel Analyst.bat` | Mọi người | **Mỗi lần** mở app trên máy local |
| `scripts\dong_goi_phat_hanh.bat` | Dev / IT | Đóng gói zip phát hành |
| `scripts\tao_link_trinh_duyet.ps1` | IT | Tạo file `.url` cho user |

### Cách 4 — File .exe (phát hành rộng, chưa có sẵn)

Đóng gói PyInstaller (~200–400 MB), user cài như phần mềm thường — không thấy folder `.py`. Cần thêm thời gian build & test.

### Lưu ý khi hỗ trợ user

- Cửa sổ đen = app đang chạy; đóng cửa sổ = tắt app
- App local: `http://localhost:8501` — không cần Internet
- App server: dùng link LAN IT cung cấp
- Antivirus đôi khi chặn `.bat` lần đầu → chọn **Allow**

---

## 🖥️ Giao diện chung

### Sidebar (thanh bên trái)

| Thành phần | Mô tả |
|------------|--------|
| **Logo / tên app** | Branding — sửa `APP_NAME`, `APP_TAGLINE` ở đầu `app.py` |
| **🔧 Chọn công cụ** | Radio chọn 1 trong 16 tool |
| **🗑️ Clear cache** | Xóa `__pycache__`, session kết quả cũ, cache Streamlit |
| **Theme (menu ⋯)** | System / Light / Dark — góc phải trên |

### Khu vực chính

1. **Tiêu đề tool** — tên + mô tả ngắn
2. **📁 Upload File Excel** — nạp dữ liệu (dùng chung cho mọi tool)
3. **Cấu hình tool** — tuỳ chọn riêng từng công cụ
4. **Nút Chạy…** — thực thi phân tích
5. **Kết quả + ⬇️ Tải Excel** — xem bảng và tải file

### Lưu ý quan trọng

- **Nút "Chạy…"** chỉ chạy khi bạn **bấm** — đổi cấu hình xong phải bấm lại.
- **Rerun** (phím `R`) chạy lại script nhưng **không** tự bấm lại nút Chạy.
- Sau khi sửa code Python: **dừng và chạy lại** `streamlit run app.py`, hoặc bấm **Clear cache** nếu UI cũ.
- Hầu hết tool hỗ trợ **xuất Excel** — cột mới / dòng lỗi thường được **tô màu** (xanh = khớp, đỏ = lỗi/khác, xanh dương = dữ liệu Sheet2).

---

## 📁 Upload dữ liệu

Mở expander **📁 Upload File Excel**:

### Chế độ 1: 2 File Excel riêng biệt

- Upload **File 1** và **File 2** (`.xlsx`, `.xls`)
- Chọn **dòng header** (0 = dòng đầu tiên là tên cột)
- Chọn **sheet** trong mỗi file

### Chế độ 2: 1 File — 2 Sheet

- Upload **một file** có ít nhất 2 sheet
- Chọn **Sheet 1** và **Sheet 2**
- Chọn **dòng header** chung

> **Sheet 1** = dữ liệu chính (bên trái khi preview). **Sheet 2** = dữ liệu đối chiếu / tra cứu.

### Ví dụ upload

**Tình huống:** Đối chiếu danh sách khách hàng từ hệ thống A và hệ thống B.

| Cách | Thao tác |
|------|----------|
| 2 file riêng | `KH_he_thong_A.xlsx` → File 1, `KH_he_thong_B.xlsx` → File 2 |
| 1 file 2 sheet | File `doi_chieu.xlsx` có sheet `DanhSach_A` và `DanhSach_B` → chế độ **1 File — 2 Sheet** |

---

## 💡 Bộ dữ liệu mẫu (dùng thử nhanh)

Bạn có thể tự tạo file Excel nhỏ theo các bảng dưới để thử từng tool.

### Sheet `DonHang` (Sheet 1 — bán hàng)

| MaDH | MaSP | SoLuong | Ngay |
|------|------|---------|------|
| DH001 | SP01 | 10 | 01/03/2025 |
| DH002 | SP02 | 5 | 02/03/2025 |
| DH003 | SP01 | 3 | 15/03/2025 |

### Sheet `SanPham` (Sheet 2 — danh mục)

| MaSP | TenSP | Gia |
|------|-------|-----|
| SP01 | Ao thun | 150000 |
| SP02 | Quan jean | 350000 |
| SP03 | Giay the thao | 890000 |

### Sheet `KhachHang_A` vs `KhachHang_B` (so sánh)

**KhachHang_A**

| MaKH | HoTen | ThanhPho |
|------|-------|----------|
| KH01 | Nguyen Van A | Can Tho |
| KH02 | Tran Thi B | Ha Noi |

**KhachHang_B** *(thứ tự dòng khác, vài giá trị sai)*

| ID | Ten | City |
|----|-----|------|
| KH02 | Tran Thi B | Ha Noi |
| KH01 | Nguyen Van A | Can Tho A |

→ Dùng **So Sánh Cột** với KEY `MaKH`↔`ID`, so cặp `HoTen`↔`Ten`, `ThanhPho`↔`City` → dòng KH01 **đỏ** vì `Can Tho` ≠ `Can Tho A`.

---

## 🔧 Danh sách công cụ & hướng dẫn sử dụng

### 🔍 So Sánh Cột

**Mục đích:** So sánh **nhiều cặp cột** giữa 2 sheet theo **KEY** — dòng xanh khi **tất cả cặp** khớp.

**Khi nào dùng:** Đối chiếu 2 danh sách không cùng thứ tự dòng (ví dụ: cột A↔C, B↔D trên cùng một KEY).

**Cách dùng:**

1. Upload 2 sheet
2. Mở **⚙️ Tuỳ chọn chuẩn hoá** (trim, không phân biệt hoa/thường)
3. Chọn **KEY Sheet1** ↔ **KEY Sheet2** (cột định danh, ví dụ Mã KH)
4. Chọn **số cặp cột** và từng cặp cần so sánh
5. Bấm **🚀 Chạy So Sánh Cột**

**Kết quả:**

| Cột / màu | Ý nghĩa |
|-----------|---------|
| `_RESULT` | ✅ Khớp / ❌ Không khớp (toàn dòng) |
| Cột Sheet1 (so sánh) | Xanh = khớp cặp đó, đỏ = không |
| `TênCột (Sheet2)` | Giá trị tương ứng từ Sheet2 (nền xanh dương) |

**File xuất:** `so_sanh_cot.xlsx`

**Ví dụ:**

```
KEY: MaKH (Sheet1) ↔ ID (Sheet2)
Cặp so sánh: HoTen ↔ Ten | ThanhPho ↔ City

Sheet1          Sheet2 (tra cứu)     _RESULT
KH01 | Can Tho  | Can Tho A           ❌ Không khớp   ← ThanhPho khác
KH02 | Ha Noi   | Ha Noi              ✅ Khớp
```

Cấu hình app: bật **trim** + **không phân biệt hoa/thường** nếu dữ liệu có khoảng trắng thừa hoặc `HA NOI` vs `Ha Noi`.

---

### 📏 So Sánh Dòng

**Mục đích:** Tra cứu dòng Sheet2 theo **KEY**, so sánh **từng ô** — chỉ rõ ô nào khác.

**Khác với So Sánh Cột:** Báo chi tiết từng ô; trạng thái: ✅ Khớp toàn dòng / ⚠️ Có ô khác / ❌ Không tìm thấy KEY.

**Cách dùng:**

1. Upload 2 sheet
2. Chọn **KEY** 2 sheet
3. Chọn các **cặp cột** cần so sánh (có thể chọn nhiều cột)
4. Bấm **🚀 Chạy So Sánh Dòng**

**Kết quả:** Thống kê ô giống/khác, % khớp; tô màu từng ô và cột `_STATUS`.

**File xuất:** `so_sanh_dong.xlsx`

**Ví dụ:**

Cùng dữ liệu `KhachHang_A` / `KhachHang_B` ở trên:

```
Dòng KH01 → _STATUS: ⚠️ Có ô khác
  HoTen: xanh (khớp)
  ThanhPho: đỏ (Can Tho ≠ Can Tho A)

Dòng KH02 → _STATUS: ✅ Khớp toàn dòng
```

Nếu Sheet2 **không có** `KH99` → `_STATUS: ❌ Không tìm thấy KEY`, toàn dòng đỏ.

---

### 🔗 Mapping / VLOOKUP

**Mục đích:** Tra cứu giá trị từ Sheet2 và **gắn thêm cột** vào Sheet1 (tương tự VLOOKUP Excel).

**Cách dùng:**

1. Upload 2 sheet
2. Chọn **cột KEY** Sheet1 và Sheet2
3. Chọn **cột giá trị** cần lấy từ Sheet2 (có thể nhiều cột)
4. Đặt **tên cột đích** trên Sheet1
5. Chọn xử lý khi Sheet2 có **nhiều dòng trùng KEY:** Đầu tiên / Cuối cùng / Nối bằng dấu phẩy
6. Bấm **🚀 Chạy Mapping**

**Kết quả:** Bảng Sheet1 + cột mới; cột mapped tô xanh nếu có giá trị, đỏ nếu không tìm thấy KEY.

**File xuất:** `mapping_result.xlsx`

**Ví dụ:**

Lấy **TenSP** và **Gia** từ `SanPham` gắn vào `DonHang`:

| Cấu hình | Giá trị |
|----------|---------|
| KEY Sheet1 | `MaSP` |
| KEY Sheet2 | `MaSP` |
| Cột lấy | `TenSP`, `Gia` |
| Tên đích | `TenSanPham`, `DonGia` |
| Trùng KEY | Đầu tiên |

**Kết quả (rút gọn):**

| MaDH | MaSP | SoLuong | TenSanPham | DonGia |
|------|------|---------|------------|--------|
| DH001 | SP01 | 10 | Ao thun | 150000 |
| DH002 | SP02 | 5 | Quan jean | 350000 |

Dòng `DH003` có `MaSP=SP01` → lấy giá từ dòng đầu SP01 trong `SanPham`.

---

### 🔀 Merge / JOIN

**Mục đích:** Ghép 2 sheet theo kiểu SQL JOIN.

**Cách dùng:**

1. Upload 2 sheet
2. Chọn **cột JOIN** (key) — có thể **nhiều cột**; số cột key 2 sheet phải **bằng nhau**
3. Chọn kiểu JOIN: **LEFT / INNER / RIGHT / OUTER**
4. Đặt **suffix** cho cột trùng tên (`_s1`, `_s2`)
5. Bấm **🚀 Chạy JOIN**

| Kiểu JOIN | Mô tả |
|-----------|--------|
| LEFT | Giữ tất cả dòng Sheet1 |
| INNER | Chỉ dòng khớp cả 2 sheet |
| RIGHT | Giữ tất cả dòng Sheet2 |
| OUTER | Giữ tất cả dòng cả 2 sheet |

**File xuất:** `join_result.xlsx`

**Ví dụ:**

Ghép `DonHang` **LEFT JOIN** `SanPham` trên `MaSP`:

| MaDH | MaSP | SoLuong | TenSP | Gia |
|------|------|---------|-------|-----|
| DH001 | SP01 | 10 | Ao thun | 150000 |
| DH002 | SP02 | 5 | Quan jean | 350000 |
| DH003 | SP01 | 3 | Ao thun | 150000 |

**INNER JOIN** → chỉ giữ dòng có `MaSP` tồn tại ở cả 2 sheet (3 dòng trên).

**OUTER JOIN** → nếu `SanPham` có `SP03` chưa có đơn → vẫn xuất hiện với cột đơn hàng = NULL.

---

### 🧹 Dedup

**Mục đích:** Phát hiện và loại bỏ **dòng trùng lặp**.

**Cách dùng:**

1. Upload file (cần Sheet 2 nếu dedup Sheet 2)
2. Chọn **cột xác định trùng** (để trống = so sánh toàn bộ dòng)
3. Chọn **giữ dòng nào:** Đầu tiên / Cuối cùng / Xoá tất cả dòng trùng
4. Chọn **sheet** cần dedup
5. Bấm **🚀 Chạy Dedup**

**Kết quả:** 2 tab — dữ liệu sau dedup và dòng trùng đã xoá.

**File xuất:** `dedup_result.xlsx` (2 sheet: Đã Dedup + Dòng Trùng)

**Ví dụ:**

Bảng email đăng ký:

| Email | HoTen |
|-------|-------|
| a@x.com | An |
| b@y.com | Binh |
| a@x.com | An copy |
| A@X.COM | An |

- Cột trùng: `Email` · Giữ: **Dòng đầu tiên** · Bật **không phân biệt hoa/thường**
- **Sau dedup:** 2 dòng (`a@x.com`, `b@y.com`)
- **Dòng trùng:** 2 dòng còn lại (tô đỏ trong Excel)

---

### 🔄 Chuẩn Hoá Dữ Liệu

**Mục đích:** Làm sạch và chuẩn hoá dữ liệu trước khi so sánh / import.

**Cách dùng:**

1. Chọn **Sheet 1** hoặc **Sheet 2**
2. Chọn **cột áp dụng** (để trống = tất cả)
3. Bật các thao tác cần thiết:

**📝 Text**

| Tuỳ chọn | Mô tả |
|----------|--------|
| Trim | Bỏ khoảng trắng đầu/cuối |
| UPPERCASE / lowercase / Title Case | Đổi kiểu chữ |
| Xoá ký tự đặc biệt | Giữ chữ, số và ký tự trong danh sách (mặc định ` .,@-_/`) |
| Xoá xuống dòng trong ô | Thay `\n` bằng khoảng trắng |

**🔢 Số & Ngày**

| Tuỳ chọn | Mô tả |
|----------|--------|
| Chuẩn hoá số | Bỏ dấu phẩy, khoảng trắng trong số |
| Chuẩn hoá ngày | Hỗ trợ chuỗi, Excel serial, Timestamp |
| Tự nhận diện format | Bật: chỉ parse chuỗi có `/`, `-`, `.` |
| Format đầu ra | `%d/%m/%Y`, `%Y-%m-%d`, … |

**Thứ tự xử lý:** trim → **ngày** → số → text (case/xoá ký tự) → fill NA → xoá dòng trống → dedup.

**🩹 Giá trị trống & Dedup**

- Điền ô trống bằng giá trị (mặc định `N/A`)
- Xoá dòng có ô trống (chọn cột kiểm tra)
- Dedup nhanh theo cột (first/last)

4. Bấm **🚀 Chuẩn Hoá Ngay**

**File xuất:** `chuan_hoa.xlsx`

**Ví dụ:**

| Trước | Thao tác | Sau |
|-------|----------|-----|
| `"  Ha Noi  "` | Trim + Title Case | `Ha Noi` |
| `"1,234,567"` | Chuẩn hoá số | `1234567` |
| `01-03-2025` | Chuẩn hoá ngày → `%Y-%m-%d` | `2025-03-01` |
| `"ABC\nXYZ"` | Xoá xuống dòng | `ABC XYZ` |
| Ô trống | Fill NA = `N/A` | `N/A` |

**Quy trình gợi ý:** Chuẩn Hoá → So Sánh Cột/Dòng (tránh lỗi do khoảng trắng, format ngày khác nhau).

---

### 📊 Data Profile

**Mục đích:** Báo cáo **chất lượng dữ liệu** theo từng cột.

**Cách dùng:**

1. Chọn sheet
2. Bấm **🚀 Chạy Profile**

**Thông tin mỗi cột:** Kiểu gợi ý, số null, % null, distinct, trùng lặp, ví dụ giá trị, min/max (số/ngày).

**File xuất:** `data_profile.xlsx`

**Ví dụ:**

Với cột `Email` có 100 dòng, 5 ô trống, 95 giá trị nhưng chỉ 80 distinct → Profile báo:

| Cột | Null | % Null | Distinct | Kiểu gợi ý |
|-----|------|--------|----------|------------|
| Email | 5 | 5.0 | 80 | email-like |

→ Phát hiện nhanh cột có nhiều null hoặc trùng lặp bất thường trước khi import DB.

---

### ✅ Validate Rules

**Mục đích:** Kiểm tra **rule nghiệp vụ**, liệt kê dòng vi phạm.

**Cách dùng:**

1. Chọn sheet
2. Đặt **số rule** (1–10)
3. Mỗi rule: chọn **cột**, **loại rule**, **giá trị/pattern**

| Rule | Mô tả |
|------|--------|
| Không được trống | Ô null hoặc rỗng |
| Lớn hơn / Nhỏ hơn | So sánh số |
| Bằng / Khác | So sánh text |
| Chứa text | Chuỗi con |
| Regex khớp | Pattern regex |
| Email hợp lệ | Định dạng email |

4. Bấm **🚀 Chạy Validate**

**Kết quả:** Tab dòng lỗi + tab toàn bộ với cột `_FAIL_*`; dòng lỗi tô đỏ trong Excel.

**File xuất:** `validate_result.xlsx`

**Ví dụ:**

Bảng `DonHang` — 3 rule:

| Rule | Cột | Loại | Giá trị |
|------|-----|------|---------|
| 1 | SoLuong | Lớn hơn | 0 |
| 2 | MaSP | Không được trống | — |
| 3 | Email | Email hợp lệ | — |

Dòng `SoLuong = 0` hoặc `MaSP` rỗng → xuất hiện tab **Dòng lỗi**, tô đỏ trong Excel.

---

### ✂️ Tách / Trích Cột

**Mục đích:** Tách 1 cột thành nhiều cột hoặc trích phần bằng regex.

**Chế độ 1 — Tách theo delimiter**

- Chọn cột nguồn, delimiter (ví dụ `,`), số cột tạo ra, tiền tố tên cột mới

**Chế độ 2 — Trích bằng Regex**

- Nhập pattern có group capture, ví dụ `(\d+)` để lấy số

Bấm **🚀 Thực hiện** → **File xuất:** `split_extract.xlsx`

**Ví dụ tách delimiter:**

Cột `HoTen`: `Nguyen,Van,A` · Delimiter: `,` · 3 cột → `HoTen_1`, `HoTen_2`, `HoTen_3`

| HoTen | HoTen_1 | HoTen_2 | HoTen_3 |
|-------|---------|---------|---------|
| Nguyen,Van,A | Nguyen | Van | A |

**Ví dụ regex:**

Cột `MaHD`: `HD-2025-00123` · Pattern: `(\d+)$` · Tên cột: `SoThuTu`

→ `SoThuTu` = `00123`

---

### 📈 GroupBy

**Mục đích:** Nhóm dữ liệu và **tổng hợp** (tương tự pivot đơn giản).

**Cách dùng:**

1. Chọn sheet
2. **Nhóm theo cột** (multiselect)
3. Chọn **cột giá trị** và **hàm:** count / sum / mean / min / max / distinct count
4. Bấm **🚀 Tổng Hợp**

**File xuất:** `groupby_result.xlsx`

**Ví dụ:**

Từ `DonHang`: nhóm theo `MaSP`, cột giá trị `SoLuong`, hàm **Tổng (sum)**

| MaSP | SoLuong_sum |
|------|-------------|
| SP01 | 13 |
| SP02 | 5 |

→ Tổng số lượng bán theo từng mã sản phẩm.

---

### 🧮 Pivot Table

**Mục đích:** Tạo **bảng pivot** nhanh.

**Cách dùng:**

1. Chọn sheet
2. **Rows (index)** — hàng pivot
3. **Columns** — cột pivot
4. **Values** — giá trị tính toán
5. **Aggregation:** Sum / Count / Mean / Min / Max
6. Bấm **🚀 Tạo Pivot**

Có dòng/cột **Tổng** (margins).

**File xuất:** `pivot_result.xlsx`

**Ví dụ:**

Thêm cột `Thang` = 3 vào `DonHang`, rồi pivot:

| Cấu hình | Giá trị |
|----------|---------|
| Rows | `MaSP` |
| Columns | `Thang` |
| Values | `SoLuong` |
| Aggregation | Sum |

| MaSP | 3 | Tổng |
|------|---|------|
| SP01 | 13 | 13 |
| SP02 | 5 | 5 |

---

### 🔎 Fuzzy Match

**Mục đích:** Ghép dữ liệu khi giá trị **gần giống** nhau (lỗi chính tả, thiếu dấu…) — không cần khớp 100%.

**Cách dùng:**

1. Upload 2 sheet
2. Chọn **KEY** 2 sheet
3. Chọn **cột lấy từ Sheet2**
4. Kéo **ngưỡng khớp** (0.5–1.0, mặc định 0.8)
5. Bấm **🚀 Chạy Fuzzy Match**

> Cảnh báo nếu Sheet1 > 2000 dòng — có thể chậm.

**Kết quả:** Cột `*_fuzzy` + `_match_score` (điểm tương đồng 0–1).

**File xuất:** `fuzzy_match.xlsx`

**Ví dụ:**

Sheet1 (tên chưa chuẩn) · Sheet2 (tên chuẩn):

| KEY (S1) | Ten (S1) | KEY (S2) | Ten (S2) |
|----------|----------|----------|----------|
| 1 | Cong ty CP ABC | 1 | Công ty CP ABC |
| 2 | Tap doan XYZ | 2 | Tập đoàn XYZ |

KEY: cột ID · Lấy `Ten` từ S2 · Ngưỡng **0.8**

→ `Cong ty CP ABC` khớp fuzzy với `Công ty CP ABC`, `_match_score` ≈ 0.85.

---

### 🔽 Filter Builder

**Mục đích:** Lọc dòng theo **nhiều điều kiện** AND/OR.

**Cách dùng:**

1. Chọn sheet
2. Đặt **số điều kiện** (1–8)
3. Mỗi điều kiện: cột + toán tử + giá trị
4. Chọn **AND** (tất cả) hoặc **OR** (bất kỳ)
5. Bấm **🚀 Lọc Dữ Liệu**

**Toán tử:** Bằng, Khác, Chứa, Không chứa, >, <, >=, <=, Trống, Không trống.

**File xuất:** `filtered_data.xlsx`

**Ví dụ:**

Lọc `DonHang` — **AND**, 2 điều kiện:

| # | Cột | Điều kiện | Giá trị |
|---|-----|-----------|---------|
| 1 | SoLuong | Lớn hơn | 5 |
| 2 | MaSP | Bằng | SP01 |

→ Chỉ giữ `DH001` (SoLuong=10, MaSP=SP01).

**OR:** `MaSP = SP02` **OR** `SoLuong > 8` → `DH001` và `DH002`.

---

### 📐 Unpivot (Melt)

**Mục đích:** Chuyển bảng **wide → long** (chuẩn bị cho Power BI / phân tích).

**Cách dùng:**

1. Chọn sheet
2. **id_vars** — cột giữ nguyên (ví dụ: Mã, Tên)
3. **value_vars** — cột cần melt (để trống = các cột còn lại)
4. Đặt tên cột attribute và value
5. Bấm **🚀 Unpivot**

**File xuất:** `unpivot_result.xlsx`

**Ví dụ:**

Bảng doanh thu theo tháng (wide):

| MaCH | T1 | T2 | T3 |
|------|----|----|-----|
| CH01 | 100 | 120 | 90 |

`id_vars` = `MaCH` · melt `T1`, `T2`, `T3` · `attribute` = `Thang` · `value` = `DoanhThu`

| MaCH | Thang | DoanhThu |
|------|-------|----------|
| CH01 | T1 | 100 |
| CH01 | T2 | 120 |
| CH01 | T3 | 90 |

→ Dạng long phù hợp import Power BI / vẽ biểu đồ.

---

### 📅 Date Intelligence

**Mục đích:** Parse cột ngày và **thêm cột phụ** phục vụ báo cáo.

**Cách dùng:**

1. Chọn sheet
2. Chọn **cột ngày**
3. Bật cột cần tạo: Year, Quarter, Month, Week, Weekday, Is weekend
4. Đặt **format chuẩn hoá ngày** (mặc định `%Y-%m-%d`)
5. Bấm **🚀 Thêm Cột Ngày**

Hỗ trợ: chuỗi ngày, số serial Excel, datetime.

**File xuất:** `date_intelligence.xlsx`

**Ví dụ:**

Cột `Ngay` = `15/03/2025` · Bật Year, Quarter, Month, Weekday, Is weekend

| Ngay | Ngay_date | Ngay_year | Ngay_quarter | Ngay_month | Ngay_weekday | Ngay_is_weekend |
|------|-----------|-----------|--------------|------------|--------------|-----------------|
| 15/03/2025 | 2025-03-15 | 2025 | 1 | 3 | Saturday | True |

---

### 🧽 Dọn Trống

**Mục đích:** Dọn dữ liệu trống / khoảng trắng thừa.

**Tuỳ chọn:**

| Tuỳ chọn | Mô tả |
|----------|--------|
| Xoá dòng trống hoàn toàn | Mọi ô trong dòng đều rỗng |
| Xoá dòng có ô trống | Bất kỳ ô trống / tất cả ô trống |
| Xoá cột trống hoàn toàn | Cột không có dữ liệu |
| Trim toàn bảng | Bỏ khoảng trắng đầu/cuối mọi ô |

Bấm **🚀 Dọn Dữ Liệu** → **File xuất:** `cleaned_data.xlsx`

**Ví dụ:**

| Trước | Sau (trim + xoá dòng trống + xoá cột trống) |
|-------|---------------------------------------------|
| Dòng 1: `An`, ` `, `HN` | Giữ nguyên (có dữ liệu) |
| Dòng 2: ` `, ` `, ` ` | **Xoá** (dòng trống hoàn toàn) |
| Cột `GhiChu` toàn ô trống | **Xoá cột** |
| `"  Binh  "` | `"Binh"` sau trim |

---

## 🔄 Ví dụ quy trình làm việc (end-to-end)

**Bài toán:** Đối chiếu danh sách khách hàng từ 2 nguồn, báo cáo lỗi, xuất file sạch.

```
Bước 1 — Upload
  File doi_chieu.xlsx · sheet KhachHang_A + KhachHang_B

Bước 2 — 🔄 Chuẩn Hoá Dữ Liệu (từng sheet)
  Trim + Chuẩn hoá ngày + UPPERCASE cột ThanhPho/City

Bước 3 — 📊 Data Profile
  Kiểm tra % null, distinct trên cột KEY

Bước 4 — 🔍 So Sánh Cột
  KEY: MaKH ↔ ID
  Cặp: HoTen↔Ten, ThanhPho↔City
  → Tải so_sanh_cot.xlsx gửi nghiệp vụ

Bước 5 — (Tuỳ chọn) 📏 So Sánh Dòng
  Xem chi tiết ô nào sai trên từng dòng

Bước 6 — (Tuỳ chọn) ✅ Validate Rules
  Rule: MaKH không trống, Email hợp lệ
```

**Bài toán khác:** Ghép đơn hàng + giá sản phẩm

```
DonHang + SanPham
  → 🔗 Mapping (lấy TenSP, Gia)   HOẶC
  → 🔀 LEFT JOIN trên MaSP
  → 📈 GroupBy sum SoLuong theo MaSP
  → 🧮 Pivot: MaSP × Thang
```

---

## 🎨 Tuỳ chỉnh giao diện

| Muốn sửa | File |
|----------|------|
| Màu, font, layout | `assets/styles.css` |
| HTML component (header, badge, stat) | `assets/templates.py` |
| Tên app sidebar | `APP_NAME`, `APP_TAGLINE` trong `app.py` |
| Widget tái sử dụng | `modules/ui_components.py` |
| Cấu hình Streamlit | `.streamlit/config.toml` |

---

## 📁 Cấu trúc project

```
excel_tool/
├── app.py                      # Entry point — routing & registry tool
├── requirements.txt
├── .streamlit/
│   └── config.toml
├── assets/
│   ├── styles.css
│   └── templates.py
├── modules/
│   ├── excel_utils.py          # Đọc/ghi Excel, normalize, màu sắc
│   ├── ui_components.py        # Widget Streamlit dùng chung
│   ├── file_loader.py          # Upload & chọn sheet
│   └── cache_utils.py          # Clear cache / session
└── tools/
    ├── compare_columns.py
    ├── compare_rows.py
    ├── mapping.py
    ├── merge_join.py
    ├── dedup.py
    ├── normalize_data.py
    ├── data_profile.py
    ├── validate_rules.py
    ├── split_extract.py
    ├── groupby_agg.py
    ├── pivot_table.py
    ├── fuzzy_match.py
    ├── filter_builder.py
    ├── unpivot.py
    ├── date_intelligence.py
    └── clean_blank.py
```

---

## ➕ Thêm tool mới

1. Tạo `tools/ten_tool_moi.py` với hàm:

```python
def render(df1, df2, name1, name2) -> None:
    ...
```

2. Import trong `app.py` và thêm vào dict `TOOLS`:

```python
"🆕 Tên Tool": {
    "module":   ten_tool_moi,
    "subtitle": "Mô tả ngắn",
},
```

Không cần sửa file khác.

---

## 📦 Thư viện

| Thư viện | Dùng cho |
|----------|----------|
| streamlit | Giao diện web |
| pandas | Xử lý DataFrame |
| openpyxl | Đọc/ghi Excel có định dạng màu |
| xlsxwriter | Export Excel (dự phòng) |
| numpy | Tính toán số |

---

## ❓ Xử lý sự cố thường gặp

| Vấn đề | Cách xử lý |
|--------|------------|
| Sửa code / text mà UI không đổi | `Ctrl+C` → `streamlit run app.py` lại; hoặc **Clear cache** |
| Kết quả cũ vẫn hiện | Bấm lại nút **Chạy…** trong tool |
| Tool so sánh báo hết hạn phiên bản | Bấm **Chạy So Sánh** lại |
| Light/Dark theme lệch màu | Dùng menu ⋯ chọn theme; CSS dark chỉ áp khi theme Dark |
| File chỉ 1 sheet | Dùng chế độ **2 File Excel** hoặc thêm sheet vào file |

---

## 👤 Tác giả

Built with ❤️ for Data Analysts — **DiepNv**
