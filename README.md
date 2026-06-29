<<<<<<< HEAD
# 📊 Excel Analyst Tool v2

Công cụ phân tích Excel dành cho Data Analyst — kiến trúc module hoá.

## 🚀 Cách chạy

```bash
# 1. Cài thư viện
pip install -r requirements.txt

# 2. Chạy app
streamlit run app.py
```
Mở trình duyệt: **http://localhost:8501**

---

## 📁 Cấu trúc project

```
excel_analyst_tool/
│
├── app.py                   ← Entry point (chỉ routing, ~70 dòng)
│
├── assets/
│   ├── styles.css           ← Toàn bộ CSS (sửa giao diện ở đây)
│   └── templates.py         ← HTML snippet / component
│
├── modules/
│   ├── excel_utils.py       ← Helper Excel/Pandas (màu sắc, load, export)
│   ├── ui_components.py     ← Widget Streamlit tái sử dụng
│   └── file_loader.py       ← Logic upload file và chọn sheet
│
├── tools/
│   ├── compare_columns.py   ← 🔍 So Sánh Cột
│   ├── compare_rows.py      ← 📏 So Sánh Dòng
│   ├── mapping.py           ← 🔗 Mapping / VLOOKUP
│   ├── merge_join.py        ← 🔀 Merge / JOIN
│   ├── dedup.py             ← 🧹 Dedup
│   └── normalize_data.py    ← 🔄 Chuẩn Hoá Dữ Liệu
│
└── requirements.txt
```

## 🔧 Hướng dẫn thêm tool mới

1. Tạo file `tools/ten_tool_moi.py` với hàm `render(df1, df2, name1, name2)`
2. Import vào `app.py`: `from tools import ten_tool_moi`
3. Thêm vào dict `TOOLS` trong `app.py`:
```python
"🆕 Tên Tool Mới": {
    "module":   ten_tool_moi,
    "subtitle": "Mô tả ngắn",
},
```
Xong! Không cần sửa gì khác.

## 🎨 Sửa giao diện

- **Màu sắc, font, layout** → `assets/styles.css`
- **HTML component** (header, badge, stat box) → `assets/templates.py`
- **Widget Streamlit** tái sử dụng → `modules/ui_components.py`

## 📦 Thư viện

| Thư viện   | Dùng cho                        |
|------------|---------------------------------|
| streamlit  | Web UI                          |
| pandas     | Xử lý DataFrame                 |
| openpyxl   | Đọc/ghi Excel có màu            |
| xlsxwriter | Export Excel (backup)           |
| numpy      | Tính toán                       |
=======
# README #

This README would normally document whatever steps are necessary to get your application up and running.

### What is this repository for? ###

* Quick summary
* Version
* [Learn Markdown](https://bitbucket.org/tutorials/markdowndemo)

### How do I get set up? ###

* Summary of set up
* Configuration
* Dependencies
* Database configuration
* How to run tests
* Deployment instructions

### Contribution guidelines ###

* Writing tests
* Code review
* Other guidelines

### Who do I talk to? ###

* Repo owner or admin
* Other community or team contact
>>>>>>> a2c9ab89b8a4205fdcff647d21e62c41b21f68ef
