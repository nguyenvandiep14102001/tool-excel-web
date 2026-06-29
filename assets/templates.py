"""
assets/templates.py
────────────────────
Tất cả HTML snippet dùng trong app.
Gọi từ bất kỳ module nào, tránh inline HTML rải rác.
"""


def sidebar_logo() -> str:
    return """
    <div style='text-align:center; padding:16px 0 24px'>
        <div style='font-size:2.5rem'>📊</div>
        <div style='color:#7eb8f7; font-weight:800; font-size:1.1rem'>Excel Analyst</div>
        <div style='color:#4a5568; font-size:.75rem'>Data Tools Suite</div>
    </div>"""


def sidebar_footer() -> str:
    return "<div style='color:#4a5568; font-size:.75rem; text-align:center'>Built with ❤️ for Data Analysts DiepNv</div>"


def sidebar_divider() -> str:
    return "<hr style='border-color:#2a3045'>"


def app_header(title: str, subtitle: str) -> str:
    return f"""
    <div class="app-header">
        <h1>{title}</h1>
        <p>{subtitle}</p>
    </div>"""


def section_divider() -> str:
    return "<hr class='section-divider'>"


def card_open(title: str = "") -> str:
    title_html = f"<div class='card-title'>{title}</div>" if title else ""
    return f"<div class='card'>{title_html}"


def card_close() -> str:
    return "</div>"


def stat_row(stats: list[dict]) -> str:
    """
    stats = [
        {"num": "120",  "lbl": "Tổng dòng",      "color": None},
        {"num": "95",   "lbl": "Khớp",            "color": "#4dbb77"},
        {"num": "25",   "lbl": "Không khớp",      "color": "#ff6b6b"},
    ]
    """
    boxes = ""
    for s in stats:
        color_style = f"color:{s['color']}" if s.get("color") else ""
        boxes += f"""
        <div class='stat-box'>
            <div class='stat-num' style='{color_style}'>{s['num']}</div>
            <div class='stat-lbl'>{s['lbl']}</div>
        </div>"""
    return f"<div class='stat-row'>{boxes}</div>"


def badge(text: str, color: str = "blue") -> str:
    """color: blue | green | red | orange"""
    return f"<span class='badge-{color}'>{text}</span>"


def sheet_info(name: str, rows: int, cols: int) -> str:
    return (
        f"<span class='badge-blue'>{name}</span>&nbsp;"
        f"<span class='badge-blue'>{rows} dòng</span>&nbsp;"
        f"<span class='badge-blue'>{cols} cột</span>"
    )
