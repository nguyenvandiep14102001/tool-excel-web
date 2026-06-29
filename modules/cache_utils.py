"""
modules/cache_utils.py
──────────────────────
Xóa __pycache__ và reset cache ứng dụng.
"""

from __future__ import annotations

import shutil
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def clear_pycache(root: Path | None = None, *, include_venv: bool = False) -> dict:
    """
    Xóa toàn bộ thư mục __pycache__ trong project.
    Mặc định bỏ qua venv (chỉ xóa code app: tools/, modules/, …).
    """
    root = root or PROJECT_ROOT
    files_removed = 0
    dirs_removed = 0

    for pycache_dir in root.rglob("__pycache__"):
        if not include_venv and "venv" in pycache_dir.parts:
            continue
        files_removed += sum(1 for f in pycache_dir.iterdir() if f.is_file())
        shutil.rmtree(pycache_dir, ignore_errors=True)
        dirs_removed += 1

    return {"pycache_files": files_removed, "pycache_dirs": dirs_removed}


def clear_tool_session_state(session_state) -> int:
    """Xóa kết quả / cấu hình cache trong Streamlit session_state."""
    prefixes = ("col_cmp", "row_cmp")
    exact_keys = ("col_cmp_v2_result", "row_cmp_v2_result")
    keys = [
        k for k in list(session_state.keys())
        if k.startswith(prefixes) or k in exact_keys
    ]
    for key in keys:
        del session_state[key]
    return len(keys)
