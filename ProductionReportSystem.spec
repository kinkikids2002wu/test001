# -*- mode: python ; coding: utf-8 -*-
# ProductionReportSystem.spec
# 目標：強制使用 _internal 資料夾結構
# 
# 期望結構：
#   dist\ProductionReportSystem\
#   ├── ProductionReportSystem.exe    ← 只有這個在外面
#   └── _internal\                     ← 所有其他檔案
#       ├── python38.dll
#       ├── *.pyd
#       ├── base_library.zip
#       ├── templates\
#       ├── static\
#       └── forms\

import os
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules

# ---------- 找到專案根目錄 ----------
def _guess_spec_dir():
    for a in reversed(sys.argv):
        try:
            p = Path(a)
        except Exception:
            continue
        if str(p).lower().endswith(".spec") and p.exists():
            return p.resolve().parent
    return Path.cwd().resolve()

HERE = _guess_spec_dir()

# ---------- 專案主要檔案 ----------
APP_PY = HERE / "app.py"

# ---------- 收集資料檔 ----------
datas = []

# 關鍵設定：使用 "_internal" 作為目標路徑
# 這會強制 PyInstaller 將檔案放到 _internal 資料夾
DEST_BASE = Path("_internal")

def add_file(src: Path, dest_dir: Path):
    """添加單個檔案"""
    datas.append((str(src), str(dest_dir)))

def add_tree(src_dir: Path, dest_dir: Path):
    """遞迴添加整個目錄"""
    if not src_dir.exists():
        return
    for f in src_dir.rglob("*"):
        if f.is_file():
            rel_parent = f.relative_to(src_dir).parent
            add_file(f, dest_dir / rel_parent)

# 1. Templates 資料夾 → _internal/templates/
tmpl_dir = HERE / "templates"
if tmpl_dir.exists():
    add_tree(tmpl_dir, DEST_BASE / "templates")
else:
    for name in ("index_table.html", "index_form.html", "index_hybrid.html", "print_template.html"):
        p = HERE / name
        if p.exists():
            add_file(p, DEST_BASE / "templates")

# 2. Static 資料夾 → _internal/static/
static_dir = HERE / "static"
if static_dir.exists():
    add_tree(static_dir, DEST_BASE / "static")
else:
    css = HERE / "style.css"
    if css.exists():
        add_file(css, DEST_BASE / "static" / "css")

# 3. Forms 資料夾 → _internal/forms/
forms_dir = HERE / "forms"
if forms_dir.exists():
    add_tree(forms_dir, DEST_BASE / "forms")
else:
    excel_tpl = HERE / "生產日報表修改申請1.xlsx"
    if excel_tpl.exists():
        add_file(excel_tpl, DEST_BASE / "forms")

# ---------- 隱藏導入 ----------
hiddenimports = []

for pkg in ("flask", "jinja2", "werkzeug", "itsdangerous", "click", "markupsafe"):
    try:
        hiddenimports += collect_submodules(pkg)
    except Exception:
        pass

for pkg in ("pandas", "openpyxl", "pyodbc"):
    try:
        hiddenimports += collect_submodules(pkg)
    except Exception:
        pass

# ---------- PyInstaller 分析階段 ----------
block_cipher = None

a = Analysis(
    [str(APP_PY)],
    pathex=[str(HERE)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'IPython', 'notebook', 'PIL'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ---------- 建立執行檔 ----------
# 關鍵：exclude_binaries=True 才會產生 onedir 模式
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,          # ← 必須是 True
    name="ProductionReportSystem",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=None,
)

# ---------- 收集檔案 ----------
# 關鍵：這裡決定了是否使用 _internal 資料夾
# PyInstaller 5.13+ 會自動將 binaries 和 zipfiles 放到 _internal
coll = COLLECT(
    exe,
    a.binaries,      # DLL 檔案 → 會進 _internal
    a.zipfiles,      # base_library.zip → 會進 _internal
    a.datas,         # 自訂資料檔（templates, static）
    strip=False,
    upx=True,
    upx_exclude=[],
    name="ProductionReportSystem",
)
