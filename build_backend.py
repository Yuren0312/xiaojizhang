# -*- coding: utf-8 -*-
"""小虾记账 - 后端引擎打包脚本（PyInstaller onedir）

产出:
    backend-server/server.exe + backend-server/_internal/...
    该目录会被 electron-builder 作为 extraResources 一并打进安装包。

用法:
    python build_backend.py            # 打包
    python build_backend.py --clean    # 先清理再打包
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BACKEND_DIR = ROOT / "backend"
ENTRY = BACKEND_DIR / "server.py"

# 每次构建使用独立目录：避免 PyInstaller 清理旧产物时被文件锁/安全软件卡住
# 产物放在 build-py/，与 electron-builder 的 buildResources（assets）隔离
STAMP = time.strftime("%Y%m%d_%H%M%S")
WORK = ROOT / "build-py" / f"pyi_{STAMP}"
DIST_TMP = ROOT / "build-py" / f"dist_{STAMP}"
OUT_ROOT = ROOT / "backend-server"
FINAL_DIR = OUT_ROOT / "server"

ARGS = [
    "--noconfirm",
    "--onedir",                          # 目录模式：冷启动比 onefile 快 2-3 倍
    "--name", "server",
    "--distpath", str(DIST_TMP),
    "--workpath", str(WORK),
    "--specpath", str(WORK),
    "--console",                         # 保留控制台，便于排障（Electron 隐藏窗口）
    # ---- 业务代码与资源 ----
    "--paths", str(ROOT),
    "--collect-submodules", "accounting_report",
    "--add-data", f"{ROOT / 'accounting_report' / 'rules'};accounting_report/rules",
    "--add-data", f"{ROOT / 'accounting_report' / 'templates'};accounting_report/templates",
    # ---- 显式隐藏导入（避免被裁剪） ----
    "--hidden-import", "uvicorn.logging",
    "--hidden-import", "uvicorn.loops.auto",
    "--hidden-import", "uvicorn.protocols.http.auto",
    "--hidden-import", "uvicorn.protocols.websockets.auto",
    "--hidden-import", "uvicorn.lifespan.on",
    "--hidden-import", "multipart",
    "--hidden-import", "email.mime.multipart",
    # PDF 解析为函数内延迟导入，需显式声明
    "--hidden-import", "fitz",
    "--collect-all", "fitz",
    "--collect-all", "pymupdf",
    "--hidden-import", "xlrd",
    # ---- 排除体积大户 ----
    "--exclude-module", "matplotlib",
    "--exclude-module", "pandas",
    "--exclude-module", "scipy",
    "--exclude-module", "tkinter",
    "--exclude-module", "PyQt5",
    "--exclude-module", "IPython",
]


def clean() -> None:
    """尽量清理旧产物（文件被占用时忽略）"""
    for p in (OUT_ROOT, ROOT / "build-py"):
        if p.exists():
            shutil.rmtree(p, ignore_errors=True)
            print(f"[clean] {p}")


def build() -> None:
    if not ENTRY.exists():
        sys.exit(f"入口文件不存在: {ENTRY}")

    cmd = [sys.executable, "-m", "PyInstaller", *ARGS, str(ENTRY)]
    print("[build] 正在打包后端引擎（约需 1-2 分钟）...")
    result = subprocess.run(cmd, cwd=str(ROOT))
    if result.returncode != 0:
        sys.exit("[error] 后端打包失败")

    built = DIST_TMP / "server"
    if not built.is_dir():
        sys.exit(f"[error] 未生成 {built}")

    # 落位到 electron-builder 约定的 backend-server/server
    OUT_ROOT.mkdir(parents=True, exist_ok=True)
    if FINAL_DIR.exists():
        shutil.rmtree(FINAL_DIR, ignore_errors=True)
    shutil.copytree(built, FINAL_DIR, dirs_exist_ok=True)

    exe = FINAL_DIR / "server.exe"
    if not exe.exists():
        sys.exit(f"[error] 未生成 {exe}")

    size_mb = sum(f.stat().st_size for f in FINAL_DIR.rglob("*") if f.is_file()) / 1024 / 1024
    print(f"[ok] {exe}  合计 {size_mb:.1f} MB")


if __name__ == "__main__":
    if "--clean" in sys.argv:
        clean()
    build()
