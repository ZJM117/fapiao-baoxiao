# -*- coding: utf-8 -*-
"""确保 PDF 后端缺失时不会误报成扫描件；不打开任何真实 PDF。"""
import builtins
import sys
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import invparse

FAIL = []


def chk(cond, label):
    print(("  [OK] " if cond else "  [XX] ") + label)
    if not cond:
        FAIL.append(label)


original_import = builtins.__import__


def without_pdf_backends(name, *args, **kwargs):
    if name == "pypdf" or name == "pdfminer.high_level":
        raise ModuleNotFoundError(name)
    return original_import(name, *args, **kwargs)


with patch("builtins.__import__", side_effect=without_pdf_backends):
    rec = invparse.parse_pdf(Path("diagnostic-only.pdf"))

warning = " ".join(rec.get("warn", []))
chk(rec.get("kind") == "PDF 解析依赖缺失",
    "依赖缺失显示为解析依赖错误，而非图片型 PDF")
chk("requirements.txt" in warning,
    "错误提示给出本机依赖安装命令")
chk("没有文字层" not in warning,
    "缺依赖时不误称原文件是扫描件")

print("\n" + ("全部通过" if not FAIL else f"{len(FAIL)} 项失败"))
sys.exit(1 if FAIL else 0)
