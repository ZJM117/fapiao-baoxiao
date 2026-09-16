# -*- coding: utf-8 -*-
"""再看两份样本的完整文本：航空客票行程单 + 滴滴行程单 + 已有的正规报销单版式。只读。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

FILES = [
    r"E:\桌\审计\个人\报销凭证\某某项目项目\兰州—济南（行程单）.pdf",
    r"E:\桌\审计\个人\报销凭证\某某投标\报销打印.pdf",
    r"E:\桌\审计\个人\报销凭证\某某投标\差旅费报销单.pdf",
    r"E:\桌\审计\个人\报销凭证\某某项目项目\某某项目差旅费（新）\打印版本.pdf",
]

for s in FILES:
    p = Path(s)
    print("=" * 78)
    print(f"【{p.name}】{p.stat().st_size // 1024} KB" if p.exists() else f"【{p.name}】不存在")
    if not p.exists():
        continue
    from pypdf import PdfReader
    r = PdfReader(str(p))
    print(f"  页数 {len(r.pages)}")
    for i, pg in enumerate(r.pages[:2]):
        t = pg.extract_text() or ""
        print(f"  --- p{i+1} 共 {len(t)} 字 ---")
        for ln in [x.rstrip() for x in t.splitlines()]:
            if ln.strip():
                print("   | " + ln)
