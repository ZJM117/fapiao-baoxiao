# -*- coding: utf-8 -*-
"""对 E:\桌\湖北报销发票整理 整目录跑一遍现有解析器，看识别情况。
只读，不改磁盘、不弹窗。
"""
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import invparse  # noqa: E402

SRC = Path(r"E:\桌\湖北报销发票整理")

files = invparse.collect_files(SRC, recursive=True)
print(f"待解析文件：{len(files)}\n")

rows = []
for f in files:
    try:
        r = invparse.parse_file(f)
    except Exception as e:
        r = {"no": "", "src": str(f), "kind": "", "date": "", "total": None,
             "warn": [f"异常 {type(e).__name__}: {e}"], "src_kind": "err"}
    if r:
        rows.append(r)

ok_no = [r for r in rows if r.get("no")]
ok_no_amt = [r for r in rows if r.get("no") and r.get("total") is not None]
no_no = [r for r in rows if not r.get("no")]

print(f"解析成功（有发票号码）：{len(ok_no)}")
print(f"  · 号码+金额都有：{len(ok_no_amt)}")
print(f"  · 有号码但没金额：{len(ok_no) - len(ok_no_amt)}")
print(f"没有发票号码：{len(no_no)}\n")

print("=== 有号码但缺金额 / 缺日期 的 ===")
for r in ok_no:
    if r.get("total") is None or not r.get("date"):
        print(f"  {Path(r['src']).name[:56]:58s} kind={r.get('kind','')[:18]:20s} "
              f"date={r.get('date') or '-':11s} total={r.get('total')}")

print()
print("=== 没有发票号码的（按文件名）===")
for r in no_no:
    print(f"  {Path(r['src']).name[:60]:62s} kind={r.get('kind','')!r} total={r.get('total')}")

print()
print("=== 票种分布 ===")
for k, v in Counter((r.get("kind") or "(空)") for r in ok_no + no_no).most_common(30):
    print(f"  {v:4d}  {k}")

print()
print("=== src_kind 分布 ===")
for k, v in Counter(r.get("src_kind") for r in rows).most_common():
    print(f"  {v:4d}  {k}")

print()
print("=== warn 汇总 ===")
w = Counter()
for r in rows:
    for x in (r.get("warn") or []):
        w[x] += 1
for k, v in w.most_common(20):
    print(f"  {v:4d}  {k}")
