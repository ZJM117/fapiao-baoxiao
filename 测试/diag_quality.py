# -*- coding: utf-8 -*-
"""质量清单：把「没认出来 / 缺关键字段」的逐条列出来，带完整路径。只读。"""
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import invparse  # noqa: E402

SRC = Path(r"E:\桌\湖北报销发票整理")
files = invparse.collect_files(SRC)
rows = []
for f in files:
    try:
        r = invparse.parse_file(f)
    except Exception as e:
        r = {"src": str(f), "warn": [f"异常 {type(e).__name__}: {e}"], "kind": "", "ctype": ""}
    if r:
        r["_f"] = f
        rows.append(r)

print("=" * 78)
print("【A】kind 为空的")
for r in rows:
    if not (r.get("kind") or "").strip():
        print(f"  {r['_f'].relative_to(SRC)}")
        print(f"      ctype={r.get('ctype')} src_kind={r.get('src_kind')} "
              f"no={r.get('no')!r} total={r.get('total')} warn={r.get('warn')}")

print()
print("=" * 78)
print("【B】该有金额却没有的（行程单/登机牌除外，它们本来可能没金额）")
for r in rows:
    if r.get("total") is None and r.get("ctype") not in (invparse.T_TRIP, invparse.T_BOARD):
        print(f"  {r['_f'].relative_to(SRC)}")
        print(f"      ctype={r.get('ctype')} kind={r.get('kind')!r} no={r.get('no')!r} warn={r.get('warn')}")

print()
print("=" * 78)
print("【C】按凭证类型的明细抽样（看 route/detail/traveler 抽对没有）")
from collections import defaultdict
g = defaultdict(list)
for r in rows:
    g[r.get("ctype") or "(空)"].append(r)
for t, rs in sorted(g.items()):
    print(f"\n--- {t}（{len(rs)} 份）---")
    for r in rs[:6]:
        print(f"  文件：{r['_f'].name[:52]}")
        print(f"    no={r.get('no') or '—'}  date={r.get('date') or '—'}  total={r.get('total')}")
        print(f"    route={r.get('route')!r}  traveler={r.get('traveler')!r}  train={r.get('train_no')!r}")
        print(f"    detail={r.get('detail')!r}")

print()
print("=" * 78)
print("【D】合肥.zip 里面是什么")
z = SRC / r"某某公司\待报销的发票\网上购票系统-电子发票通知\合肥.zip"
if z.exists():
    with zipfile.ZipFile(z) as zf:
        for n in zf.namelist()[:20]:
            print(f"   {n}  ({zf.getinfo(n).file_size} B)")
