# -*- coding: utf-8 -*-
"""老台账升级（表头迁移）回归测试
=================================
真实台账是「旧表头」——只有 18 列，没有后加的「凭证类型 / 行程/明细」。
用户点一次「凭证入库」之后：

  1. 不能因为指纹变了（老行的凭证类型是空的）把行程单当新记录重复入库；
  2. 老行要能按「文件路径」对上号，并把新字段补起来；
  3. 不能覆盖用户已经改过的内容（费用类别、报销人、状态…）；
  4. 落盘后表头要变成 20 列的新格式。

跑法：
  C:\\Users\\JM\\.workbuddy\\binaries\\python\\envs\\gui\\Scripts\\python.exe 测试\\diag_migrate.py
"""
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

SRC = Path(r"E:\桌\湖北报销发票整理")
tmp = Path(tempfile.mkdtemp(prefix="invmig_"))

import app_paths  # noqa: E402
app_paths.LEDGER_FILE = tmp / "发票台账.xlsx"
app_paths.LEDGER_BAK = tmp / "台账备份"
app_paths.WORK_DIR = tmp / "缓存"
app_paths.OUTPUT_DIR = tmp / "输出"
for _d in (app_paths.LEDGER_BAK, app_paths.WORK_DIR, app_paths.OUTPUT_DIR):
    _d.mkdir(parents=True, exist_ok=True)

import openpyxl  # noqa: E402
import db  # noqa: E402
import invparse  # noqa: E402
import ledger  # noqa: E402

fail = []


def check(name, ok, extra=""):
    print(f"  {'✅' if ok else '❌'} {name}" + (f"　{extra}" if extra else ""))
    if not ok:
        fail.append(name)


# 旧表头（注意：没有 凭证类型 / 行程/明细）
OLD_COLUMNS = ["序号", "状态", "发票号码", "开票日期", "票种", "购方名称", "销方名称",
               "不含税金额", "税额", "价税合计", "项目/事由", "项目编号", "费用类别",
               "报销人", "报销批次", "入库时间", "文件路径", "提示"]


def write_old_ledger(rows):
    """造一份「旧表头」的 Excel —— 并且把它当成「外部已有的台账」导进库里。

    ⚠️ 台账真身已经是 SQLite 了，光写 Excel 文件不算数（Excel 只是导出快照）。
    现实里这个动作就是用户在设置页点「从 Excel 导入」——那条路会先备份原文件。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = ledger.SHEET
    ws.append(OLD_COLUMNS)
    for r in rows:
        ws.append([r.get(c, "") for c in OLD_COLUMNS])
    wb.save(str(app_paths.LEDGER_FILE))
    db.migrate_from_excel(app_paths.LEDGER_FILE, force=True)


try:
    files = invparse.collect_files(SRC)
    recs = invparse.dedupe_same_invoice([r for r in (invparse.parse_file(f) for f in files) if r])
    print(f"数据源 {SRC}：{len(files)} 个文件 → {len(recs)} 条凭证\n")

    # 挑两条「没有发票号码」的凭证（行程单/图片）来当老记录 —— 它们只有靠路径才能对上号
    no_no = [r for r in recs if not (r.get("no") or "").strip()]
    check("样本里有 ≥2 条无发票号的凭证", len(no_no) >= 2, f"{len(no_no)} 条")
    sample = no_no[:2]

    old_rows = []
    for i, rec in enumerate(sample, start=1):
        old_rows.append({
            "序号": i, "状态": "未报销", "发票号码": "",
            "开票日期": rec.get("date") or rec.get("trip_date") or "",
            "价税合计": rec.get("total"), "费用类别": "交通费", "报销人": "老张（手工填的）",
            "入库时间": "2026-01-01 09:00", "文件路径": rec.get("src", ""), "提示": "",
        })
    write_old_ledger(old_rows)
    print(f"  造了 {len(old_rows)} 行旧表头的台账：")
    for r in old_rows:
        print(f"    · {Path(str(r['文件路径'])).name[:46]}  ¥{r['价税合计']}")

    # ---------- 升级 ----------
    print("\n=== 用同一批凭证再入库一次（模拟点「凭证入库」）===")
    res = ledger.add_records(recs, batch="")

    hit = {str(x["rec"].get("src")) for x in res["in_ledger"] + res["dup"]}
    check("老记录被认出来（不重复登记）", all(str(r["文件路径"]) in hit for r in old_rows),
          f"命中 {len(hit)} / 样本 {len(old_rows)}")
    added_paths = {str(r.get("文件路径")) for r in res["added"]}
    check("老记录没有被当成新行重复插入",
          not any(str(r["文件路径"]) in added_paths for r in old_rows),
          "重复插入：" + str(sorted(added_paths & {str(r['文件路径']) for r in old_rows})))
    check("补全计数 > 0", res.get("patched", 0) > 0, f"patched={res.get('patched')}")

    # ---------- 落盘后的样子 ----------
    print("\n=== 落盘后的台账 ===")
    rows = ledger.load()
    check("总行数 = 新增 + 老的", len(rows) == len(res["added"]) + len(old_rows),
          f"{len(rows)} 行")
    old_now = [r for r in rows if str(r.get("入库时间")) == "2026-01-01 09:00"]
    check("老记录还在（按入库时间找回）", len(old_now) == len(old_rows), f"{len(old_now)} 行")
    check("老记录的「凭证类型」补上了",
          all(str(r.get("凭证类型") or "").strip() for r in old_now),
          str([str(r.get("凭证类型")) for r in old_now]))
    check("老记录的「行程/明细」补上了",
          all(str(r.get("行程/明细") or "").strip() for r in old_now),
          str([str(r.get("行程/明细"))[:28] for r in old_now]))
    check("老记录手工填的报销人没被覆盖",
          all(str(r.get("报销人")) == "老张（手工填的）" for r in old_now),
          str([str(r.get("报销人")) for r in old_now]))
    check("老记录的费用类别没被覆盖",
          all(str(r.get("费用类别")) == "交通费" for r in old_now))

    wb = openpyxl.load_workbook(str(app_paths.LEDGER_FILE))
    hdr = [c.value for c in wb[ledger.SHEET][1]]
    check("表头已升级成 20 列含新列",
          len(hdr) == len(ledger.COLUMNS) and "凭证类型" in hdr and "行程/明细" in hdr,
          f"{len(hdr)} 列：{hdr}")

    # ---------- 老台账里「原本是数电票」的行也要能补 ----------
    print("\n=== 发票类老行（本来有号码）===")
    with_no = next(r for r in recs if (r.get("no") or "").strip())
    write_old_ledger([{
        "序号": 1, "状态": "已报销", "发票号码": with_no["no"], "开票日期": with_no.get("date", ""),
        "价税合计": with_no.get("total"), "费用类别": "材料费", "报销人": "",
        "入库时间": "2026-01-01 09:00", "文件路径": with_no.get("src", ""), "提示": "",
    }])
    res2 = ledger.add_records(recs, batch="")
    r2 = ledger.load()
    tgt = [x for x in r2 if str(x.get("发票号码")) == str(with_no["no"])][0]
    check("有号码的老行被标为重复风险（已报销）", any(
        str(d["rec"].get("no")) == str(with_no["no"]) for d in res2["dup"]), f"{len(res2['dup'])} 条")
    check("有号码的老行补上了凭证类型/票种", bool(str(tgt.get("凭证类型")).strip())
          and bool(str(tgt.get("票种")).strip()),
          f"{tgt.get('凭证类型')} / {tgt.get('票种')}")
    check("老行的状态没被改回未报销", tgt.get("状态") == "已报销", str(tgt.get("状态")))

    print()
    print(f"❌ 未通过 {len(fail)} 项：" + "、".join(fail) if fail else "✅ 全部通过")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

sys.exit(1 if fail else 0)
