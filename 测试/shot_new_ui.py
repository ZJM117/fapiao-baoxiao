# -*- coding: utf-8 -*-
"""
生成「新版界面」预览图 + 报销单样例（PDF / Excel）
================================================================================
全程无窗口：进程内起本地服务 → 无头 Edge 截图；报销单一并出真实 PDF 和 Excel。
所有台账 / 生成记录 / 缓存都落在 %TEMP% 里，不碰用户的真实数据。
产出：
  输出/界面预览/6-台账查询与删除.png      台账页（新查询条 + 删除选中 / 清空台账）
  输出/界面预览/7-报销单新版式.png        新报销单版式（与 PDF 完全同源）
  输出/新版报销单样例/*.pdf / *.xlsx      可以直接打开看的两份样例
"""
import io
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE = Path(__file__).resolve().parent.parent
REAL_OUT = BASE / "输出"
PREVIEW = REAL_OUT / "界面预览"
SAMPLES = REAL_OUT / "新版报销单样例"
PREVIEW.mkdir(parents=True, exist_ok=True)
SAMPLES.mkdir(parents=True, exist_ok=True)

TMP = Path(tempfile.mkdtemp(prefix="inv_shot_"))
sys.path.insert(0, str(BASE))

import app_paths                                                     # noqa: E402
app_paths.LEDGER_FILE = TMP / "发票台账.xlsx"
app_paths.LEDGER_BAK = TMP / "台账备份"
app_paths.OUTPUT_DIR = TMP / "输出"
app_paths.WORK_DIR = TMP / "缓存"
app_paths.CONFIG_FILE = TMP / "config.json"
app_paths.LAUNCH_LOG = TMP / "launch.log"
app_paths.UI_PORT_FILE = TMP / "ui_port.txt"
for _d in (app_paths.LEDGER_BAK, app_paths.OUTPUT_DIR, app_paths.WORK_DIR):
    _d.mkdir(parents=True, exist_ok=True)

import app_web                                                       # noqa: E402
import ledger                                                        # noqa: E402
import report_pdf as rp                                              # noqa: E402

EDGE = app_paths.EDGE


def row(seq, status, ctype, no, date, kind, trip, seller, amt, tax, total,
        proj, cat, person, batch, tip=""):
    return {"序号": seq, "状态": status, "凭证类型": ctype, "发票号码": no,
            "开票日期": date, "票种": kind, "行程/明细": trip, "购方名称": "国网湖北省电力有限公司",
            "销方名称": seller, "不含税金额": amt, "税额": tax, "价税合计": total,
            "项目/事由": proj, "项目编号": "1815P825000P", "费用类别": cat,
            "报销人": person, "报销批次": batch, "入库时间": "2026-09-15 13:20",
            "文件路径": rf"E:\桌\湖北报销发票整理\{no or (ctype or 'x')}.pdf", "提示": tip}


ROWS = [
    row(1, "未报销", "火车票", "", "2025-03-02", "电子发票（铁路电子客票）",
        "G1234 武汉 → 宜昌东 2025-03-02 08:15 二等座 07车12F", "中国铁路武汉局集团有限公司",
        71.56, 3.44, 75.00, "恩施建始龙坪网格10千伏线路 1815P825000P", "交通费", "张三", "9月现场审计"),
    row(2, "未报销", "火车票", "", "2025-03-04", "电子发票（铁路电子客票）",
        "D2233 宜昌东 → 武汉 2025-03-04 17:40 二等座 03车05A", "中国铁路武汉局集团有限公司",
        61.47, 2.53, 64.00, "恩施建始龙坪网格10千伏线路 1815P825000P", "交通费", "张三", "9月现场审计"),
    row(3, "未报销", "发票", "26349119423005870096", "2025-03-05", "电子发票（普通发票）",
        "", "武汉某某酒店管理有限公司", 1037.74, 62.26, 1100.00,
        "恩施建始龙坪网格10千伏线路 1815P825000P", "住宿费", "张三", "9月现场审计"),
    row(4, "已报销", "发票", "26349119423005870097", "2025-02-10", "电子发票（专用发票）",
        "", "武汉某办公用品有限公司", 353.98, 46.02, 400.00,
        "黄石阳新富池供电所台区改造 1815P825001P", "办公费", "李四", "8月报销单",
        tip="⚠ 重复：台账中已存在且状态为【已报销】"),
    row(5, "未报销", "打车行程单", "", "", "",
        "悦行优选 合肥南站 → 财智中心 13.58元", "南京领行科技股份有限公司",
        13.58, 0, 13.58, "", "交通费", "张三", "9月现场审计"),
    row(6, "未报销", "飞机行程单", "", "2025-03-01", "航空运输电子客票行程单",
        "CZ3846 武汉天河 → 昆明长水 2025-03-01 09:20 经济舱", "中国南方航空股份有限公司",
        1372.48, 123.52, 1496.00, "恩施建始龙坪网格10千伏线路 1815P825000P", "交通费", "张三", "9月现场审计"),
]
ledger.save(ROWS, do_backup=False)

api = app_web.Api()
seqs_trip = [1, 2, 3, 5, 6]
meta_trip = {"kind": "差旅费报销单", "person": "张三", "dept": "审计部",
             "reason": "恩施建始龙坪网格10千伏线路工程现场审计",
             "note": "附件：电子发票及行程单共 5 份", "date": "2026年09月15日",
             "start": "2025-03-01", "end": "2025-03-05", "place": "湖北宜昌 / 云南昆明"}
rows_trip = [r for r in ROWS if r["序号"] in seqs_trip]

# 报销单样例（真实 PDF + Excel）
res = api.make_report(meta_trip, seqs_trip, "both")
for p in res.get("outs", []):
    shutil.copy2(p, SAMPLES / Path(p).name)
    print("样例：", Path(p).name)
res2 = api.make_report({"kind": "费用报销单", "person": "张三", "dept": "审计部",
                        "reason": "项目现场审计费用", "date": "2026年09月15日"},
                       [1, 2, 3, 5], "both")
for p in res2.get("outs", []):
    shutil.copy2(p, SAMPLES / Path(p).name)
    print("样例：", Path(p).name)

# ---------- 1. 报销单版式（HTML 与 PDF 同源，直接截图）----------
report_html = rp.build_report_html(rows_trip, meta_trip)
tmp_html = TMP / "report_preview.html"
tmp_html.write_text(report_html, encoding="utf-8")
out7 = PREVIEW / "7-报销单新版式.png"
subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--hide-scrollbars",
                f"--user-data-dir={TMP / 'edge1'}", "--window-size=880,1240",
                f"--screenshot={out7}", tmp_html.as_uri()],
               capture_output=True, timeout=120,
               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
print("截图：", out7.exists(), out7)

# ---------- 2. 界面（进程内起服务 + 无头 Edge 截图）----------
srv = app_web.start_server(api, app_paths.GUI_DIR, 0)
port = srv.server_address[1]
shots = [("6-台账查询与删除.png", "page-ledger", 1500, 1180),
         ("8-生成文件可删除.png", "page-report", 1500, 1500)]
for name, page, w, h in shots:
    out = PREVIEW / name
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--hide-scrollbars",
                    f"--user-data-dir={TMP / ('edge_' + page)}", f"--window-size={w},{h}",
                    "--virtual-time-budget=12000",
                    f"--screenshot={out}", f"http://127.0.0.1:{port}/index.html#{page}"],
                   capture_output=True, timeout=180,
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    print("截图：", out.exists(), out, out.stat().st_size if out.exists() else "")
try:
    srv.shutdown()
except Exception:
    pass
print("临时目录：", TMP)
