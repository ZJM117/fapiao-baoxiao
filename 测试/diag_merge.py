# -*- coding: utf-8 -*-
"""
隔离自检：合并发票时「按票面内容认票」
======================================
复现用户在真实数据上遇到的 bug：
    同一张票既有 OFD 又有 PDF，但用户手存的 PDF 文件名是人话
    （「张三-汉口-郑州东.pdf」），里面没有 20 位发票号 →
    旧的 _resolve_pdf 只按文件名找 → 找不到 → 选 4 张只合进去 1 张。

这里造一棵和用户一样的目录树（4 张票：1 张本来就是 PDF，3 张只有 OFD，
但其中 2 张的 PDF 以人话文件名躺在上一层目录），验证现在能按票面认出来。

全部在 %TEMP% 副本里跑，绝不碰真实台账 / 输出 / 配置。无窗口。
跑法：  python 测试\\diag_merge.py
"""
import io
import sys
import tempfile
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))

TMP = Path(tempfile.mkdtemp(prefix="inv_merge_"))

# ⚠️ 必须在 import ledger / app_web / report_pdf 之前改路径
import app_paths                                                    # noqa: E402
app_paths.LEDGER_FILE = TMP / "发票台账.xlsx"
app_paths.LEDGER_BAK = TMP / "台账备份"
app_paths.OUTPUT_DIR = TMP / "输出"
app_paths.WORK_DIR = TMP / "缓存"
app_paths.CONFIG_FILE = TMP / "config.json"
app_paths.LAUNCH_LOG = TMP / "launch.log"
app_paths.UI_PORT_FILE = TMP / "ui_port.txt"
for _d in (app_paths.LEDGER_BAK, app_paths.OUTPUT_DIR, app_paths.WORK_DIR):
    _d.mkdir(parents=True, exist_ok=True)

import app_web                                                      # noqa: E402
import ledger                                                       # noqa: E402
import report_pdf as rp                                             # noqa: E402

FAIL = []


def chk(cond, label, extra=""):
    print(("  [OK] " if cond else "  [XX] ") + label + ("" if cond else f"   → {extra}"))
    if not cond:
        FAIL.append(label)


def section(t):
    print("\n" + "=" * 66 + f"\n{t}\n" + "=" * 66)


# ---------------------------------------------------------------- 造假数据
# 和用户现场一模一样：4 张票
#   1) 直接就是 PDF（张三-地铁.pdf，名字不含号码）
#   2) OFD 放在以号码命名的子文件夹里，PDF 放在上一层但叫「张三-汉口-郑州东.pdf」
#   3) 同上，「张三-济南-武汉.pdf」
#   4) 只有 OFD，没有 PDF（用来验「少一张」的提示）
NO_METRO = "26427000000413532641"
NO_TRAIN1 = "26429121050002258892"
NO_TRAIN2 = "26379116295001315516"
NO_TRAIN3 = "26419165773005057163"

SRC = TMP / "报销凭证"
SRC.mkdir(parents=True, exist_ok=True)


def make_pdf(path: Path, no: str, title: str):
    """用无头 Edge 打一张「假发票」PDF：票面里有 20 位发票号，文字层可读"""
    html = f"""<!doctype html><html><head><meta charset="utf-8"><style>
    @page {{ size: A4; margin: 14mm; }}
    body {{ font: 14px/1.9 "Microsoft YaHei",sans-serif; }}
    h1 {{ font-size: 17px; }}
    </style></head><body>
    <h1>电子发票（铁路电子客票）</h1>
    <div>发票号码：{no}</div>
    <div>{title}</div>
    <div>价税合计：￥263.00</div>
    </body></html>"""
    rp.html_to_pdf(html, path)
    return path


def make_ofd(path: Path):
    """OFD 只是个占位文件：本测试只验证「找不到 PDF 时怎么提示」，不去解析它"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"OFD-DUMMY")
    return path


# 1) 地铁票：直接是 PDF
make_pdf(SRC / "张三-地铁.pdf", NO_METRO, "武汉市轨道交通 6.00 元")
# 2)、3) 火车票：OFD 在子文件夹里，PDF 在上一层、名字不含号码
for no, name, title in ((NO_TRAIN1, "张三-汉口-郑州东.pdf", "汉口 → 郑州东 263.00 元"),
                        (NO_TRAIN2, "张三-济南-武汉.pdf", "济南 → 武汉 485.00 元")):
    make_ofd(SRC / no / f"{no}.ofd")
    make_pdf(SRC / name, no, title)
# 4) 只有 OFD，没有 PDF
make_ofd(SRC / NO_TRAIN3 / f"{NO_TRAIN3}.ofd")


def row(seq, no, ctype, date, total, path):
    return {"序号": seq, "状态": "未报销", "凭证类型": ctype, "发票号码": no,
            "开票日期": date, "票种": "电子发票", "行程/明细": "", "购方名称": "某某公司",
            "销方名称": "中国铁路", "不含税金额": "", "税额": "", "价税合计": total,
            "项目/事由": "现场审计", "项目编号": "1815P825000P", "费用类别": "交通费",
            "报销人": "张三", "报销批次": "", "入库时间": "2026-09-15 14:20",
            "文件路径": str(path), "提示": ""}


ROWS = [
    row(1, NO_METRO, "发票", "2026-05-22", 6.00, SRC / "张三-地铁.pdf"),
    row(2, NO_TRAIN1, "火车票", "2026-05-22", 263.00, SRC / NO_TRAIN1 / f"{NO_TRAIN1}.ofd"),
    row(3, NO_TRAIN2, "火车票", "2026-05-22", 485.00, SRC / NO_TRAIN2 / f"{NO_TRAIN2}.ofd"),
    row(4, NO_TRAIN3, "火车票", "2026-05-22", 249.00, SRC / NO_TRAIN3 / f"{NO_TRAIN3}.ofd"),
]
ledger.save(ROWS)

# ---------------------------------------------------------------- 开跑
section("一、_find_pdf：每张票该找到什么")
api = app_web.Api()
for r in ROWS:
    p, why = api._find_pdf(r)
    hit = "找到" if p else "找不到"
    print(f"  {hit} 序号{r['序号']} {Path(str(r['文件路径'])).name}")
    print(f"        → {p if p else why}")
chk(api._find_pdf(ROWS[0])[0] is not None, "① 本来就是 PDF 的直接用")
chk(str(api._find_pdf(ROWS[1])[0]).endswith("张三-汉口-郑州东.pdf"),
    "② 按票面内容认出「张三-汉口-郑州东.pdf」（文件名不含发票号）",
    api._find_pdf(ROWS[1])[0])
chk(str(api._find_pdf(ROWS[2])[0]).endswith("张三-济南-武汉.pdf"),
    "③ 同上，认出「张三-济南-武汉.pdf」", api._find_pdf(ROWS[2])[0])
chk(api._find_pdf(ROWS[3])[0] is None, "④ 真的只有 OFD 的那张，返回 None")
chk("OFD" in api._find_pdf(ROWS[3])[1], "④ 并且说明了原因（提到 OFD）",
    api._find_pdf(ROWS[3])[1])

section("二、合并 4 张 → 3 张进得去、1 张明确提示")
res = api.merge_invoices([1, 2, 3, 4], "2up")
print("  ", {k: res.get(k) for k in ("files", "pages", "missing", "asked", "error")})
chk(not res.get("error"), "没有报错", res.get("error"))
chk(res.get("asked") == 4, f"选中的是 4 张（实际 {res.get('asked')}）")
chk(res.get("files") == 3, f"合并进去 3 张（实际 {res.get('files')}）")
chk(res.get("pages") == 2, f"2up 版式 → 2 页 A4（实际 {res.get('pages')}）")
chk(res.get("missing") == 1, f"提示少 1 张（实际 {res.get('missing')}）")
ml = res.get("missing_list") or []
chk(len(ml) == 1 and NO_TRAIN3 in str(ml[0].get("name", "")) + str(ml[0].get("no", "")),
    "少的那张点名说是哪一条", ml)
chk("OFD" in str(ml[0].get("why") if ml else ""), "并说明了原因", ml)

section("三、真的落了 3 张的票面进 PDF")
from pypdf import PdfReader                                          # noqa: E402
reader = PdfReader(res["out"])
print(f"   合并件：{Path(res['out']).name}，{len(reader.pages)} 页")
text = "\n".join((p.extract_text() or "") for p in reader.pages)
for no, tag in ((NO_METRO, "地铁票"), (NO_TRAIN1, "汉口-郑州东"), (NO_TRAIN2, "济南-武汉")):
    chk(no in text, f"合并件里含 {tag}（{no}）")
chk(NO_TRAIN3 not in text, "只有 OFD 的那张确实没混进来")
chk(len(reader.pages) == 2, f"PDF 真的是 2 页（实际 {len(reader.pages)}）")

section("四、全都没 PDF 时，报错并逐条说明原因")
res2 = api.merge_invoices([4], "2up")
chk(res2.get("error"), "返回 error", res2)
chk(res2.get("missing") == 1, f"missing=1（实际 {res2.get('missing')}）")
chk(res2.get("missing_list") and "OFD" in res2["missing_list"][0]["why"],
    "error 里也带了原因列表", res2.get("missing_list"))

section("五、新增同票 PDF 后（不重启程序）能自动认出来")
# 模拟用户把第 4 张的 PDF 也拷进上一层
make_pdf(SRC / "张三-郑州东-济南东.pdf", NO_TRAIN3, "郑州东 → 济南东 249.00 元")
api2 = app_web.Api()
res3 = api2.merge_invoices([1, 2, 3, 4], "2up")
chk(res3.get("files") == 4, f"这次 4 张全进（实际 {res3.get('files')}）", res3.get("missing_list"))
chk(res3.get("missing") == 0, f"不再有 missing（实际 {res3.get('missing')}）")
txt3 = "\n".join((p.extract_text() or "") for p in PdfReader(res3["out"]).pages)
chk(NO_TRAIN3 in txt3, "第 4 张的票面也进去了")
chk(len(PdfReader(res3["out"]).pages) == 2, "4 张 2up 仍是 2 页")

section("六、生成记录里带上了「少几张」")
hist = app_web.Api().list_outputs().get("items", [])
merge_items = [h for h in hist if h.get("kind") == "发票合并"]
print("   ", [(h["name"][-16:], h.get("count"), h.get("missing")) for h in merge_items])
chk(any(h.get("missing") == 1 for h in merge_items),
    "第一条合并记录带着 missing=1，界面上会显示「少 1 张」",
    [(h.get("name"), h.get("missing")) for h in merge_items])

section("七、同名不同后缀 / 号码在文件名里的老路子没被破坏")
chk(str(api._find_pdf(ROWS[0])[0]).endswith(".pdf"), "PDF 行还是直接命中自己")

print("\n" + "=" * 66)
if FAIL:
    print(f"❌ {len(FAIL)} 项没通过：")
    for f in FAIL:
        print("   -", f)
else:
    print("✅ 全部通过")
print("=" * 66)
sys.exit(1 if FAIL else 0)
