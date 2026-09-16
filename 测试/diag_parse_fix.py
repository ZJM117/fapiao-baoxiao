# -*- coding: utf-8 -*-
"""
隔离自检：识别修复 + 票面红字标记 + 台账「票面标记」列 + 重新识别
================================================================================
覆盖用户报的几个问题：
  ① 「事由 / 摘要」从哪来 —— 来自发票备注栏，票上备注空着时不能把明细行当备注；
  ② 机票销方识别出一堆乱七八糟 —— 「填开单位:」标签与值分行的版式；
  ③ 交通费 13.58 的销售方变成「购方开户银行:-;银行账号:-;」—— 购/销方并排同行的版式；
  ④ 火车票的退票 / 改签要标注 —— 票面是**红字**，按字符颜色拎出来。

台账写入一律在 %TEMP% 副本里做；真实发票目录只读。
跑法：  python 测试\\diag_parse_fix.py
"""
import io
import sys
import tempfile
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))

TMP = Path(tempfile.mkdtemp(prefix="inv_pfix_"))

# ⚠️ 必须在 import ledger / app_web 之前改路径（它们用 from app_paths import XXX 取值）
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
import invparse as ip                                               # noqa: E402
import ledger as lg                                                 # noqa: E402
import report_pdf as rp                                             # noqa: E402

FAIL = []
SKIP = []

# 真实票据目录（只读；文件不在了就跳过对应用例，不误报）
SRC = Path(r"E:\桌\湖北报销发票整理\报销")
F_GAODE1358 = SRC / "王小明（某某公司）" / "【合肥南-财智中心-13.58元-1个行程】高德打车电子发票.pdf"
F_FLIGHT = SRC / "李强（某某公司）" / "武汉-烟台  李强电子行程单748元.pdf"
F_REFUND = SRC / "李强（某某公司）" / "济南东-烟台 李强 10.pdf"
F_CHANGE = SRC / "李强（某某公司）" / "烟台-济南  李强 198.pdf"
F_DIFF = SRC / "王小明（某某公司）" / "烟台-济南东改签差价 1.pdf"
F_BADTXT = SRC / "赵大勇（某某公司）" / "某某科技公司-烟台站通行费10.pdf"
F_TRAIN = SRC / "李强（某某公司）" / "烟台-济南  李强 198.pdf"


def chk(cond, label, extra=""):
    print(("  [OK] " if cond else "  [XX] ") + label + ("" if cond else f"   → {extra!r}"))
    if not cond:
        FAIL.append(label)


def skip(label):
    print(f"  [--] {label}（文件不在，跳过）")
    SKIP.append(label)


def section(t):
    print("\n" + "=" * 70 + f"\n{t}\n" + "=" * 70)


def parse(p: Path) -> dict:
    return ip.parse_file(p)


# ======================================================================
section("1. 纯函数：票面标记 / 单位名判定 / 税号行 / 去空白")
# ---- _stamp_from_text：红字优先、顺序不能乱 ----
chk(ip._stamp_from_text("电子客票号:123 差额退票") == "差额退票",
    "正文里的「差额退票」认得出来")
chk(ip._stamp_from_text("电子客票号:123 退票") == "退票", "正文里的「退票」认得出来")
chk(ip._stamp_from_text("电子客票号:123 始发改签") == "改签", "「始发改签」归到「改签」")
chk(ip._stamp_from_text("", "国家税务总局…监制章 差额退票") == "差额退票",
    "只在红字里写的标记也能认出来")
chk(ip._stamp_from_text("票价:￥198.00 中国铁路") == "", "普通车票没有标记")
chk(ip._stamp_from_text("差额退票 退票") == "差额退票",
    "同时出现时先判「差额退票」，不被「退票」吃掉")

# ---- _like_unit_name ----
chk(ip._like_unit_name("某某航空股份有限公司"), "航空股份公司算单位名")
chk(ip._like_unit_name("添猫科技（浙江）有限公司"), "带括号的公司名算单位名")
chk(not ip._like_unit_name("国内正常"), "「国内正常」不算单位名（原来误当购方）")
chk(not ip._like_unit_name("购方开户银行"), "「购方开户银行」不算单位名（原来误当销方）")
chk(not ip._like_unit_name("国家税务总局全国统一发票监制章浙江省税务局"),
    "监制章文字不算单位名")
chk(not ip._like_unit_name("旅客运输服务"), "「旅客运输服务」不算单位名")
chk(not ip._like_unit_name("2026年08月31日"), "日期不算单位名")

# ---- _tax_ids_of ----
ids = ip._tax_ids_of("91310101MA1FL0XXXX 91330108596628463C")
chk(len(ids) == 2, "一行里两个税号都能取到", ids)
chk(ip._tax_ids_of("26332000007500040846") == [],
    "20 位纯数字是发票号码，不能当税号")
chk(ip._tax_ids_of("*交通运输服务*客运服务费 115.67 15.68 3% 0.47") == [],
    "明细行里的数字不能当税号")
chk(ip._tax_ids_of("购买方地址:山东省济南市…;电话:-;") == [],
    "带冒号的地址/电话行不算税号行")

# ---- _despace：浏览器打印件里夹着的 NUL ----
chk(ip._despace("\x002\x006\x003\x007") == "2637", "去空白时把 \\x00 一并去掉")

# ---- _pick_parties：三种版式 ----
# ①「名称A 名称B」「税号A 税号B」两行并排（13.58 那张的版式）
rl1 = ["发票号码：", "26332000007500040846", "2026年08月31日",
       "测试用会计师事务所（普通合伙） 添猫科技（浙江）有限公司",
       "91310101MA1FL0XXXX 91330108596628463C", "旅客运输服务"]
b, bid, s, sid = ip._pick_parties(rl1, rl1)
chk((b, s) == ("测试用会计师事务所（普通合伙）", "添猫科技（浙江）有限公司"),
    "版式①：两个名字并排一行也能劈成购方/销方", (b, s))
chk(bid == "91310101MA1FL0XXXX" and sid == "91330108596628463C",
    "版式①：税号跟着配对", (bid, sid))

# ②「税号行在下、名字分两行在上」（65.67 那张的版式）
rl2 = ["91310101MA1FL0XXXX 91310115MA1K427762", "2026年08月31日", "李金丽",
       "测试用会计师事务所（普通合伙）", "享道出行（上海）科技股份有限公司",
       "26317000003130255066", "*交通运输服务*客运服务费 63.76"]
b2, bid2, s2, sid2 = ip._pick_parties(rl2, rl2)
chk((b2, s2) == ("测试用会计师事务所（普通合伙）", "享道出行（上海）科技股份有限公司"),
    "版式②：名字分开两行、号码在后面也能认全", (b2, s2))

# ③ 名字和税号一行一对
rl3 = ["北京某某科技有限公司 91110108MA01ABCD7X",
       "上海某某服务有限公司 91310115MA1K427762"]
b3, bid3, s3, sid3 = ip._pick_parties(rl3, rl3)
chk(b3 == "北京某某科技有限公司" and s3 == "上海某某服务有限公司",
    "版式③：名称+税号逐对出现", (b3, s3))

# ④ 干扰行不能污染（监制章 + 银行行）
rl4 = rl1 + ["国家税务总局全国统一发票监制章浙江省税务局",
             "购方开户银行:-;    银行账号:-;",
             "销方开户银行:杭州银行滨江支行;    银行账号:78708100557994;"]
b4, _, s4, _ = ip._pick_parties(rl4, rl4)
chk((b4, s4) == ("测试用会计师事务所（普通合伙）", "添猫科技（浙江）有限公司"),
    "监制章 / 开户银行行不会顶上购销方", (b4, s4))


# ======================================================================
section("2. 真实票据：13.58 高德票（购销方 / 事由 / 明细）")
if F_GAODE1358.exists():
    r = parse(F_GAODE1358)
    chk(r["seller"] == "添猫科技（浙江）有限公司",
        "销方=添猫科技（浙江）有限公司（原来是「购方开户银行:-;银行账号:-;」）", r["seller"])
    chk(r["buyer"] == "测试用会计师事务所（普通合伙）",
        "购方不再把两个公司名粘一起", r["buyer"])
    chk(r["buyer_id"] != r["no"], "购方税号没被发票号码顶掉", (r["buyer_id"], r["no"]))
    chk("*" not in str(r["project"]),
        "事由不再取到「*交通运输服务*客运服务费…」这串明细", r["project"])
    chk(r["item"] == "客运服务费", "服务名称不再吞掉金额数字", r["item"])
    chk(abs(r["total"] - 13.58) < 0.001, "价税合计仍是 13.58", r["total"])
else:
    skip("13.58 高德票")


section("3. 真实票据：机票行程单（销方不能是一整页文本）")
if F_FLIGHT.exists():
    r = parse(F_FLIGHT)
    chk(r["seller"] == "某某航空股份有限公司",
        "销方=某某航空股份有限公司（原来是整页文本拼起来的一串）", r["seller"][:60])
    chk(len(r["seller"]) <= 30, "销方长度正常（没有把整页塞进来）", len(r["seller"]))
    chk(r["buyer"] == "测试用会计师事务所（普通合伙）",
        "购方是测试用会计师事务所（不是票种说明「国内正常」）", r["buyer"])
    chk(r["no"] == "26378324111065015780", "补上了发票号码", r["no"])
    chk(abs(r["total"] - 748.0) < 0.01, "合计金额 748.00", r["total"])
else:
    skip("机票行程单")


section("4. 真实票据：火车票的退票 / 改签标记（票面红字）")
for path, want, label in ((F_REFUND, "退票", "退票费票"),
                          (F_DIFF, "差额退票", "差额退票票"),
                          (F_CHANGE, "改签", "始发改签票")):
    if not path.exists():
        skip(label)
        continue
    r = parse(path)
    chk(r["stamp"] == want, f"{label}：票面标记 = {want}", r["stamp"])
    chk(want in r["red"], f"{label}：红字里确实读到了「{want}」", r["red"][-20:])
if F_TRAIN.exists():
    r = parse(F_TRAIN)
    chk(abs(r["total"] - 198.0) < 0.01, "标记之外，金额照常读出来（198.00）", r["total"])


section("5. 真实票据：网页打印件（文字层夹 NUL）也要读出发票号码")
if F_BADTXT.exists():
    r = parse(F_BADTXT)
    chk(r["no"] == "26377000000470133595", "号码从 NUL 夹缝里拼回来了", r["no"])
    chk(any("编码" in w for w in r["warn"]), "并明确提示「文字层编码坏了」", r["warn"])
else:
    skip("通行费打印件")


section("6. 台账「票面标记」列")
chk("票面标记" in lg.FIELDS, "台账列里有「票面标记」")
chk(lg.COLUMNS[3][0] == "票面标记", "它排在「凭证类型」后面（第 4 列）",
    [c[0] for c in lg.COLUMNS[:5]])
chk(any(c == "票面标记" for c, _ in lg._BACKFILL_FIELDS),
    "老台账升级时会补这一列")

# 入库时把 rec 的 stamp 写进去
rec = {"ctype": "火车票", "no": "26379217601000118606", "date": "2026-08-27",
       "kind": "电子发票（铁路电子客票）", "total": 10.0, "amount": 10.0,
       "stamp": "退票", "detail": "G5367 济南东-烟台", "buyer": "测试用会计师事务所（普通合伙）",
       "seller": "中国铁路", "src": r"E:\x\10.pdf", "warn": []}
res = lg.add_records([rec], source_name="自检")
row = (res["added"] or [{}])[0]
chk(row.get("票面标记") == "退票", "入库时写进「票面标记」列", row.get("票面标记"))


section("7. 重新识别（refresh_rows）")
API = app_web.Api()
if F_GAODE1358.exists():
    rows = [{"序号": 1, "状态": "已报销", "凭证类型": "发票", "票面标记": "",
             "发票号码": "", "开票日期": "", "票种": "", "行程/明细": "",
             "购方名称": "测试用会计师事务所（普通合伙）添猫科技（浙江）有限公司",
             "销方名称": "购方开户银行:-;银行账号:-;", "不含税金额": "", "税额": "",
             "价税合计": 13.58, "项目/事由": "*交通运输服务*客运服务费 115.67 15.68 3%",
             "项目编号": "", "费用类别": "交通费", "报销人": "王小明",
             "报销批次": "9月报销单", "入库时间": "2026-09-15 10:00",
             "文件路径": str(F_GAODE1358), "提示": ""}]
    lg.save(rows, do_backup=False)
    out = API.refresh_rows(["1"])
    got = lg.load()[0]
    chk(out.get("updated") == 1, "重新识别了 1 条", out)
    chk(got["销方名称"] == "添猫科技（浙江）有限公司", "销方被刷成正确的", got["销方名称"])
    chk(got["购方名称"] == "测试用会计师事务所（普通合伙）", "购方被刷成正确的", got["购方名称"])
    chk(got["项目/事由"] == "", "误取的乱码事由被清掉", got["项目/事由"])
    chk(got["状态"] == "已报销" and got["费用类别"] == "交通费"
        and got["报销人"] == "王小明" and got["报销批次"] == "9月报销单",
        "人工填的状态/类别/报销人/批次一个没动",
        (got["状态"], got["费用类别"], got["报销人"], got["报销批次"]))
    chk(got["发票号码"] == "26332000007500040846", "号码刷进来了", got["发票号码"])

    # 手写的事由要保留
    rows2 = [dict(rows[0], **{"项目/事由": "某某公司项目现场审计费", "销方名称": "错的值"})]
    lg.save(rows2, do_backup=False)
    API.refresh_rows(["1"])
    chk(lg.load()[0]["项目/事由"] == "某某公司项目现场审计费",
        "手写过的「项目/事由」不被刷新覆盖", lg.load()[0]["项目/事由"])

    # 原件找不到 → 记 miss，不炸
    rows3 = [dict(rows[0], **{"文件路径": str(TMP / "根本没有这个文件.pdf")})]
    lg.save(rows3, do_backup=False)
    out3 = API.refresh_rows(["1"])
    chk(out3.get("updated") == 0 and out3.get("miss") == 1,
        "原件不在时记为 miss、不报错", out3)
else:
    skip("重新识别（13.58 票不在）")


section("8. 报销单明细里带上票面标记")
r = {"凭证类型": "火车票", "票面标记": "改签", "开票日期": "2026-09-02",
     "发票号码": "26379118557000832477", "费用类别": "交通费",
     "项目/事由": "", "行程/明细": "G1094 烟台-济南", "价税合计": 198.0,
     "销方名称": "中国铁路"}
vals = rp._detail_values(r, "费用报销单")
chk("（改签）" in str(vals[1]), "凭证类型写成了「火车票（改签）」", vals[1])
vals2 = rp._detail_values(dict(r, **{"票面标记": ""}), "费用报销单")
chk(vals2[1] == "火车票", "没有标记时还是干净的「火车票」", vals2[1])
d = rp.report_data([r], {"kind": "费用报销单"})
chk(list(d["by_type"].keys()) == ["火车票"], "按类型统计不受标记影响", d["by_type"])


# ======================================================================
print("\n" + "=" * 70)
if SKIP:
    print(f"跳过 {len(SKIP)} 组（源文件不在）：" + "、".join(SKIP))
print(f"❌ {len(FAIL)} 项不通过" if FAIL else "✅ 全部通过")
for f in FAIL:
    print("   -", f)
sys.exit(1 if FAIL else 0)
