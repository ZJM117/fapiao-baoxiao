# -*- coding: utf-8 -*-
"""
隔离自检：出行人的「人工核对」那一半
================================================================================
用户提的问题：「三层兜底+人工核对呢？我说的不是这个吗，怎么只有三层兜底」

三层兜底只管**自动认人**（票面姓名 → 手机号对照表 → 文件夹名），认完没人过一遍
就是一句空话：手机号是译出来的、文件夹名是推出来的，都可能不对，票面读不到的
干脆是空的。这里验证补上的后半截：

  · 核对清单把每行「现在能认成谁 / 依据哪一层 / 要不要人看一眼」摆出来并排序；
  · 保存后台账落地、「提示」里的「请核对」摘掉（别的提示留着）、手机号记进对照表；
  · 核对过的记一笔「已核对」，名字再被改掉 → 提醒自动回来；
  · 认不出又留空的，永远不算核对过（角标会一直挂着）。

全部在 %TEMP% 副本里做，真实台账一个字都不写；真实票据也不碰。
跑法：  python 测试\\diag_traveler_review.py
"""
import io
import sys
import tempfile
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))

TMP = Path(tempfile.mkdtemp(prefix="inv_tvr_"))

# ⚠️ 必须在 import ledger / app_web 之前改路径（它们 from app_paths import XXX 取值）
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
import traveler as tv                                               # noqa: E402

FAIL = []


def chk(cond, label, extra=""):
    print(("  [OK] " if cond else "  [XX] ") + label + ("" if cond else f"   → {extra!r}"))
    if not cond:
        FAIL.append(label)


api = app_web.Api()

# =====================================================================
# 素材：文件只是个占位，核对时走解析缓存，不会真去读那个桩文件
# =====================================================================
MAT = TMP / "素材"
_cache = {}


def stub(folder, name, ctype="打车行程单", **kw):
    d = MAT / folder
    d.mkdir(parents=True, exist_ok=True)
    f = d / name
    f.write_bytes(b"%PDF-1.4 stub")
    r = ip._blank(str(f), "test", ctype)
    r.update(kw)
    _cache[ip._ckey(f)] = r
    return f


F_TRAIN = stub("王小明（某某公司）", "train.pdf", ctype="火车票",
               traveler="王小明", date="2026-08-19", total=204.0,
               detail="G123 济南东-烟台")
F_PH = stub("王小明（某某公司）", "trip-a.pdf", phone="13800000000",
            date="2026-08-19", total=13.58, detail="合肥南-财智中心")
F_FOLD = stub("李强（某某公司）", "trip-b.pdf", date="2026-08-20",
              total=16.0, detail="烟台站-万华")
F_BLANK = stub("报销", "trip-c.pdf", date="2026-08-21",
               total=18.0, detail="济南西-大明湖")
F_DIFF = stub("李强（某某公司）", "trip-d.pdf", phone="13900000000",
              date="2026-08-22", total=20.0, detail="烟台-蓬莱")
ip.save_cache(_cache)


def row(seq, ctype, travel, path, tip="", date="2026-08-19", total=13.58, detail=""):
    r = {f: "" for f in lg.FIELDS}
    r.update({"序号": seq, "状态": "未报销", "凭证类型": ctype, "出行人": travel,
              "开票日期": date, "价税合计": total, "行程/明细": detail,
              "费用类别": "交通费", "报销人": "张三", "文件路径": str(path), "提示": tip})
    return r


# 手机号对照表：13800000000 是王小明的（模拟自动学到 / 人工记过）
tv.save_map({"13800000000": "王小明"})

lg.save([
    row("1", "火车票", "王小明", F_TRAIN, date="2026-08-19", total=204.0,
        detail="G123 济南东-烟台"),
    row("2", "打车行程单", "王小明", F_PH, tip="出行人「王小明」由手机号对照表译出，请核对",
        date="2026-08-19", total=13.58, detail="合肥南-财智中心"),
    row("3", "打车行程单", "李强", F_FOLD, tip="出行人「李强」按文件夹名推断，请核对",
        date="2026-08-20", total=16.0, detail="烟台站-万华"),
    row("4", "打车行程单", "", F_BLANK, tip="没认出行人，可在台账里手填；重复报销风险",
        date="2026-08-21", total=18.0, detail="济南西-大明湖"),
    row("5", "打车行程单", "张伟", F_DIFF, date="2026-08-22", total=20.0,
        detail="烟台-蓬莱"),
])

# =====================================================================
print("\n=== 1. 核对记录：读写 + 改名字就失效 ===")
tv.mark_checked({str(F_TRAIN): "王小明"})
chk(tv.is_checked(str(F_TRAIN), "王小明"), "记过的名字 → 算核对过")
chk(not tv.is_checked(str(F_TRAIN), "别的名字"), "台账里的名字被改掉 → 核对标记失效")
chk(not tv.is_checked(str(F_TRAIN), ""), "名字被清空 → 不算核对过")
chk(not tv.is_checked(str(F_TRAIN), "王小明") is False
    and tv.checked_path().exists(), "核对记录落在 缓存/出行人核对.json")
chk(tv.mark_checked({str(F_TRAIN): ""}) == {}, "名字给空 = 撤掉这条核对标记")
chk(not tv.is_checked(str(F_TRAIN), "王小明"), "撤掉之后就不算核对过了")

# =====================================================================
print("\n=== 2. 核对清单：谁是怎么认出来的、要不要人看 ===")
rev = api.traveler_review()
chk(not rev.get("error"), "核对清单能读出来", rev.get("error"))
by = {x["seq"]: x for x in (rev.get("rows") or [])}
chk(len(by) == 5, "5 张全在清单里", list(by))

chk(by["1"]["src"] == "票面" and by["1"]["state"] == "ok" and not by["1"]["need"],
    "火车票：票面读到姓名 → 默认不用人工核", by.get("1"))
chk(by["2"]["src"] == "手机号" and by["2"]["need"],
    "行程单：手机号译出来的 → 要核一下", by.get("2"))
chk(by["3"]["src"] == "文件夹" and by["3"]["need"],
    "行程单：文件夹名推出来的 → 要核一下", by.get("3"))
chk(by["4"]["state"] == "blank" and by["4"]["auto"] == "" and by["4"]["need"],
    "认不出来的 → 标 blank，必须手填", by.get("4"))
chk(by["5"]["state"] == "diff" and by["5"]["auto"] == "李强" and by["5"]["cur"] == "张伟",
    "台账里是「张伟」、重算应是「李强」→ 标 diff 摆出来", by.get("5"))

st = rev.get("stat") or {}
chk(st.get("total") == 5 and st.get("need") == 4 and st.get("blank") == 1,
    "统计：5 张 / 4 张待核对 / 1 张没认出来", st)
chk(st.get("手机号") == 1 and st.get("文件夹") == 2 and st.get("票面") == 1,
    "来源分布：票面 1、手机号 1、文件夹 2", st)

order = [x["seq"] for x in rev["rows"]]
chk(order[0] == "4", "没认出来的排最前面（最急）", order)
chk(order.index("5") < order.index("2"), "「跟自动值不符」排在「手机号译出」前面", order)
chk(order[-1] == "1", "不用核的排最后", order)

cands = rev.get("people") or []
chk("王小明" in cands and "李强" in cands and "张伟" in cands,
    "候选名单含台账里出现过的名字（界面拿它做补全）", cands)
chk(by["2"]["phone"] == "13800000000", "清单里带上了票面手机号（好看清是哪部手机开的）",
    by["2"].get("phone"))

# =====================================================================
print("\n=== 3. 保存核对结果：台账落地 + 提示摘干净 + 学手机号 ===")
items = [
    {"seq": "2", "name": "王小明", "phone": "13800000000"},
    {"seq": "3", "name": "李强", "phone": ""},
    {"seq": "4", "name": "王小明", "phone": ""},
    {"seq": "5", "name": "李强", "phone": "13900000000"},
]
res = api.save_traveler_review(items, remember=True)
chk(res.get("updated") == 4, "4 行写进台账了", res)
after = {str(r.get("序号")): r for r in lg.load()}
chk(after["4"]["出行人"] == "王小明", "原本空着的填上了「王小明」", after["4"].get("出行人"))
chk(after["5"]["出行人"] == "李强", "原本是「张伟」的改成了「李强」", after["5"].get("出行人"))
chk(after["4"]["提示"] == "重复报销风险",
    "「提示」里跟出行人有关的那句摘掉了，别的提示原样留着", after["4"].get("提示"))
chk(after["2"]["提示"] == "", "纯「请核对」的提示变成空", after["2"].get("提示"))
chk(after["1"]["出行人"] == "王小明" and after["1"]["提示"] == "",
    "没提交的行一个字没动", (after["1"].get("出行人"), after["1"].get("提示")))

pm = tv.load_map()
chk(pm.get("13900000000") == "李强", "勾了「记住手机号」→ 写进对照表，下次自动认人", pm)
chk(res.get("remembered") == 2, "记住的手机号个数对得上（13800000000 本来就有，重复记不算新增）",
    res)
chk(tv.is_checked(str(F_FOLD), "李强"), "核对过的记了一笔「已核对」")

# =====================================================================
print("\n=== 4. 核完再进一次：不再提醒；改了名字提醒回来 ===")
rev2 = api.traveler_review()
by2 = {x["seq"]: x for x in rev2["rows"]}
chk(by2["2"]["done"] and by2["3"]["done"] and by2["4"]["done"] and by2["5"]["done"],
    "这 4 张都标成「已核对」了", {k: v["done"] for k, v in by2.items()})
chk(rev2["stat"]["need"] == 0, "待核对归零 → 按钮角标会消失", rev2["stat"])
chk([x["seq"] for x in rev2["rows"] if x["need"]] == [],
    "没有一行挂着「待核对」")

lg.update_rows(["3"], **{"出行人": "赵四"})            # 核对完又被改掉了
rev3 = api.traveler_review()
by3 = {x["seq"]: x for x in rev3["rows"]}
chk(not by3["3"]["done"] and by3["3"]["need"] and by3["3"]["state"] == "diff",
    "名字被改掉 → 核对标记失效、重新变成待核对", by3["3"])

# 全部都核对完的那一轮，按钮角标应该是 0
chk(api.get_ledger({}).get("tv_todo") == 0,
    "上个用例（全部核对过）时角标是 0 —— 核对过的就不再挂数字",
    api.get_ledger({}).get("tv_todo"))

# 留空的不算核对过
lg.update_rows(["4"], **{"出行人": ""})
rev4 = api.traveler_review()
by4 = {x["seq"]: x for x in rev4["rows"]}
chk(by4["4"]["state"] in ("blank", "todo") and by4["4"]["need"] and not by4["4"]["done"],
    "台账被清空 → 重新变成待核对（不会因为「核对过」就跳过）", by4["4"])
api.save_traveler_review([{"seq": "4", "name": "", "phone": ""}], remember=False)
rev5 = {x["seq"]: x for x in api.traveler_review()["rows"]}
chk(rev5["4"]["need"] and not rev5["4"]["done"],
    "核对时留空 = 撤掉核对标记，下次还提醒（空着不算核对完）", rev5["4"])

# =====================================================================
print("\n=== 5. 台账角标：还差几张没落实 ===")
g = api.get_ledger({})
expect = sum(1 for r in lg.load()
             if not str(r.get("出行人") or "").strip()
             or "请核对" in str(r.get("提示") or ""))
chk(g.get("tv_todo") == expect, f"角标数 = 空着的 + 标了「请核对」的（{expect} 张）", g.get("tv_todo"))
chk(expect >= 1, "4 号空着 → 角标还挂着（3 号改过名不算：那已经等于人工处理过了）", expect)
chk(g.get("tv_todo") == 1, "此时只剩 4 号一张挂着", g.get("tv_todo"))

# =====================================================================
print("\n=== 6. 容错：原件不在、路径空着也不能崩 ===")
lg.save([row("1", "打车行程单", "", MAT / "没有这个文件夹" / "没有这个文件.pdf")])
rev6 = api.traveler_review()
chk(not rev6.get("error"), "原件不在时清单照样出得来", rev6.get("error"))
chk(rev6["rows"][0]["state"] == "blank", "读不到票面 + 文件夹认不出人 → blank", rev6["rows"][0])
api.save_traveler_review([{"seq": "1", "name": "李四", "phone": ""}], remember=False)
chk({str(r.get("序号")): r for r in lg.load()}["1"]["出行人"] == "李四",
    "原件不在也能人工填上名字")

# =====================================================================
print("\n" + "=" * 76)
if FAIL:
    print(f"❌ {len(FAIL)} 项没过：")
    for x in FAIL:
        print("   · " + x)
else:
    print("✅ 全部通过")
print(f"（临时目录：{TMP}）")
sys.exit(1 if FAIL else 0)
