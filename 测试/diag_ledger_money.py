# -*- coding: utf-8 -*-
"""
隔离自检：「金额对不上」的三件事 + 票面姓名被覆盖
====================================================
用户报（原话）：
  「这里边的发票到底多少钱。首先，识别的 4361，加了两遍打车的；而且报销的里边有
    高铁票，高铁票不是有名字吗，怎么识别不到，全都得是我自己输入的人名。」

查下来是三件**互相独立**的事，各有各的根：

  ① 「加了两遍打车的」—— 台账页的「合计」把**附件（打车行程单）**也加了一遍。
     单据那边是排除附件的（report_data 算的），台账页是自己硬加每一行，
     于是同一份数据：台账 4361.00、单据 4284.90，差的正是那张 76.10 的行程单。
     现在台账页改用**跟单据同一个口径**，并且把「另外有几张附件没算钱」写在界面上。

  ② 「4361 到底是多少」—— 它其实是**两个文件夹的票加在一起**：
     太仓批 3795.80 ＋ 烟台批 565.20。两批本来该分开报。
     现在后端按「报销批次」分组下发（by_batch），界面上一批一行、点一下只看这一批。

  ③ 「高铁票的姓名识别不到」—— 恰好相反，**票面上读得完全正确**
     （孙振强 / 邵自杰 / 陆兰玲），是被「设置出行人」批量填名字盖成了同一个名字。
     现在票面白纸黑字写着姓名的行**默认保护**，要覆盖得二次确认。

跑法：  python 测试\\diag_ledger_money.py
说明：用**用户真实台账的只读副本**跑（sqlite backup，连 WAL 一起带走），
      全程在 %TEMP% 里，绝不改原件。真实库不在时会跳过真数据那段，只跑受控那段。
"""
import io
import json
import sqlite3
import sys
import tempfile
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

REAL_DB = BASE_DIR / "发票台账.db"
REAL_CACHE = BASE_DIR / "缓存"
TMP = Path(tempfile.mkdtemp(prefix="inv_money_"))

import app_paths                                                    # noqa: E402
app_paths.DATA_DIR = TMP
app_paths.LEDGER_FILE = TMP / "发票台账.xlsx"
app_paths.LEDGER_BAK = TMP / "台账备份"
app_paths.OUTPUT_DIR = TMP / "输出"
app_paths.WORK_DIR = TMP / "缓存"
app_paths.DEBUG_DIR = TMP / "browser_debug"
app_paths.RECYCLE_DIR = TMP / "回收站"
app_paths.CONFIG_FILE = TMP / "config.json"
app_paths.LAUNCH_LOG = TMP / "launch.log"
app_paths.PASSWORD = ""

for _d in (TMP, app_paths.LEDGER_BAK, app_paths.OUTPUT_DIR, app_paths.WORK_DIR,
           app_paths.RECYCLE_DIR, TMP / "票据"):
    _d.mkdir(parents=True, exist_ok=True)
app_paths.CONFIG_FILE.write_text(
    json.dumps({"source": str(TMP / "票据"), "recursive": True}), encoding="utf-8")

HAS_REAL = REAL_DB.exists()
if HAS_REAL:
    # ⚠️ 库开了 WAL：只 copy .db 会漏掉还没 checkpoint 的行，必须走 sqlite backup
    _s = sqlite3.connect(str(REAL_DB))
    _d = sqlite3.connect(str(TMP / "发票台账.db"))
    _s.backup(_d)
    _d.close()
    _s.close()
if REAL_CACHE.is_dir():
    import shutil
    shutil.copytree(REAL_CACHE, app_paths.WORK_DIR, dirs_exist_ok=True)

import app_web                                                      # noqa: E402
import attachment                                                   # noqa: E402
import ledger as lg                                                 # noqa: E402
import report_pdf as rp                                             # noqa: E402

FAIL = []


def chk(cond, label, extra=""):
    print(("  [OK] " if cond else "  [XX] ") + label + ("" if cond else f"   → {extra}"))
    if not cond:
        FAIL.append(label)


def section(t):
    print("\n" + "=" * 66 + f"\n{t}\n" + "=" * 66)


def near(a, b, tol=0.005):
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:                                               # noqa: BLE001
        return False


api = app_web.Api()
print(f"临时目录：{TMP}\n真实台账：{'已复制副本' if HAS_REAL else '不存在（跳过真数据段）'}")

# ======================================================================
section("一、合计口径：附件不许重复计钱")
# ======================================================================
if HAS_REAL:
    r = api.get_ledger({})
    if r.get("error"):
        chk(False, "get_ledger 能读出台账", r.get("error"))
    else:
        rows = r["rows"]
        print(f"  台账 {len(rows)} 行（未删）")
        print(f"  计费合计 sum      = {r['sum']:,.2f}")
        print(f"  全部行硬加 sum_raw= {r['sum_raw']:,.2f}")
        print(f"  附件金额 attach_sum={r['attach_sum']:,.2f}（{r['attach_count']} 张）")

        # ① 恒等式：计费合计 + 附件 = 全部行硬加。这条不成立就说明口径算错了
        chk(near(r["sum"] + r["attach_sum"], r["sum_raw"]),
            "「计费合计 + 附件金额 = 全部行硬加」成立（口径没漏没错）",
            f"{r['sum']} + {r['attach_sum']} vs {r['sum_raw']}")

        # ② 跨模块口径一致：台账页的数必须和「生成单据时算出来的数」一样
        d = rp.report_data(r["rows"], {"kind": "费用报销单"})
        chk(near(r["sum"], d["total"]),
            "台账页「合计」与生成单据用的金额完全一致（以前这里是两个数）",
            f"台账 {r['sum']} vs 单据 {d['total']}")

        # ③ 锚点：这正是用户报的两个数
        chk(near(r["sum"], 4284.90),
            "计费合计 = 4,284.90（＝4361.00 − 76.10 那张行程单）", r["sum"])
        chk(near(r["sum_raw"], 4361.00),
            "全部行硬加 = 4,361.00（就是用户看到的那个数）", r["sum_raw"])
        chk(near(r["attach_sum"], 76.10),
            "被排除的附件正好是 76.10 那张打车行程单", r["attach_sum"])

        # ④ 附件行的序号必须真的下发（界面靠它把附件从合计里摘出来）
        chk(isinstance(r.get("attach_seqs"), list) and len(r["attach_seqs"]) == r["attach_count"],
            "attach_seqs 下发了（界面的合计靠它排除附件）", r.get("attach_seqs"))

# ======================================================================
section("二、按批次分开报：两批不许加成一个数")
# ======================================================================
if HAS_REAL:
    r = api.get_ledger({})
    bb = r.get("by_batch") or []
    print(f"  共 {len(bb)} 批：")
    for b in bb:
        print(f"    [{b['batch'] or '未填批次'}] {b['n']} 张（附件 {b['n_attach']}）"
              f" · {b['sum']:,.2f} · {b['first']}~{b['last']} · {b['persons']}")
    chk(len(bb) >= 2, "至少分出两批（太仓 / 烟台），不再混成一个数", len(bb))
    chk(near(sum(b["sum"] for b in bb), r["sum"]),
        "各批小计之和 = 总合计（分组没漏行也没重复算）",
        f"{sum(b['sum'] for b in bb)} vs {r['sum']}")

    by = {b["batch"]: b for b in bb}
    taicang = by.get("审计审计")
    if taicang:
        chk(near(taicang["sum"], 3795.80), "太仓批小计 = 3,795.80", taicang["sum"])
    empty = by.get("")
    if empty:
        chk(near(empty["sum"], 489.10),
            "空批次（烟台那批）小计 = 489.10（565.20 去掉 76.10 的行程单）", empty["sum"])

    # 「筛了某一批，还是能看到所有批次」—— 否则筛完就只剩一个分组，没得挑
    r2 = api.get_ledger({"batch": "审计审计"})
    chk(len(r2.get("by_batch") or []) == len(bb),
        "按批次筛过之后，批次条仍然列出全部批次（不然筛完就没得选了）",
        len(r2.get("by_batch") or []))
    chk(r2["total"] < r["total"],
        "但表格里只剩这一批的行", f"{r2['total']} vs {r['total']}")

# ======================================================================
section("三、票面姓名保护：批量填名字不许盖掉票面")
# ======================================================================
import invparse as ip                                               # noqa: E402
import traveler as tv                                               # noqa: E402

FAKE = TMP / "票据" / "邵太仓高铁去.ofd"
FAKE.write_bytes(b"fake-ofd")
CK = ip._ckey(FAKE)
FAKE2 = TMP / "票据" / "陆太仓高铁回.ofd"
FAKE2.write_bytes(b"fake-ofd-2")
CK2 = ip._ckey(FAKE2)

# 受控三行：① 票面写着邵自杰 ② 票面没人名（打车行程单） ③ 票面写着陆兰玲
_ROWS = [
    {"序号": 1, "状态": "未报销", "凭证类型": "铁路电子客票",
     "发票号码": "26379116295002204819", "开票日期": "2026-08-04", "价税合计": "456.00",
     "出行人": "邵自杰", "文件路径": str(FAKE), "行程/明细": "G2699 济南-太仓", "提示": ""},
    {"序号": 2, "状态": "未报销", "凭证类型": "打车行程单",
     "发票号码": "", "开票日期": "2026-08-04", "价税合计": "39.90",
     "出行人": "", "文件路径": str(TMP / "票据" / "打车.pdf"), "行程/明细": "太仓站-宾馆", "提示": ""},
    {"序号": 3, "状态": "未报销", "凭证类型": "铁路电子客票",
     "发票号码": "26329116565000255994", "开票日期": "2026-08-05", "价税合计": "23.00",
     "出行人": "陆兰玲", "文件路径": str(FAKE2), "行程/明细": "D2135 太仓-上海虹桥", "提示": ""},
]

LG_CALLS = []
_orig_load, _orig_patch = lg.load, lg.patch_rows
_orig_cache = ip.load_cache
lg.load = lambda: [dict(x) for x in _ROWS]
lg.patch_rows = lambda m: (LG_CALLS.append(dict(m)), len(m))[1]
# 缓存里：两张 ofd 的票面姓名分别就是「邵自杰」「陆兰玲」，来源＝票面（解析器本来就读对了）
ip.load_cache = lambda: {
    CK: {"traveler": "邵自杰", "tv_from": "票面", "src": str(FAKE)},
    CK2: {"traveler": "陆兰玲", "tv_from": "票面", "src": str(FAKE2)},
}

sec_ok = True
try:
    # 用户当时的操作：勾一片、填一个名字「孙振强」
    r = api.save_traveler_review([{"seq": "1", "name": "孙振强"},
                                  {"seq": "2", "name": "孙振强"},
                                  {"seq": "3", "name": "孙振强"}])
    chk(not r.get("error"), "save_traveler_review 正常返回", r.get("error"))
    prot = r.get("protected") or []
    chk(len(prot) == 2, "票面写了姓名的 2 张被拦下（邵自杰、陆兰玲）", prot)
    chk({p["face"] for p in prot} == {"邵自杰", "陆兰玲"},
        "拦下时带上了票面真正写的姓名（界面要列给用户看）",
        [p.get("face") for p in prot])
    chk(r["updated"] == 1, "票面没人名的那 1 张照填（不能因为保护就全都不写）", r["updated"])
    _m = LG_CALLS[-1] if LG_CALLS else {}
    chk("1" not in _m and "3" not in _m and "2" in _m,
        "真正落库的只有第 2 行：票面那两行**一个字都没动**",
        sorted(_m.keys()))

    # 用户确认之后（protect=False）：这次才覆盖
    LG_CALLS.clear()
    r2 = api.save_traveler_review([{"seq": "1", "name": "孙振强"}], True, False)
    chk(r2["updated"] == 1 and not (r2.get("protected") or []),
        "二次确认后（protect=False）允许覆盖", r2)
    chk("1" in (LG_CALLS[-1] if LG_CALLS else {}),
        "这次第 1 行确实被写成了孙振强", LG_CALLS[-1] if LG_CALLS else None)

    # 填的正好等于票面姓名 → 不算覆盖，照写
    LG_CALLS.clear()
    r3 = api.save_traveler_review([{"seq": "1", "name": "邵自杰"}])
    chk(r3["updated"] == 1 and not (r3.get("protected") or []),
        "填的就是票面上的名字 → 不算覆盖，正常写入", r3)

    # 留空（清掉）不该被拦 —— 那是明确的清空动作
    LG_CALLS.clear()
    r4 = api.save_traveler_review([{"seq": "1", "name": ""}])
    chk(r4["updated"] == 1 and r4["cleared"] == 1,
        "留空 = 明确要清掉这一格，不被保护拦住", r4)
except Exception as e:                                              # noqa: BLE001
    sec_ok = False
    chk(False, "票面姓名保护这段自己抛错了", f"{type(e).__name__}: {e}")
finally:
    lg.load, lg.patch_rows, ip.load_cache = _orig_load, _orig_patch, _orig_cache

# ======================================================================
section("四、源码守则：口径只有一处、保护不许被绕过")
# ======================================================================
src = (BASE_DIR / "app_web.py").read_text(encoding="utf-8")
js = (BASE_DIR / "gui" / "app.js").read_text(encoding="utf-8")

chk("attach_seqs" in src and '"sum_raw"' in src,
    "get_ledger 同时给出「计费合计」和「全部行硬加」（界面能解释差在哪）")
chk("by_batch" in src and "_by_batch" in src,
    "get_ledger 下发按批次分组（分开统计/分开出单的数据源）")
chk("def _face_names" in src,
    "后端有「查票面姓名」的能力（_face_names）")
chk("protect=True" in src and "tv.SRC_FACE" in src,
    "save_traveler_review 默认保护票面来源的姓名")

# 界面侧：合计不许再自己硬加
chk("function isAttach" in js and "attachSeqs" in js,
    "界面认得附件行（isAttach / attachSeqs）")
chk("function renderBatchBar" in js and "全选这批" in js,
    "界面有「按批次分开报」的批次条 + 「全选这批」")
chk("function askFaceOverwrite" in js,
    "界面有票面姓名保护的二次确认弹层")
chk("save_traveler_review', items, remember, false" in js,
    "「出行人核对」面板落库时显式 protect=false（那是逐行看过来源的）")
chk("const r = await call('save_traveler_review', items, false);" in js,
    "「设置出行人」走默认保护（protect 不传＝true）")

print("\n" + "=" * 66)
print(f"{'✅ 全部通过' if not FAIL else '❌ ' + str(len(FAIL)) + ' 项不通过'}")
if FAIL:
    for x in FAIL:
        print("   - " + x)
print("=" * 66)
sys.exit(1 if FAIL else 0)
