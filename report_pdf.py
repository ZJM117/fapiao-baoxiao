# -*- coding: utf-8 -*-
"""
报销单生成（PDF / Excel 双格式）+ 发票 PDF 合并
================================================
两条路都刻意避开「装一堆 PDF 排版库」：

1. 报销单 PDF：先拼一张 HTML，再交给**无头 Edge 打印成 PDF**。
   好处：中文字体、A4 版心、表格跨页都由浏览器保证，不用管字体嵌不进 PDF 的坑；
   实测 3.6 秒出一张，且全程无窗口、不抢焦点。

2. 报销单 Excel：用 openpyxl 直接画（台账本来就是 Excel，同一套依赖，不再多装东西）。
   版式与 PDF 一致，方便用户拿去再改、再算，或者直接打印。

3. 发票合并：用 pypdf 把每两张发票缩放到 A4 的上下两半，中间画一条裁切线
   （按用户要求：一张 A4 上下两个，中间裁开）。裁切线本身也是「Edge 打一张带虚线的
   A4 空白页」当水印叠上去，省得引入画图库。

报销单的栏目是照着「比较全面的报销单」定的（用户 2026-09-15 要求「金额要全面」）：
  抬头 → 基本信息（报销人/部门/日期/张数/事由）→ 费用明细（含凭证类型、摘要、发票号码、
  销方）→ 金额汇总（按费用类别，分 不含税 / 税额 / 价税合计 三列）→ 合计大写+小写
  → 附件清单 → 预借款 / 应退补 → 签字栏 → 制表信息
"""
import os
import shutil
import subprocess
import tempfile
import time
import uuid
from datetime import datetime
from pathlib import Path

import attachment
from app_paths import EDGE, WORK_DIR

A4_W, A4_H = 595.28, 841.89        # A4 尺寸（点）
MARGIN = 8.0                        # 上下半页各自的留白

_CUTLINE_CACHE = {}


# ---------------------------------------------------------------- 人民币大写
def rmb_upper(amount) -> str:
    """
    14310.29 → 壹万肆仟叁佰壹拾圆贰角玖分

    写法按「数电票票面实际口径」定的（拿 136 张真实发票的官方大写字段回归挑出来的，
    见 tools/test_rmb.py 的对照结果）：
      · 用「圆」不用「元」
      · 元位是 0 且角位不为 0 时，角前补一个「零」（1820.35 → …贰拾圆零叁角伍分）
      · 角位不为 0、分位为 0 时，角后**不写**「整」（1849.40 → …肆角）
      · 分位为 0 且角位也为 0 时才写「整」（1314.00 → …圆整）
      · 负数用「（负数）」前缀，这是税局对红字发票的写法
    这样 136 张里能对上 135 张（剩 1 张是开票端自己写成了「元」）。
    """
    try:
        num = round(float(amount or 0), 2)
    except (TypeError, ValueError):
        return ""
    if num == 0:
        return "零圆整"
    neg = num < 0
    cents = int(round(abs(num) * 100))
    yuan, jiao, fen = cents // 100, cents // 10 % 10, cents % 10

    D = "零壹贰叁肆伍陆柒捌玖"
    U = ("", "拾", "佰", "仟")
    B = ("", "万", "亿", "万亿")

    groups, y = [], yuan
    while y > 0:
        groups.append(y % 10000)
        y //= 10000

    out, zero_flag = "", False
    for gi in range(len(groups) - 1, -1, -1):
        seg = groups[gi]
        if seg == 0:
            zero_flag = zero_flag or bool(out)
            continue
        seg_s, pending = "", False
        for i in range(4):
            d = (seg // 10 ** (3 - i)) % 10
            if d == 0:
                pending = True
            else:
                if pending and seg_s:
                    seg_s += D[0]
                pending = False
                seg_s += D[d] + U[3 - i]
        if out and (zero_flag or seg < 1000):
            out += D[0]
        out += seg_s + B[gi]
        zero_flag = False

    # 不足一元时不写「圆」（0.50 → 伍角整）
    s = (out + "圆") if yuan else ""
    if jiao == 0 and fen == 0:
        return ("（负数）" if neg else "") + s + "整"
    if jiao:
        if yuan and yuan % 10 == 0:      # 元位是 0，角前补零
            s += D[0]
        s += D[jiao] + "角"
        if fen:
            s += D[fen] + "分"
    else:
        s += D[0] + D[fen] + "分"
    return ("（负数）" if neg else "") + s


# ---------------------------------------------------------------- HTML → PDF
def _run_edge_print(tmp_html: Path, out_path: Path, profile, timeout: int = 120):
    """起一个无头 Edge/Chromium 把 HTML 打成 PDF（无窗口、不抢焦点、不留页眉页脚）"""
    cmd = [EDGE, "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
           f"--user-data-dir={profile}", "--no-pdf-header-footer", "--disable-extensions"]
    if os.name != "nt":
        # 容器里（Docker/NAS）是以 root 跑的：Chromium 默认拒绝给 root 开沙箱，不加这个
        # 参数会直接启动失败 → 表现成「没有产出文件」。另外容器 /dev/shm 默认只有 64MB，
        # 明细一多就会崩，所以让它别用 /dev/shm。
        # 这里只是把一段自己生成的静态 HTML 打成 PDF，不开沙箱没有额外风险。
        cmd += ["--no-sandbox", "--disable-dev-shm-usage"]
    cmd += [f"--print-to-pdf={out_path}", tmp_html.as_uri()]
    si = None
    if hasattr(subprocess, "STARTUPINFO"):
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    subprocess.run(cmd, capture_output=True, timeout=timeout, startupinfo=si,
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def html_to_pdf(html: str, out_path, timeout: int = 120):
    """
    无头 Edge 把 HTML 打成 A4 PDF。

    ⚠️ Edge 是「同一个 --user-data-dir 只允许一个实例」的：上一次打印的 Edge 要是还没退干净，
    新起的这个会把命令行**转交给老实例然后自己退出** —— 表面上什么都没发生，PDF 也没落盘，
    报出来就是「无头 Edge 没有产出文件」，看着像 Edge 被策略禁用了（2026-09-15 实际撞到过，
    连着出两张单子时第二张就没了）。所以：第 1 遍用公用 profile（快），
    没出文件就换一个一次性 profile 再打一遍（稳）；还不行就歇一下再来第三遍
    （上一轮被打断时会留个 Edge 半死不活地占着，等它彻底退干净就好了）。
    「有没有出文件」靠 out_path 落盘判断，不看 Edge 的返回码 —— 它转交命令行时也是 rc=0。
    """
    out_path = Path(os.path.abspath(out_path))
    # ⚠️ 一定要把 out_path 变成绝对路径：Edge 落盘时用的是**它自己进程的工作目录**
    # （Windows 上它会转交给 broker 进程，cwd 就不受我们控制了），
    # 传相对路径的话 PDF 会写到别处、表现成「没出文件」（2026-09-16 实际踩到）。
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    tmp_html = WORK_DIR / f"_print_{uuid.uuid4().hex[:8]}.html"
    tmp_html.write_text(html, encoding="utf-8")
    tmp_prof = None
    try:
        if out_path.exists():
            out_path.unlink()
        # 第 1 遍公用 profile（快）；第 2、3 遍各起一个一次性 profile
        for attempt in (1, 2, 3):
            if attempt == 1:
                profile = WORK_DIR / "_edge_print_profile"
            else:
                if attempt == 3:
                    time.sleep(1.5)          # 给残留的 Edge 一点时间退干净
                tmp_prof = Path(tempfile.mkdtemp(prefix="edge_print_"))
                profile = tmp_prof
            _run_edge_print(tmp_html, out_path, profile, timeout)
            # Edge 打印是异步落盘的，等一下文件出现（后两遍多等一会儿 —— 机器忙的时候冷启动慢）
            for _ in range(48 if attempt == 1 else 120):
                if out_path.exists() and out_path.stat().st_size > 800:
                    break
                time.sleep(0.25)
            if out_path.exists() and out_path.stat().st_size > 800:
                break
    finally:
        try:
            tmp_html.unlink()
        except OSError:
            pass
        if tmp_prof is not None:
            shutil.rmtree(tmp_prof, ignore_errors=True)
    if not out_path.exists() or out_path.stat().st_size <= 800:
        raise RuntimeError("生成 PDF 失败：无头 Edge 没有产出文件"
                           "（连着出单子、或上次打断了留了 Edge 没退干净时最容易碰上，"
                           "已经自动重试 3 遍了；再不行就把 Edge 全退出、或检查策略是否禁用了无头打印）")
    return out_path


# ---------------------------------------------------------------- 报销单：数据
def _esc(s) -> str:
    return (str(s if s is not None else "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _money(v) -> str:
    try:
        return f"{float(v):,.2f}"
    except (TypeError, ValueError):
        return ""


def _safe(v):
    try:
        return float(str(v).replace(",", "") or 0)
    except ValueError:
        return 0.0


def _days_between(start: str, end: str) -> str:
    """出差起止 → 天数（填了才算）"""
    try:
        a = datetime.strptime(str(start).strip()[:10], "%Y-%m-%d")
        b = datetime.strptime(str(end).strip()[:10], "%Y-%m-%d")
    except (ValueError, TypeError):
        return ""
    d = (b - a).days + 1
    return f"{d} 天" if d > 0 else ""


def _ctype_of(r) -> str:
    c = str(r.get("凭证类型") or "").strip()
    if c:
        return c
    return "发票" if str(r.get("发票号码") or "").strip() else "其他凭证"


# ---------------------------------------------------------------- 单据版式
# 「模板二」是用户 2026-09-15 提的第二套模板，样子照
# E:\桌\报销明细-某某公司-7.28.xlsx 的「报销单」工作表复刻：标题「出差报销明细」，
# 明细六列（出行人 / 类型 / 日期 / 行程 / 报销金额 / 备注），末尾「小计」，
# 再接一块**按人**的出差补助（每人：补助金额 ＋ 这个人名下票据合计 ＝ 实际），最后「合计」。
JY_KIND = "模板二"
JY_TITLE = "出差报销明细"


def style_of(kind) -> str:
    """单据类型 → 版式：cost（费用报销单）/ trip（差旅费报销单）/ jy（模板二版）。"""
    k = str(kind or "").strip()
    if k == JY_KIND:
        return "jy"
    if k == "差旅费报销单":
        return "trip"
    return "cost"


# 模板里「类型」列写的是 高铁票 / 打车 / 住宿 / 公交 这种短标签，
# 台账里存的是 凭证类型（火车票 / 打车行程单 / 飞机行程单 / 发票），两边得对一下。
_CAT_HINT = (("住宿", "住宿"), ("餐饮", "餐费"), ("餐费", "餐费"), ("办公", "办公"),
             ("交通", "交通"), ("话费", "话费"), ("通行", "过路费"), ("加油", "加油"))


def _short_type(r, trip="") -> str:
    """台账行 → 模板「类型」列的短标签。"""
    ctype = _ctype_of(r)
    trip = trip or str(r.get("行程/明细") or "").strip()
    cat = str(r.get("费用类别") or "").strip()
    if ctype == "火车票":
        # G/D/C 开头的车次是高铁动车，其余按普速车票写
        return "高铁票" if trip[:1].upper() in ("G", "D", "C") else "火车票"
    if ctype in ("飞机行程单", "登机牌"):
        return "机票"
    if ctype == "打车行程单":
        for k in ("公交", "地铁", "轨道", "巴士"):
            if k in trip:
                return "公交"
        return "打车"
    for k, v in _CAT_HINT:              # 普通发票：先看费用类别，再看行程里有没有线索
        if k in cat:
            return v
    for k, v in _CAT_HINT:
        if k in trip:
            return v
    return ctype or "发票"


def _alw_by_items(meta: dict) -> list:
    """
    模板二版式的差旅补助：**按人**填，一个出行人一行（跟模板一致）。

    meta["allowance_by"] = [{"name": "聂法银", "amount": 1200}, ...]
    票据合计不在这里算 —— 那是台账里的事实，由 report_data 按出行人分组加出来。
    """
    out = []
    for it in (meta.get("allowance_by") or []):
        if not isinstance(it, dict):
            continue
        name = str(it.get("name") or "").strip()
        amt = round(_safe(it.get("amount")), 2)
        if not name and amt <= 0:
            continue
        out.append({"name": name or "（未填）", "amount": amt})
    # 同名合并（界面上正常不会重复，防手滑）
    merged = {}
    for a in out:
        merged[a["name"]] = round(merged.get(a["name"], 0.0) + a["amount"], 2)
    return [{"name": k, "amount": v} for k, v in merged.items()]


def _alw_items(meta: dict) -> list:
    """
    差旅费补助（用户要求：「要能够增加差旅费，因为出差有差旅费啊，要计算几个人几天，
    多少钱，可以我自己设置这个，我自己填写」）。

    界面上一行 = 一个补助项目：项目 / 人数 / 天数 / 标准（元/人·天）/ 金额。
    金额 = 人数 × 天数 × 标准（界面上实时算给你看），也允许直接改金额（手改优先）。
    金额为 0 又没填项目名的行直接丢掉 —— 界面默认摆的那行空白不会污染单据。
    """
    out = []
    for it in (meta.get("allowance") or []):
        if not isinstance(it, dict):
            continue
        name = str(it.get("name") or "").strip()
        people = _safe(it.get("people"))
        days = _safe(it.get("days"))
        rate = _safe(it.get("rate"))
        amt = _safe(it.get("amount"))
        if amt <= 0:                        # 金额没手改过 → 按 人数 × 天数 × 标准 算
            amt = round(people * days * rate, 2)
        if amt <= 0 and not name:
            continue
        out.append({"name": name or "差旅费补助",
                    "people": f"{people:g}" if people else "",
                    "days": f"{days:g}" if days else "",
                    "rate": round(rate, 2), "amount": round(amt, 2)})
    return out


def report_data(rows: list, meta: dict) -> dict:
    """
    报销单要用到的全部数据一次算好（PDF 和 Excel 两个版式共用，保证两边数字一模一样）。

    ⚠️ 不含税 / 税额的补齐规则：行程单、登机牌这类凭证台账里只有一个「价税合计」，
    不含税和税额是空的。如果直接相加，会出现「不含税 + 税额 ≠ 价税合计」的对不上账。
    所以这里统一口径：以价税合计为准，缺的差额补进不含税金额。

    ⚠️ 附件不重复计钱（用户 2026-09-15 要求）：配到发票上的打车行程单只是那张发票的
    证明附件，钱已经算在发票里了，再单独加一遍就成了双倍（76.10 会变成 152.20）。
    所以「金额」相关的小计只算报销凭证（bills），附件只出现在明细表里、标「（附件）」。
    """
    meta = meta or {}
    kind = (meta.get("kind") or "费用报销单").strip()

    # 附件归位：行程单挪到它对应发票的紧后面，并算出「哪些行是配上的附件」
    att = attachment.analyze(sorted(list(rows or []), key=attachment.row_sort_key))
    rows, attach_seqs = att["rows"], att["attach_seqs"]
    bills = [r for r in rows if str(r.get("序号") or "") not in attach_seqs]

    total = round(sum(_safe(r.get("价税合计")) for r in bills), 2)
    amount = round(sum(_safe(r.get("不含税金额")) for r in bills), 2)
    tax = round(sum(_safe(r.get("税额")) for r in bills), 2)
    if abs(amount + tax - total) > 0.01:
        amount = round(total - tax, 2)

    # 张数按「实际打印出来的凭证张数」算（含附件，和明细表、附件清单对得上）；
    # 金额只算报销凭证，见上面的说明。
    by_cat = {}
    for r in rows:
        k = str(r.get("费用类别") or "").strip() or "(未分类)"
        d = by_cat.setdefault(k, {"n": 0, "amount": 0.0, "tax": 0.0, "total": 0.0})
        d["n"] += 1
        if str(r.get("序号") or "") in attach_seqs:
            continue
        d["amount"] += _safe(r.get("不含税金额"))
        d["tax"] += _safe(r.get("税额"))
        d["total"] += _safe(r.get("价税合计"))
    for v in by_cat.values():
        v["amount"] = round(v["amount"], 2)
        v["tax"] = round(v["tax"], 2)
        v["total"] = round(v["total"], 2)
        if abs(v["amount"] + v["tax"] - v["total"]) > 0.01:
            v["amount"] = round(v["total"] - v["tax"], 2)

    by_type = {}
    for r in rows:
        t = _ctype_of(r)
        by_type[t] = by_type.get(t, 0) + 1

    style = style_of(kind)
    if style == "jy":
        # 模板二版式：补助是**按人**填的（一个出行人一行），不走「人数×天数×标准」那块
        alw, alw_by = [], _alw_by_items(meta)
        alw_total = round(sum(x["amount"] for x in alw_by), 2)
        if alw_total <= 0:
            alw_by = []                 # 一分补助都没填 → 整块不出现，免得单子上挂一排 0
    else:
        alw, alw_by = _alw_items(meta), []
        alw_total = round(sum(x["amount"] for x in alw), 2)
    grand = round(total + alw_total, 2)

    date = meta.get("date") or datetime.now().strftime("%Y年%m月%d日")
    n = len(rows)
    attach = "、".join(f"{t} {c} 张" for t, c in
                       sorted(by_type.items(), key=lambda kv: -kv[1])) or "—"
    if attach_seqs:
        attach += (f"，合计 {n} 张（其中 {len(attach_seqs)} 张行程单是发票的证明附件，"
                   f"金额已含在对应发票里，不重复计）")
    else:
        attach += f"，合计 {n} 张。"

    # 出行人：明细里出现过的人（按首次出现排序），用来做「出差人」的兜底 ——
    # 一份差旅单里往往不止一个人，让用户自己一个个填不现实。
    travelers = []
    for r in rows:
        who = str(r.get("出行人") or "").strip()
        if who and who not in travelers:
            travelers.append(who)
    person = str(meta.get("person") or "").strip()
    if not person and kind == "差旅费报销单" and travelers:
        person = "、".join(travelers[:4]) + (f" 等 {len(travelers)} 人" if len(travelers) > 4 else "")
    elif not person and style == "jy" and travelers:
        # 模板二的抬头只有一格「报销人」，一个人就写名字，多人就写「甲、乙 等 N 人」
        person = travelers[0] if len(travelers) == 1 else \
            "、".join(travelers[:3]) + (f" 等 {len(travelers)} 人" if len(travelers) > 3 else "")

    # 按人汇总（模板二版式）：每个人的票据合计 ＋ 这个人自己的补助 ＝ 实际。
    # ⚠️ 票据合计只算报销凭证（bills），附件（打车行程单）的钱已经含在发票里了，不能加第二遍。
    sums = {}
    for r in bills:
        who = str(r.get("出行人") or "").strip() or "（未标注）"
        sums[who] = round(sums.get(who, 0.0) + _safe(r.get("价税合计")), 2)
    order = list(sums.keys())
    alw_map = {}
    for a in alw_by:
        alw_map[a["name"]] = round(alw_map.get(a["name"], 0.0) + a["amount"], 2)
        if a["name"] not in sums:        # 只发生补助、没票的人也要出现在名单里
            order.append(a["name"])
    by_person = [{"name": nm, "bills": sums.get(nm, 0.0), "alw": alw_map.get(nm, 0.0),
                  "real": round(sums.get(nm, 0.0) + alw_map.get(nm, 0.0), 2)} for nm in order]

    return {
        "kind": kind, "n": n, "date": date,
        "style": style, "title": JY_TITLE if style == "jy" else kind,
        "rows": rows, "attach_seqs": attach_seqs, "n_attach": len(attach_seqs),
        "link_of": att.get("link_of") or {},
        "n_bill": len(bills),
        "total": total, "amount": amount, "tax": tax,
        "bill_upper": rmb_upper(total),
        "alw": alw, "alw_total": alw_total,
        "alw_by": alw_by, "by_person": by_person,
        "grand": grand, "grand_upper": rmb_upper(grand),
        # 「报销金额（大写）」= 凭证合计 + 差旅费补助，两个版式都用它
        "upper": rmb_upper(grand),
        "by_cat": by_cat, "by_type": by_type, "attach": attach,
        "days": _days_between(meta.get("start", ""), meta.get("end", "")),
        "person": person, "dept": meta.get("dept", ""),
        "travelers": travelers,
    }


_COST_COLS = [("序号", "5%", "c"), ("开票日期", "9%", "c"), ("凭证类型", "9%", "c"),
              ("出行人", "8%", "c"),
              ("费用类别", "8%", "c"), ("事由 / 摘要", "18%", ""), ("发票号码", "16%", "c"),
              ("销方名称", "15%", ""), ("金额（元）", "12%", "r")]
_TRIP_COLS = [("序号", "5%", "c"), ("日期", "9%", "c"), ("凭证类型", "9%", "c"),
              ("出行人", "8%", "c"),
              ("行程 / 事由", "23%", ""), ("发票号码", "17%", "c"),
              ("销方名称", "17%", ""), ("金额（元）", "12%", "r")]


def _detail_values(r, kind, attach=False) -> list:
    """明细行里除「序号」以外的单元格值，顺序必须跟 _COST_COLS / _TRIP_COLS 对齐。
    attach=True 表示这一行是「跟在发票后面的证明附件」（打车行程单），
    凭证类型上标出来、金额列照旧显示（方便核对），但不计入合计（见 report_data 说明）。"""
    no = str(r.get("发票号码") or "").strip()
    # 票面标记（退票 / 差额退票 / 改签 / 红字）跟在凭证类型后面 —— 财务看单子时必须
    # 一眼看到「这张不是普通票价」，否则退票费会被当成车票钱报掉。
    mark = str(r.get("票面标记") or "").strip()
    ctype = _ctype_of(r) + (f"（{mark}）" if mark else "") + ("（附件）" if attach else "")
    trip = str(r.get("行程/明细") or "").strip()
    proj = str(r.get("项目/事由") or "").strip()
    # 出行人：火车票的乘车人 / 机票旅客 / 打车票出行人；票面读不到就靠手机号或文件夹
    # 兜底推出来（见 traveler.py），推不出来就留空 —— 不瞎猜。
    who = str(r.get("出行人") or "").strip()
    if kind == "差旅费报销单":
        return [r.get("开票日期", ""), ctype, who, trip or proj, no,
                r.get("销方名称", ""), _money(r.get("价税合计"))]
    return [r.get("开票日期", ""), ctype, who, r.get("费用类别", ""), proj or trip, no,
            r.get("销方名称", ""), _money(r.get("价税合计"))]


# ---------------------------------------------------------------- 报销单：HTML
_BASE_CSS = """
@page { size: A4 portrait; margin: 12mm 10mm 10mm 10mm; }
* { box-sizing: border-box; }
body { font-family: "Microsoft YaHei", "SimSun", sans-serif; color: #111; font-size: 11px;
       -webkit-print-color-adjust: exact; }
h1 { text-align: center; font-size: 19px; letter-spacing: 8px; margin: 0 0 4px; font-weight: 600; }
.sub { text-align: center; font-size: 9.5px; color: #777; margin: 0 0 8px; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; }
th, td { border: 1px solid #444; padding: 3px 4px; font-size: 10px; line-height: 1.32;
         word-break: break-all; vertical-align: middle; }
th { background: #f0f2f5; font-weight: 600; text-align: center; }
td.c { text-align: center; }
td.r { text-align: right; font-variant-numeric: tabular-nums; }
thead { display: table-header-group; }
tr { page-break-inside: avoid; }
tr.sum td { font-weight: 700; background: #fafafa; }
.sec { margin: 9px 0 4px; font-size: 11px; font-weight: 600;
       border-left: 3px solid #333; padding-left: 6px; }
.info th { width: 58px; background: #f7f8fa; text-align: center; font-weight: 600; }
.info td { height: 21px; padding-left: 6px; }
.money { margin-top: 6px; display: flex; }
.money .l { flex: 1; border: 1px solid #444; border-right: none; padding: 6px 8px; }
.money .r { border: 1px solid #444; padding: 6px 8px; text-align: right;
            white-space: nowrap; font-weight: 600; }
.attach { margin-top: 6px; border: 1px solid #444; padding: 5px 8px; font-size: 10px; }
.sign { margin-top: 12px; display: flex; border: 1px solid #444; }
.sign > div { flex: 1; border-right: 1px solid #444; padding: 4px 6px 30px; font-size: 10px; }
.sign > div:last-child { border-right: none; }
.foot { margin-top: 8px; font-size: 9px; color: #666;
        display: flex; justify-content: space-between; }
/* 附件行（跟在发票后面的打车行程单）：淡底 + 灰字，一眼看出它不是独立的报销项 */
tr.att td { background: #fbfbfb; color: #555; }
.note { margin-top: 4px; font-size: 9px; color: #666; line-height: 1.4; }
"""


def build_report_html(rows: list, meta: dict) -> str:
    if style_of((meta or {}).get("kind")) == "jy":
        return build_jy_html(rows, meta)          # 模板二版式另开一套，见下面
    d = report_data(rows, meta or {})
    kind, n = d["kind"], d["n"]
    trip_kind = kind == "差旅费报销单"
    cols = _TRIP_COLS if trip_kind else _COST_COLS

    # ---------- 基本信息区 ----------
    who = "出差人" if trip_kind else "报销人"
    info1 = [(who, d["person"]), ("部门", d["dept"]),
             ("报销日期", d["date"]), ("单据张数", f"{n} 张")]
    row1 = "".join(f"<th>{t}</th><td>{_esc(v) or '&nbsp;'}</td>" for t, v in info1)
    reason_label = "出差事由" if trip_kind else "报销事由"
    info = f'<table class="info"><tr>{row1}</tr>'
    info += f'<tr><th>{reason_label}</th><td colspan="7">{_esc(meta.get("reason", "")) or "&nbsp;"}</td></tr>'
    if trip_kind:
        info += ('<tr><th>出差起止</th><td colspan="3">'
                 + (_esc(f"{meta.get('start','')} 至 {meta.get('end','')}".strip(" 至")) or "&nbsp;")
                 + '</td><th>出差地点</th><td>' + (_esc(meta.get("place", "")) or "&nbsp;")
                 + '</td><th>出差天数</th><td class="c">' + (_esc(d["days"]) or "&nbsp;")
                 + "</td></tr>")
    info += "</table>"

    # ---------- 费用明细 ----------
    th = "".join(f'<th style="width:{w}">{t}</th>' for t, w, _ in cols)
    body = []
    for i, r in enumerate(d["rows"], start=1):
        at = str(r.get("序号") or "") in d["attach_seqs"]
        vals = [i, *_detail_values(r, kind, attach=at)]
        body.append(("<tr class='att'>" if at else "<tr>") + "".join(
            f'<td class="{cls}">{_esc(v)}</td>' for (t, w, cls), v in zip(cols, vals)) + "</tr>")
    if not body:
        body.append(f'<tr><td colspan="{len(cols)}" class="c" style="height:34px">（未选择凭证）</td></tr>')
    footnote = ""
    if d["n_attach"]:
        footnote = (f'<div class="note">注：标记「（附件）」的 {d["n_attach"]} 张行程单是上方发票的'
                    f'证明附件（金额以对应发票为准），已计在发票行内，不再重复计入合计。</div>')

    # ---------- 金额汇总（按费用类别）----------
    cat = "".join(
        f'<tr><td>{_esc(k)}</td><td class="c">{v["n"]}</td>'
        f'<td class="r">{_money(v["amount"])}</td><td class="r">{_money(v["tax"])}</td>'
        f'<td class="r">{_money(v["total"])}</td></tr>'
        for k, v in sorted(d["by_cat"].items(), key=lambda kv: -kv[1]["total"]))
    cat += (f'<tr class="sum"><td class="c">合计</td><td class="c">{n}</td>'
            f'<td class="r">{_money(d["amount"])}</td><td class="r">{_money(d["tax"])}</td>'
            f'<td class="r">{_money(d["total"])}</td></tr>')
    cat_th = ("<th>费用类别</th><th style='width:64px'>单据张数</th>"
              "<th style='width:96px'>不含税金额</th><th style='width:84px'>税额</th>"
              "<th style='width:100px'>价税合计</th>")

    # ---------- 差旅费补助（用户自己填：几个人 × 几天 × 多少钱）----------
    alw_html = ""
    if d["alw"]:
        alw_body = "".join(
            f'<tr><td>{_esc(a["name"])}</td><td class="c">{_esc(a["people"])}</td>'
            f'<td class="c">{_esc(a["days"])}</td><td class="r">{_money(a["rate"])}</td>'
            f'<td class="r">{_money(a["amount"])}</td></tr>' for a in d["alw"])
        alw_body += (f'<tr class="sum"><td class="c" colspan="4">差旅费补助合计</td>'
                     f'<td class="r">{_money(d["alw_total"])}</td></tr>')
        alw_html = f"""<div class="sec">差旅费补助（{len(d['alw'])} 项）</div>
<table><thead><tr><th>补助项目</th><th style="width:60px">人数</th><th style="width:60px">天数</th>
<th style="width:108px">标准（元/人·天）</th><th style="width:100px">金额（元）</th></tr></thead>
<tbody>{alw_body}</tbody></table>"""

    # ---------- 报销金额：凭证合计 + 差旅费补助 ----------
    if d["alw_total"]:
        tail = f"""<table class="info" style="margin-top:6px"><tr>
  <th>预借款</th><td>&nbsp;</td><th>应退 / 补</th><td>&nbsp;</td>
  <th>附件张数</th><td colspan="3" class="c">{n} 张</td></tr>
  <tr><th>凭证金额</th><td class="r">{_money(d['total'])}</td>
  <th>差旅费补助</th><td class="r">{_money(d['alw_total'])}</td>
  <th>报销合计</th><td colspan="3" class="r">{_money(d['grand'])}</td></tr></table>"""
    else:
        tail = f"""<table class="info" style="margin-top:6px"><tr>
  <th>预借款</th><td>&nbsp;</td><th>应退 / 补</th><td>&nbsp;</td>
  <th>附件张数</th><td colspan="3" class="c">{n} 张</td></tr></table>"""

    return f"""<!doctype html><html><head><meta charset="utf-8"><style>{_BASE_CSS}</style></head><body>
<h1>{_esc(kind)}</h1>
<div class="sub">（本单由发票报销工具按台账自动生成，金额以所附原始凭证为准）</div>
{info}
<div class="sec">费用明细（共 {n} 张凭证）</div>
<table><thead><tr>{th}</tr></thead><tbody>{''.join(body)}
<tr class="sum"><td colspan="{len(cols) - 1}" class="r">合计（价税合计）</td>
<td class="r">{_money(d['total'])}</td></tr>
</tbody></table>
{footnote}
<div class="sec">金额汇总</div>
<table><thead><tr>{cat_th}</tr></thead><tbody>{cat}</tbody></table>
{alw_html}
<div class="money">
  <div class="l">报销金额（大写）：{_esc(d['upper'])}</div>
  <div class="r">小写：¥{_money(d['grand'])}</div>
</div>
{tail}
<div class="attach">附件清单：{_esc(d['attach'])}</div>
<div class="sign">
  <div>报销人（签字）</div><div>部门负责人</div><div>财务审核</div>
  <div>分管领导</div><div>领款人</div>
</div>
<div class="foot"><span>{_esc(meta.get('note', ''))}</span>
<span>制表时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}　共 {n} 张凭证</span></div>
</body></html>"""


def make_report_pdf(rows: list, meta: dict, out_path):
    return html_to_pdf(build_report_html(rows, meta), out_path)


# ---------------------------------------------------------------- 报销单：Excel
def build_report_xlsx(rows: list, meta: dict, out_path):
    """
    跟 PDF 同一份数据的 Excel 版报销单：抬头 → 基本信息 → 费用明细 → 金额汇总
    → 大写/小写     → 附件清单 → 签字栏。A4 竖版、按宽度缩放到一页，可直接打印。
    """
    if style_of((meta or {}).get("kind")) == "jy":
        return build_jy_xlsx(rows, meta, out_path)   # 模板二版式另开一套，见下面
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins
    from openpyxl.worksheet.properties import PageSetupProperties

    d = report_data(rows, meta or {})
    kind, n = d["kind"], d["n"]
    trip_kind = kind == "差旅费报销单"
    NCOL = 9

    wb = Workbook()
    ws = wb.active
    ws.title = "报销单"
    for i, w in enumerate([6, 12, 11, 9, 9, 26, 20, 17, 13], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    thin = Side(style="thin", color="444444")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    lft = Alignment(horizontal="left", vertical="center", wrap_text=True)
    rgt = Alignment(horizontal="right", vertical="center")
    hdr_fill = PatternFill("solid", fgColor="F0F2F5")
    sum_fill = PatternFill("solid", fgColor="FAFAFA")

    def band(r, c1, c2, border=True, fill=None, align=None, font=None):
        """给一行里的一段单元格统一套样式（合并单元格时边框才不会缺）"""
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            if border:
                cell.border = box
            if fill:
                cell.fill = fill
            if align:
                cell.alignment = align
            if font:
                cell.font = font

    def text(r, c1, c2, value, align=None, font=None, fill=None, border=True):
        if c2 > c1:
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        cell = ws.cell(row=r, column=c1, value=value)
        band(r, c1, c2, border=border, fill=fill,
             align=align or lft, font=font or Font(size=10))
        return cell

    def pair(r, pairs):
        """一行里塞若干 标签+值 对（每对 2 列）；多出来的零头列由最后一格吃掉，保证铺满整行"""
        for i, (k, v) in enumerate(pairs):
            c = 1 + i * 2
            c2 = (c + 1) if i + 1 < len(pairs) else NCOL
            text(r, c, c, k, align=ctr, font=Font(size=10, bold=True), fill=hdr_fill)
            text(r, c + 1, c2, v or "", align=ctr)

    row = 1
    text(row, 1, NCOL, kind, align=ctr, font=Font(size=16, bold=True), border=False)
    ws.row_dimensions[row].height = 30
    row += 1
    text(row, 1, NCOL, "（本单由发票报销工具按台账自动生成，金额以所附原始凭证为准）",
         align=ctr, font=Font(size=9, color="888888"), border=False)
    row += 1

    who = "出差人" if trip_kind else "报销人"
    pair(row, [(who, d["person"]), ("部门", d["dept"]),
               ("报销日期", d["date"]), ("单据张数", f"{n} 张")])
    ws.row_dimensions[row].height = 20
    row += 1
    text(row, 1, 1, "出差事由" if trip_kind else "报销事由",
         align=ctr, font=Font(size=10, bold=True), fill=hdr_fill)
    text(row, 2, NCOL, (meta or {}).get("reason", ""))
    row += 1
    if trip_kind:
        pair(row, [("出差起止", f"{(meta or {}).get('start','')} 至 {(meta or {}).get('end','')}".strip(" 至")),
                   ("出差地点", (meta or {}).get("place", "")),
                   ("出差天数", d["days"])])
        row += 1
    row += 1                                     # 空一行

    text(row, 1, NCOL, f"费用明细（共 {n} 张凭证）", align=lft,
         font=Font(size=11, bold=True), border=False)
    row += 1
    cols = _TRIP_COLS if trip_kind else _COST_COLS
    headers = [c[0] for c in cols]
    # 明细表列宽按 PDF 的百分比折算到 9 列网格上（跟 PDF 版式一一对应）
    grid = list(range(1, NCOL + 1))
    if trip_kind:
        grid = [1, 2, 3, 4, (5, 6), 7, 8, 9]
    for i, h in enumerate(headers):
        g = grid[i]
        c1 = g if isinstance(g, int) else g[0]
        c2 = g if isinstance(g, int) else g[1]
        text(row, c1, c2, h, align=ctr, font=Font(size=10, bold=True), fill=hdr_fill)
    ws.row_dimensions[row].height = 20
    head_row = row
    row += 1

    for i, r in enumerate(d["rows"], start=1):
        at = str(r.get("序号") or "") in d["attach_seqs"]
        vals = [i, *_detail_values(r, kind, attach=at)]
        for j, v in enumerate(vals):
            g = grid[j]
            c1 = g if isinstance(g, int) else g[0]
            c2 = g if isinstance(g, int) else g[1]
            if isinstance(v, str) and v and v.replace(",", "").replace(".", "").isdigit() \
                    and j == len(vals) - 1:
                cell = text(row, c1, c2, float(v.replace(",", "")), align=rgt)
                cell.number_format = "#,##0.00"
            else:
                cell = text(row, c1, c2, "" if v is None else v,
                            align=rgt if j == len(vals) - 1 else (ctr if cols[j][2] == "c" else lft))
            if at:                                   # 附件行：灰字，跟发票区分开
                cell.font = Font(size=10, color="888888")
        row += 1
    if not d["rows"]:
        text(row, 1, NCOL, "（未选择凭证）", align=ctr)
        row += 1
    text(row, 1, NCOL - 1, "合计（价税合计）", align=rgt,
         font=Font(size=10, bold=True), fill=sum_fill)
    cell = text(row, NCOL, NCOL, d["total"], align=rgt, font=Font(size=10, bold=True), fill=sum_fill)
    cell.number_format = "#,##0.00"
    row += 2

    text(row, 1, NCOL, "金额汇总", align=lft, font=Font(size=11, bold=True), border=False)
    row += 1
    text(row, 1, NCOL - 4, "费用类别", align=ctr, font=Font(size=10, bold=True), fill=hdr_fill)
    for k, c in (("单据张数", NCOL - 3), ("不含税金额", NCOL - 2),
                 ("税额", NCOL - 1), ("价税合计", NCOL)):
        text(row, c, c, k, align=ctr, font=Font(size=10, bold=True), fill=hdr_fill)
    row += 1
    for k, v in sorted(d["by_cat"].items(), key=lambda kv: -kv[1]["total"]):
        text(row, 1, NCOL - 4, k)
        text(row, NCOL - 3, NCOL - 3, v["n"], align=ctr)
        for c, val in ((NCOL - 2, v["amount"]), (NCOL - 1, v["tax"]), (NCOL, v["total"])):
            cell = text(row, c, c, val, align=rgt)
            cell.number_format = "#,##0.00"
        row += 1
    text(row, 1, NCOL - 4, "合计", align=ctr, font=Font(size=10, bold=True), fill=sum_fill)
    text(row, NCOL - 3, NCOL - 3, n, align=ctr, font=Font(size=10, bold=True), fill=sum_fill)
    for c, val in ((NCOL - 2, d["amount"]), (NCOL - 1, d["tax"]), (NCOL, d["total"])):
        cell = text(row, c, c, val, align=rgt, font=Font(size=10, bold=True), fill=sum_fill)
        cell.number_format = "#,##0.00"
    row += 2

    # ---------- 差旅费补助（用户自己填：几个人 × 几天 × 多少钱）----------
    if d["alw"]:
        text(row, 1, NCOL, f"差旅费补助（{len(d['alw'])} 项）", align=lft,
             font=Font(size=11, bold=True), border=False)
        row += 1
        for k, c1, c2 in (("补助项目", 1, NCOL - 5), ("人数", NCOL - 4, NCOL - 4),
                          ("天数", NCOL - 3, NCOL - 3),
                          ("标准（元/人·天）", NCOL - 2, NCOL - 1), ("金额（元）", NCOL, NCOL)):
            text(row, c1, c2, k, align=ctr, font=Font(size=10, bold=True), fill=hdr_fill)
        row += 1
        for a in d["alw"]:
            text(row, 1, NCOL - 5, a["name"])
            text(row, NCOL - 4, NCOL - 4, _safe(a["people"]) or "", align=ctr)
            text(row, NCOL - 3, NCOL - 3, _safe(a["days"]) or "", align=ctr)
            for c1, c2, val in ((NCOL - 2, NCOL - 1, a["rate"]), (NCOL, NCOL, a["amount"])):
                cell = text(row, c1, c2, val, align=rgt)
                cell.number_format = "#,##0.00"
            row += 1
        text(row, 1, NCOL - 1, "差旅费补助合计", align=rgt,
             font=Font(size=10, bold=True), fill=sum_fill)
        cell = text(row, NCOL, NCOL, d["alw_total"], align=rgt,
                    font=Font(size=10, bold=True), fill=sum_fill)
        cell.number_format = "#,##0.00"
        row += 2

    text(row, 1, 2, "报销金额（大写）", align=ctr, font=Font(size=10, bold=True), fill=hdr_fill)
    text(row, 3, NCOL - 1, d["upper"], align=ctr)
    cell = text(row, NCOL, NCOL, d["grand"], align=rgt, font=Font(size=11, bold=True))
    cell.number_format = '"¥"#,##0.00'
    row += 1
    if d["alw_total"]:                       # 凭证金额 + 差旅费补助 = 报销合计
        for lay, val in ((1, "凭证金额"), (4, "差旅费补助"), (7, "报销合计")):
            text(row, lay, lay, val, align=ctr, font=Font(size=10, bold=True), fill=hdr_fill)
        for val, c1, c2 in ((d["total"], 2, 3), (d["alw_total"], 5, 6), (d["grand"], 8, NCOL)):
            cell = text(row, c1, c2, val, align=rgt, font=Font(size=10, bold=True))
            cell.number_format = "#,##0.00"
        row += 1
    pair(row, [("预借款", ""), ("应退 / 补", ""), ("附件张数", f"{n} 张")])
    row += 1
    text(row, 1, NCOL, f"附件清单：{d['attach']}")
    ws.row_dimensions[row].height = 22
    row += 1
    if d["n_attach"]:
        text(row, 1, NCOL,
             f"注：标记「（附件）」的 {d['n_attach']} 张行程单是上方发票的证明附件"
             f"（金额以对应发票为准），已计在发票行内，不再重复计入合计。",
             font=Font(size=9, color="666666"), border=False)
        row += 1
    row += 1

    sign = ["报销人（签字）", "部门负责人", "财务审核", "分管领导"]
    per = max(1, NCOL // len(sign))                 # 每个签字栏占几列
    for i, label in enumerate(sign):
        c1 = 1 + i * per
        c2 = (c1 + per - 1) if i + 1 < len(sign) else NCOL   # 最后一栏吃掉零头，铺满整行
        text(row, c1, c2, label + "\n", align=Alignment(
            horizontal="left", vertical="top", wrap_text=True))
    ws.row_dimensions[row].height = 44
    row += 1
    text(row, 1, NCOL, "领款人（签字）：", align=lft, font=Font(size=10), border=False)
    ws.row_dimensions[row].height = 22
    row += 1
    text(row, 1, 4, (meta or {}).get("note", ""), border=False,
         font=Font(size=9, color="666666"))
    text(row, 5, NCOL, f"制表时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}　共 {n} 张凭证",
         align=rgt, border=False, font=Font(size=9, color="666666"))

    ws.print_area = f"A1:{get_column_letter(NCOL)}{row}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.35, right=0.35, top=0.5, bottom=0.4)
    ws.freeze_panes = f"A{head_row}"

    out_path = Path(out_path)
    try:
        wb.save(str(out_path))
    except PermissionError as e:
        raise RuntimeError(f"报销单 Excel 正被占用（多半是开着），请先关掉：{out_path}") from e
    return out_path


# ---------------------------------------------------------------- 报销单：模板二版式
# 用户 2026-09-15：「报销单还需要一个别的模板，这个设置为模板二」。
# 样子照 E:\桌\报销明细-某某公司-7.28.xlsx 的「报销单」工作表复刻：
#
#   出差报销明细                                   ← 标题（横跨整宽）
#   报销人 | 甲 | 日期 | 事由长文本            | 报销金额 | 备注
#   时间：7.23-8.19        |   | 同行人：李强、聂法银、郑凯华
#   甲 | 高铁票 | 2026-08-27 | 济南东-烟台      |  174.00 |
#   …（每张凭证一行）
#   小计                    |   |                | 2144.84 |
#                           |   |                | 出差补助 | 实际
#          出差补助 |       | 甲             |  120.00 | 1164.30
#   合计                    |   |                | 3974.84 |
#
# 两处**故意**跟参考文件不一样（都是为了让打印出来更好读，想改回去说一声）：
#   1. 「行程」列左对齐（参考文件是居中 —— 行程串很长，居中换行很难看）
#   2. 金额列右对齐（参考文件是居中 —— 钱右对齐才对得上位）
_JY_COLS = [16.2, 14.86, 13.0, 45.8, 13.66, 12.86]      # A~F 列宽，照参考文件量的
_JY_ROW_H = 35.0                                        # 参考文件的行高
_JY_CSS = """
@page { size: A4 portrait; margin: 14mm 12mm 12mm 12mm; }
* { box-sizing: border-box; }
body { font-family: "Microsoft YaHei", "SimSun", sans-serif; color: #000; font-size: 11px;
       -webkit-print-color-adjust: exact; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; }
td, th { border: 1px solid #000; height: 35pt; padding: 2px 5px; vertical-align: middle;
         text-align: center; word-break: break-all; line-height: 1.35; }
td.title { font-size: 16px; font-weight: 700; letter-spacing: 9px; height: 33pt; }
td.lab { font-weight: 700; }
td.tl { text-align: left; }
td.r { text-align: right; font-variant-numeric: tabular-nums; }
/* 分页：单行不许被拦腰截断；「小计→补助→合计」那一坨（tbody.keep）必须整块挪页；
   页脚说明别跟表格分家。加这一段的起因是 30 行台账打出来时，
   补助块被拆到两页、只剩一行「合计」孤零零在末页（2026-09-16）。 */
tr { page-break-inside: avoid; break-inside: avoid; }
tbody.keep { page-break-inside: avoid; break-inside: avoid; }
.foot { margin-top: 5px; font-size: 9px; color: #555; line-height: 1.5;
        page-break-before: avoid; break-before: avoid; }
"""


def _jy_layout(rows: list, meta: dict):
    """
    模板二版式的**行模型**：HTML（→PDF）和 Excel 两个输出都照它画，
    保证「两边长得一模一样、数字也一模一样」。
    每行是 [(值, 跨几列, 类型)]，类型只管对齐/加粗：
      title 标题 / lab 标签(加粗居中) / trip 行程(左) / money 金额(右) / "" 普通(居中)
    """
    d = report_data(rows, meta or {})
    meta = meta or {}
    lay = []

    def row(*cells):
        lay.append(list(cells))

    # 「类型」列：打车数电票本身是张「发票」，光看票面只会映射成「交通」，
    # 而它那张打车行程单写的是「打车」—— 同一笔钱两行两个叫法很难看。
    # 配上的附件是什么类型，主票就跟着写什么（见 attachment.analyze 的 link_of）。
    by_seq = {str(r.get("序号") or ""): r for r in d["rows"]}
    forced = {}
    for aseq, info in (d.get("link_of") or {}).items():
        a = by_seq.get(str(aseq))
        m = by_seq.get(str((info or {}).get("main_seq") or ""))
        if a is not None and m is not None and _ctype_of(a) == "打车行程单":
            forced[str(m.get("序号") or "")] = _short_type(a)

    row(("出差报销明细", 6, "title"))
    # 抬头行：A/C/E/F 是标签，B 填报销人、D 填事由（照参考文件的样子）
    row(("报销人", 1, "lab"), (d["person"], 1, "lab"), ("日期", 1, "lab"),
        (str(meta.get("reason") or ""), 1, "lab"), ("报销金额", 1, "lab"), ("备注", 1, "lab"))
    # 左边写时间（填了出差起止就是区间，否则写报销日期），右边写同行人
    span = f"{meta.get('start', '')} - {meta.get('end', '')}".strip(" -")
    left = f"时间：{span}" if span else f"日期：{d['date']}"
    who = "、".join(d["travelers"])
    row((left, 2, ""), ("", 1, ""), (f"同行人：{who}" if who else "", 3, ""))
    # 明细：出行人 / 类型 / 日期 / 行程 / 金额 / 备注
    for r in d["rows"]:
        at = str(r.get("序号") or "") in d["attach_seqs"]
        mark = str(r.get("票面标记") or "").strip()
        note = "　".join(x for x in (mark, "（附件）" if at else "") if x)
        trip = str(r.get("行程/明细") or "").strip()
        proj = str(r.get("项目/事由") or "").strip()
        seq = str(r.get("序号") or "")
        row((str(r.get("出行人") or "").strip(), 1, ""),
            (forced.get(seq) or _short_type(r, trip), 1, ""),
            (r.get("开票日期", ""), 1, ""),
            (trip or proj, 1, "trip"),
            (_safe(r.get("价税合计")), 1, "money"),
            (note, 1, ""))
    if not d["rows"]:
        row(("（未选择凭证）", 6, ""))
    row(("小计", 2, "lab"), ("", 1, ""), ("", 1, ""), (d["total"], 1, "money"), ("", 1, ""))
    if d["by_person"] and d["alw_total"]:
        # 出差补助：每人一行，补助 ＋ 这个人的票据合计 ＝ 实际（表头照参考文件放在 E/F）
        row(("", 1, ""), ("", 1, ""), ("", 1, ""), ("", 1, ""),
            ("出差补助", 1, "lab"), ("实际", 1, "lab"))
        for p in d["by_person"]:
            row(("", 1, ""), ("出差补助", 1, "lab"), ("", 1, ""), (p["name"], 1, "lab"),
                (p["alw"], 1, "money"), (p["real"], 1, "money"))
    row(("合计", 1, "lab"), ("", 1, ""), ("", 1, ""), ("", 1, ""),
        (d["grand"], 1, "money"), ("", 1, ""))
    return d, lay


def _jy_tail_index(d: dict) -> int:
    """
    行模型里「小计」那一行是第几行（0 基）。小计 → 补助块 → 合计 这一坨在打印时必须
    抱在一起：掉一行「合计」到下一页、或者补助块被拦腰截断，财务看着就是个半截单子。
    （行序见 _jy_layout：0 标题 / 1 抬头 / 2 时间 / 3… 明细 / 小计 …）
    """
    return 3 + (len(d["rows"]) or 1)


def _jy_foot_notes(d: dict, meta: dict) -> list:
    out = []
    if d["n_attach"]:
        out.append(f"注：标「（附件）」的 {d['n_attach']} 张行程单是上方发票的证明附件，"
                   f"金额已计在对应发票里，不再重复合计。")
    if str((meta or {}).get("note") or "").strip():
        out.append(str(meta["note"]).strip())
    out.append(f"制表时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}　共 {d['n']} 张凭证")
    return out


def _jy_col_pct() -> list:
    """_JY_COLS（Excel 的字符宽）换算成百分比，给 HTML 的 <colgroup> 用。
    ⚠️ 不给的话 table-layout:fixed 会把 6 列平摊，宽窄全乱 —— 「行程」被挤成窄条、
    文件长的一大串地址换行成 4 行，行数虚高、白白多出一页（2026-09-16 修）。"""
    tot = float(sum(_JY_COLS)) or 1.0
    return [round(w / tot * 100, 2) for w in _JY_COLS]


def build_jy_html(rows: list, meta: dict) -> str:
    """模板二版式的 PDF 版（HTML → 无头 Edge 打印）。"""
    d, lay = _jy_layout(rows, meta)
    tail = _jy_tail_index(d)          # 小计那一行起，整块不许分页拆开

    def render(cells):
        tds = []
        for v, cs, kind in cells:
            if kind == "money":
                txt, cls = _money(v), "r"
            else:
                txt = _esc(v if v is not None else "")
                cls = {"title": "title", "lab": "lab", "trip": "tl"}.get(kind, "")
            span = f' colspan="{cs}"' if cs > 1 else ""
            tds.append(f'<td class="{cls}"{span}>{txt or "&nbsp;"}</td>')
        return "<tr>" + "".join(tds) + "</tr>"

    body = "".join(render(c) for c in lay[:tail])
    keep = "".join(render(c) for c in lay[tail:])
    foot = "<br>".join(_esc(x) for x in _jy_foot_notes(d, meta or {}))
    cols = "".join(f'<col style="width:{p}%">' for p in _jy_col_pct())
    return f"""<!doctype html><html><head><meta charset="utf-8"><style>{_JY_CSS}</style></head><body>
<table><colgroup>{cols}</colgroup><tbody>{body}</tbody><tbody class="keep">{keep}</tbody></table>
<div class="foot">{foot}</div>
</body></html>"""


def build_jy_xlsx(rows: list, meta: dict, out_path):
    """模板二版式的 Excel 版：跟 PDF 同一份行模型，版式、数字完全一致。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins
    from openpyxl.worksheet.properties import PageSetupProperties

    d, lay = _jy_layout(rows, meta)
    wb = Workbook()
    ws = wb.active
    ws.title = "报销单"
    for i, w in enumerate(_JY_COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    thin = Side(style="thin", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    base = Font(name="微软雅黑", size=11)
    bold = Font(name="微软雅黑", size=11, bold=True)
    big = Font(name="微软雅黑", size=14, bold=True)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    lft = Alignment(horizontal="left", vertical="center", wrap_text=True)
    rgt = Alignment(horizontal="right", vertical="center")

    r = 1
    for cells in lay:
        c = 1
        first = str(cells[0][0] or "")
        for v, cs, kind in cells:
            if cs > 1:
                ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + cs - 1)
            ws.cell(row=r, column=c, value=(v if v != "" else None))
            for cc in range(c, c + cs):          # 合并格里每一格都要上样式，边框才不缺
                x = ws.cell(row=r, column=cc)
                x.border = box
                if kind == "title":
                    x.font, x.alignment = big, ctr
                elif kind == "lab":
                    x.font, x.alignment = bold, ctr
                elif kind == "money":
                    x.font, x.alignment = base, rgt
                    x.number_format = "#,##0.00"
                elif kind == "trip":
                    x.font, x.alignment = base, lft
                else:
                    x.font, x.alignment = base, ctr
            c += cs
        # 行高照参考文件：标题 33、抬头 60.6、小计/合计 28.05、其余 35
        if r == 1:
            ws.row_dimensions[r].height = 33.0
        elif r == 2:
            ws.row_dimensions[r].height = 60.6
        elif first in ("小计", "合计"):
            ws.row_dimensions[r].height = 28.05
        else:
            ws.row_dimensions[r].height = _JY_ROW_H
        r += 1

    # 页脚几行说明（附件提示 / 用户备注 / 制表时间）：表格下方，无边框，右对齐
    for note in _jy_foot_notes(d, meta or {}):
        cell = ws.cell(row=r, column=1, value=note)
        cell.font = Font(name="微软雅黑", size=9, color="555555")
        cell.alignment = Alignment(horizontal="right", vertical="center")
        r += 1

    last = r - 1
    ws.print_area = f"A1:{get_column_letter(len(_JY_COLS))}{last}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.7, right=0.7, top=0.75, bottom=0.75)
    ws.freeze_panes = "A3"

    out_path = Path(out_path)
    try:
        wb.save(str(out_path))
    except PermissionError as e:
        raise RuntimeError(f"报销单 Excel 正被占用（多半是开着），请先关掉：{out_path}") from e
    return out_path


def make_report(rows: list, meta: dict, base_path, fmt: str = "pdf"):
    """
    按格式生成报销单。base_path 传「不带后缀」的基名，返回生成的文件路径列表。
      fmt = 'pdf' | 'xlsx' | 'both'
    """
    base_path = Path(base_path)
    fmt = (fmt or "pdf").lower()
    outs = []
    if fmt in ("pdf", "both"):
        outs.append(make_report_pdf(rows, meta, base_path.with_suffix(".pdf")))
    if fmt in ("xlsx", "excel", "both"):
        outs.append(build_report_xlsx(rows, meta, base_path.with_suffix(".xlsx")))
    if not outs:
        raise ValueError(f"不认识的报销单格式：{fmt}")
    return outs


# ---------------------------------------------------------------- 合并发票
def _cutline_pdf() -> Path:
    """一张只有中间虚线的 A4 空白页，合并时叠上去当裁切线"""
    key = "cutline"
    if key in _CUTLINE_CACHE and Path(_CUTLINE_CACHE[key]).exists():
        return Path(_CUTLINE_CACHE[key])
    html = """<!doctype html><html><head><meta charset="utf-8"><style>
@page { size: A4 portrait; margin: 0; }
html,body { margin:0; padding:0; height:100%; }
.line { position:absolute; top:50%; left:6mm; right:6mm; border-top:1px dashed #999; }
.tag { position:absolute; top:50%; left:6mm; font:9px sans-serif; color:#999;
       background:#fff; padding:0 3px; transform:translateY(-50%); }
</style></head><body>
<div class="line"></div><div class="tag">裁切线</div>
</body></html>"""
    out = WORK_DIR / "_cutline.pdf"
    html_to_pdf(html, out)
    _CUTLINE_CACHE[key] = str(out)
    return out


def merge_pdfs(paths, out_path, mode: str = "2up", units=None):
    """
    合并发票 PDF。
      mode='2up'   一张 A4 上下两张，中间一条裁切线（用户要的样式）
      mode='plain' 顺序合并，每页一张，原尺寸

    units（可选）：分组。每个单元 = [发票, 它的行程单…]，**同一个单元的两页尽量放同一张
    A4 的上下两半**（用户要求：「行程单是证明这张发票的附件，要在一块」）。
    不传就当成「一页一个单元」，跟以前的行为一样。
    """
    from pypdf import PdfReader, PdfWriter, Transformation

    paths = [Path(p) for p in paths if p and Path(p).exists()]
    keep = {str(p) for p in paths}
    if units:
        units = [[Path(p) for p in u if str(p) in keep] for u in units]
        units = [u for u in units if u]
        # units 里没覆盖到的（理论上不会有）补在最后，宁可多打印也不少页
        inside = {str(p) for u in units for p in u}
        units += [[p] for p in paths if str(p) not in inside]
    else:
        units = [[p] for p in paths]
    paths = [p for u in units for p in u]
    if not paths:
        raise RuntimeError("没有可合并的发票文件")
    writer = PdfWriter()
    skipped = []
    slot_h = A4_H / 2                   # 一张 A4 分成上下两半，各放一张

    def place(page, src_path, k):
        """把一张 PDF 缩放后贴到 A4 的上半（k=0）或下半（k=1）"""
        try:
            src = PdfReader(str(src_path)).pages[0]
        except Exception as e:
            skipped.append((Path(src_path).name, f"{type(e).__name__}: {e}"))
            return
        sw = float(src.mediabox.width) or A4_W
        sh = float(src.mediabox.height) or A4_H
        avail_w, avail_h = A4_W - MARGIN * 2, slot_h - MARGIN * 2
        s = min(avail_w / sw, avail_h / sh)
        tx = (A4_W - sw * s) / 2
        gap = MARGIN + (avail_h - sh * s) / 2
        ty = (slot_h + gap) if k == 0 else gap
        page.merge_transformed_page(src, Transformation(ctm=(s, 0, 0, s, tx, ty)))

    if mode == "plain":
        for p in paths:
            try:
                for pg in PdfReader(str(p)).pages:
                    writer.add_page(pg)
            except Exception as e:
                skipped.append((p.name, f"{type(e).__name__}: {e}"))
    else:
        overlay = None
        try:
            overlay = PdfReader(str(_cutline_pdf())).pages[0]
        except Exception:
            overlay = None
        # 排版：一行 A4 两个位置。凑不满时，一个「发票+附件」小组不会被拆到两张纸上
        # （拆了就违背了「要在一块」的初衷），宁可少用半张纸。
        sheets, cur = [], []
        for u in units:
            if len(cur) + len(u) <= 2:
                cur.extend(u)
            else:
                if cur:
                    sheets.append(cur)
                    cur = []
                if len(u) <= 2:
                    cur = list(u)
                else:                       # 一个小组超过 2 页（1 发票 + 多张附件）只能拆
                    for p in u:
                        cur.append(p)
                        if len(cur) == 2:
                            sheets.append(cur)
                            cur = []
            if len(cur) == 2:
                sheets.append(cur)
                cur = []
        if cur:
            sheets.append(cur)
        for pair in sheets:
            page = writer.add_blank_page(width=A4_W, height=A4_H)
            for k, p in enumerate(pair):
                place(page, p, k)
            if overlay is not None:
                page.merge_transformed_page(overlay, Transformation(ctm=(1, 0, 0, 1, 0, 0)))
    out_path = Path(out_path)
    with open(out_path, "wb") as f:
        writer.write(f)
    return {"out": str(out_path), "pages": len(writer.pages),
            "files": len(paths), "skipped": skipped}


if __name__ == "__main__":
    for v in (0, 0.5, 1, 10, 100.05, 1000, 1001, 10000, 12345.67, 100000000.01, -88.8):
        print(f"{v:>14} → {rmb_upper(v)}")
