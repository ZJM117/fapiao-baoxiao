# -*- coding: utf-8 -*-
"""
附件归位：把「行程单」这类附件排到它所对应的发票后面
====================================================
用户要求（2026-09-15）：
    「打车行程单要跟在那个发票之后，然后合成 pdf 的时候也要和进去，
      因为他是证明这个发票的附件，要在一块，你可以考虑金额什么的」

背景
----
一张打车发票（凭证类型=发票）和它对应的打车行程单是**两条独立的凭证记录**：
发票是报销凭证、行程单是证明这张发票明细的附件，打印/装订时必须挨在一起。
但台账是按入库顺序一行行存下来的，两份东西经常隔得很远；合并 PDF 时也是各按
各的日期排，于是「发票 + 它的行程单」被拆到两张 A4 上，报表里的明细也分了家。

怎么认它们是一对（用户提示「可以考虑金额什么的」）
--------------------------------------------------
打车行程单上没有发票号码，但它和对应发票的「价税合计是分毫不差的」——
实测用户现场 5 组真实数据全部精确相同：
    76.10 / 13.58 / 11.40 / 65.67 / 51.65
所以拿「金额相等」当主判据（这是硬条件），再叠加几个加分项防止配错：

    基础              金额相等（差 ≤ 0.01）            +10
    同一文件夹        同一个人报的票通常放一个目录       +2
    报销人相同                                        +1
    日期接近          开票日期相差 ≤ 10 天             +1
    对方是「发票」    优先配给真正的发票而不是杂项       +1

一对一：一张行程单只挂一张发票，一张发票也只挂一张行程单。分数高的先配，
配完两边都出局，避免「一张行程单被两张同额发票抢」。

金额口径（重要）
----------------
配上的行程单**不再单独计钱** —— 它的金额已经含在对应发票里了
（用户原话：「他是证明这个发票的附件」）。不这么处理的话，一张 76.10 的打车票
会被算成 76.10 + 76.10 = 152.20，报销单就出错了。
没配上发票的行程单（比如行程单就是唯一凭证）还是照常计钱。
"""
import os
import re

FOLLOW_TYPES = ("打车行程单", "出租车行程单", "网约车行程单")   # 要跟在主票后面的附件
PRINCIPAL_TYPES = ("发票", "数电票", "增值税发票")             # 打车发票的「凭证类型」就是「发票」

_AMOUNT_TOL = 0.01          # 金额相等判定容差
_DAY_NEAR = 10              # 日期接近的宽松阈值（天）
_DATE_RE = re.compile(r"(\d{4})\D(\d{1,2})\D(\d{1,2})")


def amount_of(r) -> float:
    """一行凭证的价税合计（读不出来当 0）"""
    try:
        return round(float(str(r.get("价税合计") or 0).replace(",", "") or 0), 2)
    except (TypeError, ValueError):
        return 0.0


def _folder(r) -> str:
    """文件所在的目录（用于「同一批人放同一个文件夹」的加分）"""
    p = str(r.get("文件路径") or "").replace("\\", "/").strip()
    d = os.path.dirname(p)
    # 只取最后两级，避免「E:/桌」这种公共前缀让所有票都算同一个文件夹
    parts = [x for x in d.split("/") if x]
    return "/".join(parts[-2:]).lower()


def _days(a, b):
    """两张凭证开票日期相差多少天；任一读不出来返回 None"""
    def d(r):
        m = _DATE_RE.search(str(r.get("开票日期") or ""))
        if not m:
            return None
        try:
            return int(m.group(1)) * 10000 + int(m.group(2)) * 100 + int(m.group(3))
        except ValueError:
            return None
    x, y = d(a), d(b)
    if x is None or y is None:
        return None
    def to_ord(v):
        import datetime
        return datetime.date(v // 10000, v // 100 % 100, v % 100).toordinal()
    try:
        return abs(to_ord(x) - to_ord(y))
    except ValueError:
        return None


def row_sort_key(r):
    """
    出单据/合并时的排序键：先按开票日期（空日期的排最后），再按序号（**按数字排**）。
    ⚠️ 序号原来是按字符串排的：「10」<「2」，同一个日期里的行序会乱；
       而且没开票日期的行程单会跑到最前面，跟它的发票分家。
    """
    seq = str(r.get("序号") or "").strip()
    n = int(seq) if seq.isdigit() else 0
    date = str(r.get("开票日期") or "").strip()
    return (1 if not date else 0, date, n)


def _score(att, main) -> int:
    """两个候选配不配、有多配（金额不等直接 0）"""
    aa, ma = amount_of(att), amount_of(main)
    if aa <= 0 or ma <= 0 or abs(aa - ma) > _AMOUNT_TOL:
        return 0
    s = 10
    fa, fm = _folder(att), _folder(main)
    if fa and fa == fm:
        s += 2
    pa = str(att.get("报销人") or "").strip()
    if pa and pa == str(main.get("报销人") or "").strip():
        s += 1
    d = _days(att, main)
    if d is not None and d <= _DAY_NEAR:
        s += 1
    if str(main.get("凭证类型") or "").strip() in PRINCIPAL_TYPES:
        s += 1
    return s


def match_pairs(rows) -> list:
    """找出所有「附件 → 主票」的配对，返回 [(附件下标, 主票下标), ...]"""
    rows = list(rows or [])
    atts = [(i, r) for i, r in enumerate(rows)
            if str(r.get("凭证类型") or "").strip() in FOLLOW_TYPES]
    if not atts:
        return []
    cand = []
    for ai, a in atts:
        for mi, m in enumerate(rows):
            if mi == ai or str(m.get("凭证类型") or "").strip() in FOLLOW_TYPES:
                continue
            s = _score(a, m)
            if s:
                cand.append((s, ai, mi))
    # 分高的先配；同分时按先后顺序（稳定、可复现）
    cand.sort(key=lambda x: (-x[0], x[1], x[2]))
    used_a, used_m, pairs = set(), set(), []
    for _s, ai, mi in cand:
        if ai in used_a or mi in used_m:
            continue
        used_a.add(ai)
        used_m.add(mi)
        pairs.append((ai, mi))
    pairs.sort()                     # 按附件出现的先后给出，方便阅读日志
    return pairs


def order_rows(rows, pairs=None) -> list:
    """把附件那一行挪到它主票的**紧后面**；没配上主票的留在原地"""
    rows = list(rows or [])
    pairs = match_pairs(rows) if pairs is None else pairs
    follower = {mi: ai for ai, mi in pairs}      # 主票下标 → 附件下标
    moved = set(follower) | {ai for ai, _ in pairs}
    out = []
    for i, r in enumerate(rows):
        if i in follower:
            out.append(r)
            out.append(rows[follower[i]])
        elif i in moved:
            continue                              # 附件已跟着主票出去了
        else:
            out.append(r)
    return out


def analyze(rows) -> dict:
    """
    一次把该算的都算好（报表、合并、界面标记都用这一份结果）：
        rows        附件已跟到主票后面的行序
        pairs       [(附件行, 主票行, 金额), ...]
        attach_seqs 配上的附件序号集合（str），界面/报表标「附件」用
        link_of     {附件序号: 主票序号}
        n_attach    配上的组数
    """
    rows = list(rows or [])
    pairs = match_pairs(rows)
    ordered = order_rows(rows, pairs)
    detail = []
    attach_seqs, link_of = set(), {}
    for ai, mi in pairs:
        a, m = rows[ai], rows[mi]
        amount = amount_of(a)
        detail.append((a, m, amount))
        aseq = str(a.get("序号") or "")
        attach_seqs.add(aseq)
        link_of[aseq] = {"main_seq": str(m.get("序号") or ""),
                         "main_amount": amount_of(m),
                         "main_desc": (str(m.get("项目/事由") or "").strip()
                                       or str(m.get("销方名称") or "").strip()
                                       or str(m.get("发票号码") or "").strip())}
    return {"rows": ordered, "pairs": detail, "attach_seqs": attach_seqs,
            "link_of": link_of, "n_attach": len(detail)}


def link_map(rows) -> dict:
    """{附件序号: {main_seq, main_amount, main_desc}} —— 界面标「↳ 附件」用"""
    return analyze(rows)["link_of"]


def describe(pairs) -> str:
    """给日志用的一句话（如「76.10元 → 序号11」）"""
    if not pairs:
        return ""
    return "；".join(f"{amt:,.2f} 元 → 序号 {a.get('序号')}（{a.get('凭证类型')}）"
                     for a, _m, amt in pairs)


def all_pair_rows(rows):
    """返回 (排序后的行, 配对明细, 附件序号集合)，给顺序敏感的地方省一次计算"""
    d = analyze(rows)
    return d["rows"], d["pairs"], d["attach_seqs"]
