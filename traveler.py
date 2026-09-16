# -*- coding: utf-8 -*-
"""
出行人认领（三层兜底 + 自动学习）

背景（2026-09-15 用户提的）：
  「有的行程单没有人名，就是打车的单子」——打车行程单票面**只有「行程人手机号」**，
  没有姓名；而打车数电票有的填了「出行人」有的没填。要判断一张票是谁坐的，
  光靠票面不够。

三层兜底，可靠度从高到低：
  1) 票面 / 文件名 —— 火车票的乘车人、机票的旅客姓名、打车票的出行人列、以及
     用户自己给文件起的名字（`济南东-烟台 李强 10.pdf`）。读到就用。
  2) 手机号对照表 —— 行程单只有手机号，用「手机号 → 姓名」对照表翻成人名。
     对照表存在 缓存/手机号姓名.json，可以自动学（见 learn()），也可以人工改。
  3) 文件夹名 —— 报销时本来就按人分文件夹（`王小明（某某公司）`），
     票面读不到就用它兜底。

三层都没认出来 → **留空**，交给用户在界面上手填（用户明确要求保留人工参与）。

自动学习 learn() 的规则（保守，宁可学不到也不学错）：
  同一文件夹里，票面确认过姓名、且日期与这张行程单相差 ≤ window_days 天的记录，
  如果**只有一个候选姓名**，就把该手机号记到他名下；有多个候选就不学。
"""
from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path

import app_paths

# 不像人名的词：单位名、票据用语、常见地址词都会命中，避免把「某某科技」当成出行人
_NOT_NAME = (
    "公司", "中心", "事务所", "银行", "税务", "发票", "报销", "凭证", "票据", "整理",
    "酒店", "宾馆", "超市", "医院", "学校", "大学", "集团", "企业", "服务", "科技",
    "出行", "行程", "旅客", "旅游", "交通", "客运", "出租车", "网约车", "出租",
    "备注", "名称", "账号", "开户", "项目", "费用", "平台", "图片", "扫描", "待报",
)
# 文件夹名尾部常见的业务后缀，先削掉再认人
_FOLDER_SUFFIX = re.compile(
    r"(报销单|报销单子|报销|发票|票据|行程单|行程|出差|凭证|整理|待报|已报|资料|文件|扫描件?|图片)$")

# 复姓。4 个字的人名只在「复姓 + 双字名」（欧阳娜娜）时才认 ——
# 否则「某某科技」这种公司简称（某某科技公司）会被当成出行人。
_COMPOUND_SURNAMES = (
    "欧阳", "上官", "司马", "东方", "独孤", "南宫", "夏侯", "诸葛", "尉迟", "皇甫",
    "公孙", "慕容", "宇文", "司徒", "长孙", "钟离", "令狐", "西门", "百里", "呼延",
    "轩辕", "拓跋", "第五", "端木", "澹台", "赫连", "万俟", "太史", "闻人", "宗政",
    "濮阳", "公冶", "太叔", "申屠", "仲孙", "鲜于", "司空", "闾丘", "公良", "漆雕",
    "乐正", "宰父", "谷梁", "夹谷", "段干", "东郭", "南门", "微生", "梁丘",
)

# 来源标签（界面和日志直接用这几个词）
SRC_FACE = "票面"
SRC_FILENAME = "文件名"
SRC_PHONE = "手机号"
SRC_FOLDER = "文件夹"
SRC_MANUAL = "手填"
CONFIDENT = (SRC_FACE, SRC_FILENAME)        # 可以拿来"教"手机号的来源


def is_name(s) -> bool:
    """
    这两三个字像不像一个人的姓名。规则：
      · 全中文 2~4 字；
      · 不含单位 / 票据 / 地址类用语（公司、中心、行程、报销…）；
      · 4 字的必须是**复姓 + 双字名**，否则很容易把公司简称（「某某科技」）认成人。
    """
    t = str(s or "").strip()
    if not re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", t):
        return False
    if any(w in t for w in _NOT_NAME):
        return False
    if len(t) == 4 and t[:2] not in _COMPOUND_SURNAMES:
        return False
    return True


def folder_of(path) -> str:
    """文件所在的文件夹名（用来认人和做「同一个文件夹」的判断）"""
    p = str(path or "").strip()
    if not p:
        return ""
    try:
        return Path(p).parent.name
    except Exception:
        return ""


def from_folder(path) -> str:
    """
    从文件夹名里认人：`王小明（某某公司）`→王小明；`李强（某某公司）`→李强。
    认不出（比如叫「报销」「新建文件夹」）就返回空串。
    """
    name = folder_of(path)
    if not name:
        return ""
    core = re.split(r"[（(]", name)[0]
    core = re.sub(r"[\s\u3000]+", "", core)
    core = re.sub(r"[\d\W_]+", "", core, flags=re.UNICODE)
    core = _FOLDER_SUFFIX.sub("", core)
    return core if is_name(core) else ""


def _d(s):
    """把 2026-08-19 / 2026年08月19日 变成 date"""
    t = str(s or "")
    m = re.search(r"(\d{4})\D{0,2}(\d{1,2})\D{0,2}(\d{1,2})", t)
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except Exception:
        return None


def row_date(rec) -> date | None:
    return _d(rec.get("date")) or _d(rec.get("trip_date"))


# ---------------------------------------------------------------- 手机号对照表
def map_path() -> Path:
    """对照表放缓存目录（不进同步盘的代码目录，方便随时删）"""
    return Path(app_paths.WORK_DIR) / "手机号姓名.json"


def load_map() -> dict:
    try:
        d = json.loads(map_path().read_text(encoding="utf-8"))
        return {str(k): str(v) for k, v in d.items()} if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_map(d: dict) -> None:
    p = map_path()
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def set_name(phone, name) -> dict:
    """人工指定：这个手机号是谁。名字给空串 = 删掉这条。"""
    ph = re.sub(r"\D", "", str(phone or ""))
    if not ph:
        return load_map()
    d = load_map()
    nm = str(name or "").strip()
    if nm:
        d[ph] = nm
    else:
        d.pop(ph, None)
    save_map(d)
    return d


# ---------------------------------------------------------------- 三层认人
def resolve(rec, path=None, phone_map=None, folder_name="") -> tuple:
    """
    认出行人，返回 (姓名, 来源)。来源 ∈ 票面 / 文件名 / 手机号 / 文件夹 / ""（没认出来）。
    """
    p = path or rec.get("src") or ""
    tv = str(rec.get("traveler") or "").strip()
    if tv and is_name(tv):
        return tv, (rec.get("tv_from") or SRC_FACE)

    pm = load_map() if phone_map is None else phone_map
    ph = re.sub(r"\D", "", str(rec.get("phone") or ""))
    if ph and is_name(pm.get(ph, "")):
        return pm[ph], SRC_PHONE

    f = folder_name or from_folder(p)
    if f:
        return f, SRC_FOLDER
    return "", ""


def apply(recs: list, phone_map=None) -> dict:
    """
    把认好的出行人写回每个 rec（traveler / tv_from），并统计各来源多少张。
    返回 {"票面":n, "文件名":n, "手机号":n, "文件夹":n, "":n, "learned":{phone:name}}
    """
    pm = load_map() if phone_map is None else dict(phone_map)
    learned = _learn(recs, pm)
    pm.update(learned)
    stat = {SRC_FACE: 0, SRC_FILENAME: 0, SRC_PHONE: 0, SRC_FOLDER: 0, "": 0}
    for rec in recs or []:
        name, src = resolve(rec, phone_map=pm)
        rec["traveler"] = name
        rec["tv_from"] = src
        rec["tv_src_word"] = src
        stat[src] = stat.get(src, 0) + 1
    stat["learned"] = learned
    if learned:
        save_map(pm)
    return stat


def _learn(recs: list, pm: dict, window_days: int = 7) -> dict:
    """
    学「手机号 → 姓名」。只学**同文件夹 + 日期接近 + 唯一候选**的，学不准就不学。
    """
    known = []                     # (文件夹, 姓名, 日期)
    for rec in recs or []:
        nm = str(rec.get("traveler") or "").strip()
        src = rec.get("tv_from") or SRC_FACE
        if not (nm and is_name(nm) and src in CONFIDENT):
            continue
        known.append((folder_of(rec.get("src")), nm, row_date(rec)))
    out = {}
    for rec in recs or []:
        ph = re.sub(r"\D", "", str(rec.get("phone") or ""))
        if not ph or ph in pm:
            continue                   # 没手机号 / 已经记过了，不重复学也不改已有的
        folder = folder_of(rec.get("src"))
        if not folder:
            continue                   # 文件散在根目录，没有「同一批人」的范围，不猜
        d0 = row_date(rec)
        cands = set()
        for f, nm, d1 in known:
            if f != folder:
                continue
            if d0 and d1 and abs((d1 - d0).days) > window_days:
                continue
            cands.add(nm)
        if len(cands) == 1:
            out[ph] = cands.pop()
    return out


# ---------------------------------------------------------------- 人工核对记录
# 「三层兜底 + 人工核对」里的后半截：
#   自动认人有三层兜底，但**手机号译出来的、文件夹名推出来的都可能错**，
#   票面读不到的更是空的。所以要有一个人工过一遍的环节 —— 核过哪几张、
#   当时确认的是谁，记在这里；台账里的出行人跟记录不一致（说明后来又被改了），
#   核对标记自动失效，下次还会再提醒。
_CHECK_FILE = "出行人核对.json"


def _norm(p) -> str:
    """核对记录的键：文件路径规范化（Windows 路径不区分大小写）"""
    return str(Path(str(p or ""))).strip().lower()


def checked_path() -> Path:
    return Path(app_paths.WORK_DIR) / _CHECK_FILE


def load_checked() -> dict:
    """已经人工核对过的：{文件路径: 核对时确认的姓名}"""
    try:
        d = json.loads(checked_path().read_text(encoding="utf-8"))
    except Exception:
        return {}
    return {str(k): str(v) for k, v in d.items()} if isinstance(d, dict) else {}


def save_checked(d: dict) -> None:
    p = checked_path()
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def mark_checked(pairs: dict) -> dict:
    """记「这几张人工看过了」。pairs = {文件路径: 姓名}；姓名给空串 = 撤掉核对标记"""
    d = load_checked()
    for p, nm in (pairs or {}).items():
        k = _norm(p)
        if not k:
            continue
        nm = str(nm or "").strip()
        if nm:
            d[k] = nm
        else:
            d.pop(k, None)
    save_checked(d)
    return d


def is_checked(path, name, checked: dict = None) -> bool:
    """
    这张的出行人，人工核对过没有。
    口径：核对时记下的姓名 == 现在台账里的姓名（名字被改过就重新算「没核对」）。
    """
    nm = str(name or "").strip()
    if not nm:
        return False
    d = load_checked() if checked is None else checked
    return str(d.get(_norm(path)) or "") == nm


def summary_text(stat: dict) -> str:
    """给日志用的一句话"""
    parts = [f"票面 {stat.get(SRC_FACE, 0)}",
             f"文件名 {stat.get(SRC_FILENAME, 0)}",
             f"手机号 {stat.get(SRC_PHONE, 0)}",
             f"文件夹 {stat.get(SRC_FOLDER, 0)}"]
    blank = stat.get("", 0)
    s = "出行人：" + " / ".join(parts)
    if blank:
        s += f"；{blank} 张没认出来（可在台账里手填）"
    return s
