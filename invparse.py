# -*- coding: utf-8 -*-
"""
票据要素解析
============
从各种形态的票据文件里抽出登记台账要用的字段。

支持的形态（按可靠性从高到低）：
  1. `.xml`   数电票 XML 法定原件（字段最全、最准）
  2. `.zip`   里面装着 XML（税局「XML下载」下发的是 zip）
  3. `.ofd`   OFD 版式文件，内部藏三种数据，依次尝试：
                · 铁路票 → `Doc_0/Attachs/rai_issuer_*.xml`（XBRL，36 个字段，最全）
                · 数电票 → `Tags/CustomTag.xml` 给的是「字段名 → TextObject ID」映射，
                            再去 `Pages/*/Content.xml` 按 ID 把文字取回来
                · 兜底   → `OFD.xml` 里的 `CustomDatas`（号码/金额/税额/开票日期）
              三条都空时，找同目录同名的 `.pdf` 来读（OFD 旁边常常配着一份 PDF）
  4. `.pdf`   版式 PDF 文本层，按票面特征分派到五个解析器：
                数电票 / 铁路电子客票 / 航空运输电子客票行程单 / 打车行程单 / 登机牌
  5. 图片     `.jpg/.png` 等，读不出内容，登记成「图片凭证」等人工补
  6. 文件名   上面都失败时，只从文件名抠号码 / 金额 / 日期 / 起终点

统一输出（dict）：
  —— 发票要素 ——
  no 发票号码 / date 开票日期 / kind 票面名称 / ctype 凭证类型 / is_red 红字
  stamp 票面标记（退票 / 差额退票 / 改签…，来自票面红字或正文）/ red 票面红字原文
  buyer buyer_id / seller seller_id / amount tax total total_cn
  project project_code item drawer tax_bureau
  —— 行程要素 ——（发票上往往没有，行程单上才有，这正是要单独存一列的原因）
  detail 明细一行 / route 起终点 / trip_date 行程日期 / trip_time 时间
  traveler 出行人 / train_no 车次 / flight_no 航班 / seat 车厢座位
  trip_class 席别舱位 / mileage 里程 / carrier 承运方
  —— 元信息 ——
  src 文件路径 / src_kind 数据来自哪 / warn 需要人工留意的点
"""
import hashlib
import json
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

# ---------------------------------------------------------------- 常量
DOC_EXTS = (".pdf", ".ofd", ".xml", ".zip")
IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff")
SUPPORTED_EXTS = DOC_EXTS + IMAGE_EXTS

INV_NO_RE = re.compile(r"(?<!\d)\d{20}(?!\d)")
NUM_RE = re.compile(r"-?[\d,]+\.\d{2}")
# 判断「这一行是中文大写金额」的特征字。
# 只认 拾/佰/仟/圆/元/角/分 这类货币字 —— 千万别把「万」「壹」单独算进去：
# 购买方地址里的「华置万象天地」含一个「万」字，曾经把价税合计算成 0。
CN_NUM_RE = re.compile(r"[拾佰仟圆元角分]")

# 凭证类型
T_INVOICE = "发票"
T_TRAIN = "火车票"
T_FLIGHT = "飞机行程单"
T_TRIP = "打车行程单"
T_BOARD = "登机牌"
T_IMAGE = "图片凭证"
T_OTHER = "其他凭证"


# ---------------------------------------------------------------- 小工具
def _f(s):
    """'1,314.00' / '-123.96' → float；取不到返回 None"""
    if s is None:
        return None
    s = str(s).replace(",", "").replace("¥", "").replace("￥", "").strip()
    m = re.match(r"^-?\d+(\.\d+)?$", s)
    return round(float(s), 2) if m else None


def _clean(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


def _valid_ymd(y, m, d):
    """挡住从 20 位发票号码里抠出来的假日期（如 2634-91-19）"""
    try:
        y, m, d = int(y), int(m), int(d)
    except Exception:
        return False
    return 2000 <= y <= 2099 and 1 <= m <= 12 and 1 <= d <= 31


def norm_date(s: str) -> str:
    """各种日期写法 → YYYY-MM-DD（非法日期返回空串）"""
    s = (s or "").strip()
    m = re.search(r"(\d{4})\s*[-/年.]\s*(\d{1,2})\s*[-/月.]\s*(\d{1,2})", s)
    if m and _valid_ymd(*m.groups()):
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r"(?<!\d)(\d{4})(\d{2})(\d{2})(?!\d)", s)
    if m and _valid_ymd(*m.groups()):
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return ""


def _strip_station(s):
    """「随州南站」→「随州南」：跟手写台账里的叫法保持一致"""
    return re.sub(r"站$", "", _clean(s))


def _route_from_name(stem: str) -> str:
    """
    从文件名里抠区间，专治行程单 —— 高德给行程单起的名字就是「起点-终点-金额元-N个行程」：
      【合肥南-财智中心-13.58元-1个行程】高德打车电子行程单.pdf  →  合肥南-财智中心
      【某某科技-烟台站-65.67元-1个行程】高德打车电子发票.pdf      →  某某科技-烟台站
      某某科技公司-烟台站行程报销单.pdf                       →  某某科技公司-烟台站
      武汉-烟台  李强电子行程单748元.pdf                           →  武汉-烟台
    """
    s = _clean(stem)
    m = re.search(r"【([^】]{2,60})】", s)
    if m:
        s = m.group(1)
    s = re.sub(r"^(?:\d+[-_.]\s*)+", "", s)
    m = re.match(r"^([\u4e00-\u9fa5]{2,12})\s*[-–—至]\s*([\u4e00-\u9fa5]{2,12})", s)
    if not m or m.group(1) == m.group(2):
        return ""
    dep = m.group(1)
    arr = re.sub(r"(的?(?:行程报销单|电子行程单|行程单|电子发票|发票|行程))$", "", m.group(2))
    return f"{dep}-{arr}" if arr and arr != dep else ""


def _amount_from_name(stem: str):
    """文件名里的金额：「…-13.58元-1个行程】」「…行程单748元」"""
    m = re.search(r"(\d{1,6}(?:\.\d{1,2})?)\s*元", stem)
    if m:
        return _f(m.group(1))
    m = re.search(r"[-–—\s](\d{1,5}\.\d{2})", stem)
    return _f(m.group(1)) if m else None


_CN_DIGIT = {"零": 0, "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5,
             "陆": 6, "柒": 7, "捌": 8, "玖": 9}
_CN_UNIT = {"拾": 10, "佰": 100, "仟": 1000, "万": 10000, "亿": 100000000}


def cn_to_amount(s):
    """
    中文大写金额 → 数字。有的票面把「价税合计」写成「陆拾伍圆陆角柒分」单独一行，
    小写金额在别的行（甚至写成「65.67¥」，¥ 跑到数字后面），这时大写是唯一可靠的来源。

    陆拾伍圆陆角柒分            → 65.67
    壹万玖仟伍佰陆拾陆圆贰角玖分  → 19566.29
    壹拾元整                    → 10.00
    """
    s = re.sub(r"[^零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整正]", "", _clean(s))
    if not s:
        return None
    # 必须带「圆/元」或「角/分」才算金额。没有这些字的串多半只是句子里
    # 恰巧混进个「万」字（「华置万象天地」），硬算会得到 0。
    if not re.search(r"[圆元角分]", s):
        return None
    m = re.match(r"^(.*?)[圆元](.*)$", s)
    head, tail = (m.group(1), m.group(2)) if m else (s, "")

    def _int(seg):
        val = num = 0
        for ch in seg:
            if ch in _CN_DIGIT:
                num = _CN_DIGIT[ch]
            elif ch in ("拾", "佰", "仟"):
                val += (num or 1) * _CN_UNIT[ch]
                num = 0
            elif ch in ("万", "亿"):
                val = (val + num) * _CN_UNIT[ch]
                num = 0
        return val + num

    total = float(_int(head))
    m2 = re.search(r"([零壹贰叁肆伍陆柒捌玖])角", tail)
    if m2:
        total += _CN_DIGIT[m2.group(1)] * 0.1
    m3 = re.search(r"([零壹贰叁肆伍陆柒捌玖])分", tail)
    if m3:
        total += _CN_DIGIT[m3.group(1)] * 0.01
    return round(total, 2) if (head or m2 or m3) else None


def split_project(remark: str) -> tuple:
    """'<项目全名> <项目编号>' → (全名, 编号)；编号形如 1815P825000P

    PDF 里名称很长时会折行，抽出来是 `…治理工程1815L5240041` 这种「紧贴着」的形态，
    所以除了按空白切，还要认「中文/括号后面直接跟 8~20 位编号」。
    """
    remark = _clean(remark)
    m = re.search(r"^(.*?)[\s\u3000]+([0-9A-Z]{8,20})$", remark)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m = re.search(r"^(.*[\u4e00-\u9fa5）)】])[ \u3000]*([0-9A-Z]{8,20})$", remark)
    if m and re.search(r"\d", m.group(2)):
        return m.group(1).strip(), m.group(2).strip()
    return remark, ""


def _blank(src="", src_kind="", ctype=""):
    """统一空记录：所有解析器都从这里长出来，字段不会缺"""
    return {
        # 发票要素
        "no": "", "date": "", "kind": "", "ctype": ctype, "is_red": False,
        "stamp": "", "red": "",
        "buyer": "", "buyer_id": "", "seller": "", "seller_id": "",
        "amount": None, "tax": None, "total": None, "total_cn": "",
        "project": "", "project_code": "", "item": "", "drawer": "",
        "tax_bureau": "",
        # 行程要素
        "detail": "", "route": "", "trip_date": "", "trip_time": "",
        # traveler 出行人姓名；tv_from 这个名字是从哪儿来的（票面/文件名/手机号/文件夹…），
        # 界面要能把「票面读到的」和「推测出来的」区分开
        "traveler": "", "tv_from": "", "phone": "",
        "train_no": "", "flight_no": "", "seat": "",
        "trip_class": "", "mileage": None, "carrier": "",
        # 元信息
        "src": src, "src_kind": src_kind, "warn": [],
    }


# ---------------------------------------------------------------- XML（数电票法定原件）
def _xml_to_dict(data: bytes) -> dict:
    """把 XML 里所有叶子标签拍平成 {标签: 文本}（同名只取第一个）"""
    root = ET.fromstring(data)
    g = {}
    for el in root.iter():
        if el.text and el.text.strip():
            g.setdefault(el.tag, el.text.strip())
    # 嵌套标签里再挖一层（如 InherentLabel > EInvoiceType > LabelCode）
    for el in root.iter():
        for c in el:
            if c.text and c.text.strip():
                g.setdefault(f"{el.tag}/{c.tag}", c.text.strip())
    return g


def _xbrl_to_dict(data: bytes) -> dict:
    """XBRL（铁路电子客票用）拍平：标签带 {命名空间}，只留本地名"""
    root = ET.fromstring(data)
    g = {}
    for el in root.iter():
        name = el.tag.split("}")[-1] if isinstance(el.tag, str) else ""
        if not name or name in ("xbrl", "context", "unit", "schemaRef"):
            continue
        if el.text and el.text.strip():
            g.setdefault(name, el.text.strip())
    return g


def parse_xml_bytes(data: bytes, src: Path = None, src_kind: str = "xml") -> dict:
    g = _xml_to_dict(data)
    no = g.get("EIid") or g.get("InvoiceNumber") or g.get("TaxSupervisionInfo/InvoiceNumber") or ""
    date = norm_date(g.get("IssueTime") or g.get("TaxSupervisionInfo/IssueTime")
                     or g.get("RequestTime") or "")
    etype = g.get("EInvoiceType/LabelName") or ""          # 电子发票
    vat = g.get("GeneralOrSpecialVAT/LabelName") or ""     # 增值税专用发票 / 普通发票
    kind = f"{etype}（{vat}）" if etype and vat else (etype or vat)
    blue = g.get("InIssuType/LabelCode") or ""             # Y 蓝字 / N 红字
    total = _f(g.get("TotalTax-includedAmount"))
    is_red = (blue == "N") or (total is not None and total < 0)
    remark = g.get("Remark") or ""
    project, pcode = split_project(remark)
    rec = _blank(str(src) if src else "", src_kind, T_INVOICE)
    rec.update({
        "no": no,
        "date": date,
        "kind": kind or "电子发票",
        "is_red": bool(is_red),
        "buyer": _clean(g.get("BuyerName")),
        "buyer_id": _clean(g.get("BuyerIdNum")),
        "seller": _clean(g.get("SellerName")),
        "seller_id": _clean(g.get("SellerIdNum")),
        "amount": _f(g.get("TotalAmWithoutTax")) if g.get("TotalAmWithoutTax") is not None else _f(g.get("Amount")),
        "tax": _f(g.get("TotalTaxAm")) if g.get("TotalTaxAm") is not None else _f(g.get("ComTaxAm")),
        "total": total,
        "total_cn": _clean(g.get("TotalTax-includedAmountInChinese")),
        "project": project,
        "project_code": pcode,
        "item": _clean(g.get("ItemName")),
        "drawer": _clean(g.get("Drawer")),
        "tax_bureau": _clean(g.get("TaxBureauName")),
    })
    # 出行人明细（增值税发票的旅客运输服务版式会带）
    for el in ET.fromstring(data).iter():
        name = el.tag.split("}")[-1] if isinstance(el.tag, str) else ""
        if name in ("TravelerName", "PassengerName") and el.text:
            rec["traveler"] = _clean(el.text)
    if remark and not rec["detail"]:
        rec["detail"] = _trip_from_remark(remark)
    return rec


# ---------------------------------------------------------------- PDF 文本
def _pdf_text(path: Path) -> str:
    """
    取 PDF 文本层。

    注意 pypdf 会在某些字体（DengXian 这类同时挂了 FontFile2 和 FontFile3 的）上
    直接抛 `PdfReadError: More than one /FontFile found` —— 铁路电子客票 PDF 就长这样，
    所以必须保留 pdfminer 这条兜底链路，实测它能完整读出车次/席别/票价。
    """
    try:
        from pypdf import PdfReader
        r = PdfReader(str(path))
        t = "\n".join((pg.extract_text() or "") for pg in r.pages)
        if t.strip():
            return t
    except Exception:
        pass
    try:
        from pdfminer.high_level import extract_text
        return extract_text(str(path)) or ""
    except Exception:
        return ""


def _despace(s: str) -> str:
    """
    去掉行内所有空白（含全角空格）。

    为什么必须这么做：数电票 PDF 的字体是逐个字形定位的，pypdf 抽出来的文本会在
    数字之间塞空格 —— 实际见过 `2 6 3 7 2 0 0 0 0 0 0 3 8 7 4 2 9 7 9 8 1`、
    `¥1 9 82.1 7`、`2 0 2 6 年0 8 月0 7 日`、`国 网 湖 北 省 电 力 有限 公司`。
    发票上中文和数字本来就不靠空格分词（分行信息 already 在新行里），
    所以整行去空格是安全且必要的一步；不去就一个字都匹配不上。
    另一种情况是「浏览器打印成 PDF」的票（字体编码坏了），数字之间的空格是
    零宽字符，普通 \\s 吃不掉，所以把 \\u200b/\\u200c/\\u200d/\\ufeff 一并算上。
    再有一种更狠的：字符之间直接夹着 NUL（\\x00），抽出来是
    `\\x002\\x006\\x003\\x007…` —— 肉眼看像一串号码，正则却一个字符都匹配不上。
    """
    return re.sub(r"[\s\u3000\u00a0\u200b\u200c\u200d\ufeff\x00]+", "", s or "")


def _char_color(ch):
    """取 LTChar 的填充色（拿不到就返回 None）"""
    gs = getattr(ch, "graphicstate", None)
    if gs is None:
        return None
    for attr in ("ncolor", "scolor"):
        c = getattr(gs, attr, None)
        if c is None:
            continue
        try:
            vals = [float(x) for x in c]
        except Exception:
            continue
        if len(vals) in (3, 4):
            return vals
    return None


def _is_red_color(vals) -> bool:
    """够红就算红 —— 实测铁路票的红字是纯 (1,0,0)，但也留点余量给暗红/CMYK"""
    if not vals:
        return False
    if len(vals) == 4:                                    # CMYK
        c, m, y, k = vals
        return c < 0.30 and m > 0.50 and y > 0.50 and k < 0.40
    r, g, b = vals
    return r > 0.45 and (r - max(g, b)) > 0.22 and g < 0.50 and b < 0.50


_RED_CACHE = {}


def pdf_red_text(path: Path) -> str:
    """
    把 PDF 里**红色**的文字单独拎出来。

    为什么需要它：铁路电子客票把「退票」「差额退票」「始发改签」这类票面提示打成红字
    （正文黑字里也有一份，跟在「电子客票号」后面，但有的版式只在红字里写），
    所以两条都读、互为补充。

    实现：pdfminer 的 LTChar 上挂着 graphicstate（当前填充色），逐字符过一遍按颜色筛。
    ⚠️ 不能走 pypdf —— 铁路票的 DengXian 字体同时挂了 FontFile2 / FontFile3，
    pypdf 直接抛 `More than one /FontFile found`，只有 pdfminer 这条链能读。
    """
    key = str(path)
    try:
        key = f"{path}|{Path(path).stat().st_mtime_ns}"
    except Exception:
        pass
    if key in _RED_CACHE:
        return _RED_CACHE[key]
    out = []
    try:
        from pdfminer.high_level import extract_pages
        from pdfminer.layout import LTChar, LTTextContainer
        for page in extract_pages(str(path)):
            for el in page:
                if not isinstance(el, LTTextContainer):
                    continue
                for line in el:
                    for ch in line:
                        if isinstance(ch, LTChar) and _is_red_color(_char_color(ch)):
                            out.append(ch.get_text())
    except Exception:
        out = []
    txt = _clean("".join(out))
    _RED_CACHE[key] = txt
    return txt


_STAMP_WORDS = ("退票", "改签", "变更到站", "作废", "红冲")


def _stamp_from_text(flat: str, red: str = "") -> str:
    """
    认票面标记（退票 / 改签…）。红字优先，正文里的词兜底。
    顺序不能乱：先判「差额退票」，否则会被「退票」先吃掉，细节就丢了。
    """
    text = f"{flat} {red}"
    if "差额退票" in text:
        return "差额退票"
    if "退票" in text:
        return "退票"
    if "改签" in text:                 # 「始发改签」「改签差价」都落这里
        return "改签"
    if "变更到站" in text:
        return "变更到站"
    return ""


def _sniff(lines, raw) -> str:
    """按票面特征判断是哪种版式"""
    flat = "".join(lines)
    if "铁路电子客票" in flat or "买票请到12306" in flat:
        return "train"
    if "航空运输电子客票行程单" in flat:
        return "flight"
    if "行程单" in flat and "序号" in flat and ("服务商" in flat or "车型" in flat):
        return "trip"
    if ("登机牌" in flat or "BOARDING" in raw) and ("航班" in flat or "FLIGHT" in raw):
        return "boarding"
    return "einvoice"


# ---------------------------------------------------------------- PDF：铁路电子客票
def parse_train_pdf(lines, raw, path: Path) -> dict:
    rec = _blank(str(path), "pdf", T_TRAIN)
    rec["kind"] = "电子发票（铁路电子客票）"
    flat = "".join(lines)

    m = re.search(r"发票号码[：:]\s*(\d{20})", flat)
    if m:
        rec["no"] = m.group(1)
    m = re.search(r"开票日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)", flat)
    if m:
        rec["date"] = norm_date(m.group(1))
    # 乘车日期/时间：「2026年07月06日 08:00开」
    m = re.search(r"(\d{4}年\d{1,2}月\d{1,2}日)\s*(\d{1,2}:\d{2})开", flat)
    if m:
        rec["trip_date"] = norm_date(m.group(1))
        rec["trip_time"] = m.group(2)

    # 车站：票面里独立的「XX站」行，先出发、后到达
    stations = [ln for ln in lines if re.fullmatch(r"[\u4e00-\u9fa5]{2,8}站", ln)]
    dep = _strip_station(stations[0]) if stations else ""
    arr = _strip_station(stations[1]) if len(stations) > 1 else ""
    rec["route"] = f"{dep}-{arr}" if dep and arr else (dep or arr)

    # 车次：字母+数字，单独成行
    m = next((ln for ln in lines if re.fullmatch(r"[GDCZTKYLS]\d{1,4}", ln)), "")
    if m:
        rec["train_no"] = m
    # 车厢座位：「05车11F号」
    m = re.search(r"(\d{1,2}车[\dA-Z]{1,4}号)", flat)
    if m:
        rec["seat"] = m.group(1)
    # 席别：出现在座位号后面那一行
    for i, ln in enumerate(lines):
        if re.fullmatch(r"商务座|特等座|一等座|二等座|无座|软座|硬座|软卧|硬卧|一等卧|二等卧|新空调硬座", ln):
            rec["trip_class"] = ln
            break

    # 票价 / 退票费 / 改签费：注意退票和改签票上写的是「退票费」而不是「票价」
    m = re.search(r"票价[：:]*\s*[￥¥]?\s*([\d,]+\.\d{2})", flat)
    if m:
        rec["total"] = _f(m.group(1))
    else:
        m = re.search(r"(退票费|改签费)[：:]*\s*[￥¥]?\s*([\d,]+\.\d{2})", flat)
        if m:
            rec["total"] = _f(m.group(2))
            rec["warn"].append(f"这是{m.group(1)}票，票面没有票价")
    if rec["total"] is not None:
        rec["amount"] = rec["total"]

    # 票面标记：退票 / 改签。票上写「差额退票」「始发改签」，是**红字**印的；
    # 「电子客票号」行尾也跟一份同样的字。两边都看，红字优先（有的版式只在红字里写）。
    rec["red"] = pdf_red_text(path)[:120]
    rec["stamp"] = _stamp_from_text(flat, rec["red"])
    if rec["stamp"]:
        rec["warn"].append(f"票面标记：{rec['stamp']}")

    # 乘车人：身份证（带星号掩码）下一行。掩码有「370102********373X」也有
    # 「3701021980****373X」两种长度，所以星号前后都要放宽。
    for i, ln in enumerate(lines):
        if re.fullmatch(r"\d{6}\d{0,8}\*+[\dXx]{2,4}", ln) and i + 1 < len(lines):
            nxt = lines[i + 1]
            if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", nxt):
                rec["traveler"] = nxt
                rec["tv_from"] = "票面"
            break
    # 购买方
    m = re.search(r"购买方名称[：:]\s*([^统]+?)\s*统一社会信用代码[：:]\s*([0-9A-Z]{15,20})", flat)
    if m:
        rec["buyer"], rec["buyer_id"] = m.group(1), m.group(2)
    rec["seller"] = "中国铁路" if not rec["seller"] else rec["seller"]
    rec["item"] = "铁路旅客运输"
    rec["_stations"] = stations

    # 明细：车次 + 区间 + 日期时间 + 席别座位
    bits = []
    if rec["train_no"]:
        bits.append(rec["train_no"])
    if rec["route"]:
        bits.append(rec["route"])
    if rec["trip_date"]:
        bits.append(f"{rec['trip_date']}{((' ' + rec['trip_time']) if rec['trip_time'] else '')}")
    if rec["trip_class"]:
        bits.append(rec["trip_class"])
    if rec["seat"]:
        bits.append(rec["seat"])
    rec["detail"] = "  ".join(bits)
    for f_, cn in (("no", "发票号码"), ("date", "开票日期"), ("total", "票价")):
        if not rec[f_]:
            rec["warn"].append(f"没取到{cn}")
    return rec


# ---------------------------------------------------------------- PDF：航空运输电子客票行程单
def parse_flight_pdf(lines, rlines, raw, path: Path) -> dict:
    rec = _blank(str(path), "pdf", T_FLIGHT)
    rec["kind"] = "电子发票（航空运输电子客票行程单）"
    flat = "".join(lines)

    m = re.search(r"发票号码[：:]\s*(\d{20})", flat)
    if m:
        rec["no"] = m.group(1)
    # 金额行：「CNY 576.15 CNY 64.22 9% CNY 57.63 CNY 50.00 CNY 0.00 CNY 748.00」
    # 依次是 票价 / 燃油附加费 / 税率 / 增值税税额 / 民航发展基金 / 其他税费 / 合计
    for ln in lines:
        vals = re.findall(r"CNY\s*([\d,]+\.\d{2})", ln)
        if len(vals) >= 3:
            rec["amount"] = _f(vals[0])
            rec["total"] = _f(vals[-1])
            if len(vals) >= 4:
                rec["tax"] = _f(vals[3]) if _f(vals[2]) is None else _f(vals[3])
            break
    # 旅客 + 证件
    m = re.search(r"([\u4e00-\u9fa5]{2,4})\s*(\d{6}\*+[\dXx]{3,4})", flat)
    if m:
        rec["traveler"] = m.group(1)
        rec["tv_from"] = "票面"
    # 填开单位 / 填开日期 —— ⚠️ 必须**按行**找，不能拿整页文本去套正则。
    # 票面上「填开单位:」是标签独占一行（和「销售网点代号: 填开日期:」挤在同一行），
    # 值却在好几行之后：`SDH999/08612351 某某航空股份有限公司（网上直营渠道） 2026年08月27日`。
    # 用 `填开单位:(.+?)(日期)` 整页搜，group(1) 会把中间所有文字全吃进去 ——
    # 用户看到的「机票销售方一堆乱七八糟」就是这么来的。
    src = rlines or []
    for i, ln in enumerate(src):
        if "填开单位" not in ln:
            continue
        for nxt in src[i: i + 8]:
            m = re.search(r"([\u4e00-\u9fa5A-Za-z（）()]{4,40}?(?:公司|航空|旅行社|票务))", nxt)
            if m:
                rec["carrier"] = m.group(1)
                break
        break
    if not rec["date"]:
        for i, ln in enumerate(src):
            if "填开日期" not in ln:
                continue
            for nxt in src[i: i + 8]:
                m = re.search(r"(\d{4}年\d{1,2}月\d{1,2}日)", nxt)
                if m:
                    rec["date"] = norm_date(m.group(1))
                    break
            break
    if not rec["date"]:
        rec["date"] = norm_date(flat)
    # 购买方：票面上是「购方名 税号」并排一行
    # （`测试用会计师事务所（普通合伙） 91310101MA1FL0XXXX`）
    for ln in src:
        m = re.match(r"^([\u4e00-\u9fa5（）()]{4,40})\s+([0-9A-Z]{15,20})$", ln.strip())
        if m:
            rec["buyer"], rec["buyer_id"] = m.group(1), m.group(2)
            break
    if not rec["buyer"]:
        # 兜底：整页里找「单位名 + 税号」。必须确认它**像单位名**，
        # 否则会把票种说明「国内 正常」当购方（去空格后贴着号码，老逻辑就这么中的）。
        for m in re.finditer(r"([\u4e00-\u9fa5（）()]{4,40})([0-9A-Z]{15,20})", flat):
            if _like_unit_name(m.group(1)):
                rec["buyer"], rec["buyer_id"] = m.group(1), m.group(2)
                break
    rec["seller"] = rec["carrier"] or ""
    rec["item"] = "航空旅客运输"
    # 发票号码：票面上是独立的 20 位数字行（「发票号码:」标签和值不在一行）
    if not rec["no"]:
        m = INV_NO_RE.search(flat)
        if m:
            rec["no"] = m.group(0)

    # 航段：「武汉 天河 山航 SC7930 W 2026年08月19日 16:55 W 20K」+ 下一行「烟台 蓬莱」
    # 必须用未去空格的行 —— 列之间就是靠空格分开的，line 版把空格吃光了匹配不上。
    segs = []
    src_lines = rlines if rlines else lines
    for i, ln in enumerate(src_lines):
        m = re.match(r"^([\u4e00-\u9fa5]{2,5})\s+([\u4e00-\u9fa5]{2,5})\s*"
                     r"([\u4e00-\u9fa5]{2,4}?航)?\s*([A-Z]{2}\d{3,4})\s+([A-Z])\s+"
                     r"(\d{4}年\d{1,2}月\d{1,2}日)\s+(\d{1,2}:\d{2})", ln)
        if not m:
            continue
        to_city = to_air = ""
        if i + 1 < len(src_lines):
            m2 = re.match(r"^([\u4e00-\u9fa5]{2,5})\s+([\u4e00-\u9fa5]{2,5})$", src_lines[i + 1])
            if m2 and not re.search(r"航|机场", src_lines[i + 1]):
                to_city, to_air = m2.group(1), m2.group(2)
        segs.append({
            "from_city": m.group(1), "from_air": m.group(2),
            "to_city": to_city, "to_air": to_air,
            "flight": m.group(4), "class": m.group(5),
            "date": norm_date(m.group(6)), "time": m.group(7),
        })
    if segs:
        s0 = segs[0]
        rec["flight_no"] = s0["flight"]
        rec["trip_class"] = s0["class"]
        rec["trip_date"] = s0["date"]
        rec["trip_time"] = s0["time"]
        rec["route"] = f"{s0['from_city']}-{s0['to_city']}" if s0["to_city"] else s0["from_city"]
        rec["detail"] = "  ".join(
            filter(None, [s0["flight"], rec["route"],
                          f"{s0['date']} {s0['time']}".strip(),
                          (s0["class"] + "舱") if s0["class"] else ""]))
    # 航段读不出来时（有的行程单只有一条航段文字没成行），用文件名兜底
    if not rec["route"]:
        rec["route"] = _route_from_name(path.stem)
    if rec["total"] is None:
        rec["total"] = _amount_from_name(path.stem)
        rec["amount"] = rec["total"] if rec["total"] is not None else rec["amount"]
    if not rec["detail"]:
        rec["detail"] = "  ".join(filter(None, [
            rec["flight_no"], rec["route"], rec["trip_date"], rec["traveler"]]))
    for f_, cn in (("no", "发票号码"), ("total", "合计金额")):
        if not rec[f_]:
            rec["warn"].append(f"没取到{cn}")
    return rec


# ---------------------------------------------------------------- PDF：打车行程单（高德 / 滴滴）
def parse_trip_pdf(lines, raw, path: Path) -> dict:
    """行程单是发票的明细附件 —— 发票上只有一句「客运服务费」，起终点只在这儿"""
    rec = _blank(str(path), "pdf", T_TRIP)
    flat = "".join(lines)
    rec["kind"] = "打车行程单"
    rec["seller"] = "高德地图" if "高德" in flat else ("滴滴出行" if "滴滴" in flat else "")

    # 合计：「共计1单行程，合计23.58元」/「共1笔行程， 合计13.00元」
    m = re.search(r"合计\s*([\d,]+\.\d{2})\s*元", flat)
    if m:
        rec["total"] = _f(m.group(1))
        rec["amount"] = rec["total"]
    trips = ""
    m = re.search(r"(?:共计|共)\s*(\d+)\s*[单笔]\s*行程", flat)
    if m:
        trips = f"{m.group(1)} 笔行程"
    # 行程起止日期 / 申请日期
    m = re.search(r"行程起止日期[：:]\s*(\d{4}-\d{2}-\d{2})", flat) or \
        re.search(r"行程时间[：:]\s*(\d{4}-\d{2}-\d{2})", flat)
    if m:
        rec["trip_date"] = norm_date(m.group(1))
    m = re.search(r"申请日期[：:]\s*(\d{4}-\d{2}-\d{2})", flat) or \
        re.search(r"申请时间[：:]\s*(\d{4}-\d{2}-\d{2})", flat)
    if m:
        rec["date"] = norm_date(m.group(1))
    # 行程人手机号 —— ⚠️ 这是**手机号，不是人名**！以前塞进 traveler，
    # 台账「出行人」那格就变成一串数字。现在单独记在 phone 字段，
    # 由 traveler.py 用「手机号 → 姓名」对照表把它翻成人名（见 三层认人）。
    m = re.search(r"行程人手机号[：:]\s*(\d{6,})", flat)
    if m:
        rec["phone"] = m.group(1)

    # 明细行。高德/滴滴的 PDF 会把「起点 / 终点」按视觉列拆成好几行，
    # 所以先按行首序号把折行拼回一条逻辑行，再解析。
    #   滴滴：`1 特价拼车 07-10 22:47 周五 济南市 黄台南路|八涧堡地铁站B口(主路) 历城区|中海·凯旋门东区 6.39 13.00`
    #   高德：`1 T3出行 经济型 2026-07-29 18:29 恩施土家族苗族自治州 …起点… …终点… 23.58元`
    logical = []
    for ln in (raw or "").splitlines():
        t = ln.strip()
        if not t:
            continue
        if re.match(r"^\d+\s", t):
            logical.append(t)
        elif logical:
            logical[-1] += t
    rows = [t for t in logical if re.search(r"[\d,]+\.\d{2}\s*元?\s*$", t)] or logical[:1]
    if rows:
        first = rows[0]
        money = re.search(r"([\d,]+\.\d{2})\s*元?\s*$", first)
        if money and rec["total"] is None:
            rec["total"] = _f(money.group(1))
        car = re.match(r"^\d+\s+([^\s]{2,10})\s+([^\s]{2,10})", first)
        if car and not re.search(r"\d|[:：]", car.group(1)):
            rec["carrier"] = car.group(1)
            rec["trip_class"] = car.group(2)
        t0 = re.search(r"\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}"
                       r"|\d{2}-\d{2}\s+\d{1,2}:\d{2}(?:\s*周[一二三四五六日])?", first)
        if t0:
            rec["trip_time"] = t0.group(0)
        mile = re.search(r"\s(\d{1,3}\.\d{1,2})\s+[\d,]+\.\d{2}\s*元?\s*$", first)
        if mile:
            rec["mileage"] = _f(mile.group(1))
        # 起终点：滴滴写成「A|B」两个带竖线的 POI，最好认
        poi = re.findall(r"([^\s|]+\|[^\s]+)", first)
        if len(poi) >= 2:
            short = lambda x: x.split("|")[-1].strip()
            rec["route"] = f"{short(poi[0])}→{short(poi[1])}"
    # 高德的行程单里起点/终点是两列文字、没有分隔符，靠 PDF 文本分不开；
    # 它的文件名恰好写着「起点-终点」，直接拿来用。
    if not rec["route"]:
        nm = _route_from_name(path.stem)
        if nm:
            rec["route"] = nm.replace("-", "→")
    if rec["total"] is None:
        rec["total"] = _amount_from_name(path.stem)
    if rec["total"] is not None:
        rec["amount"] = rec["total"]

    bits = []
    if rec["carrier"]:
        bits.append(rec["carrier"])
    if rec["route"]:
        bits.append(rec["route"])
    if rec["trip_time"]:
        bits.append(rec["trip_time"])
    if rec["mileage"]:
        bits.append(f"{rec['mileage']}公里")
    if trips:
        bits.append(trips)
    if rec["total"] is not None:
        bits.append(f"{rec['total']:.2f}元")
    rec["detail"] = "  ".join(bits)
    if not rec["no"]:
        rec["warn"].append("行程单没有发票号码（它是发票的明细附件，金额不要重复计入）")
    if len(rows) > 1:
        rec["warn"].append(f"这份行程单有 {len(rows)} 笔行程，已取第 1 笔作为明细")
    return rec


# ---------------------------------------------------------------- PDF：登机牌
def parse_boarding_pdf(lines, raw, path: Path) -> dict:
    rec = _blank(str(path), "pdf", T_BOARD)
    rec["kind"] = "登机牌"
    flat = "".join(lines)
    # 姓名 / 航班 / 日期 / 座位 / 始发 / 目的
    m = re.search(r"([\u4e00-\u9fa5]{2,4})[A-Z]{4,20}", flat)
    if m:
        rec["traveler"] = m.group(1)
    m = re.search(r"([A-Z]{2}\d{3,4})\s*(\d{1,2}[A-Z]{3})?", flat)
    if m:
        rec["flight_no"] = m.group(1)
    m = re.search(r"(\d{1,2}[A-Z]{3})", flat)
    if m:
        rec["trip_date"] = _month_day_to_iso(m.group(1))
    m = re.search(r"座位号?SEATNO\.?\s*([\dA-Z]{1,4})", flat) or re.search(r"(\d{1,2}[A-F])", flat)
    if m:
        rec["seat"] = m.group(1)
    m = re.search(r"([\u4e00-\u9fa5]{2,6})([A-Z]{4,12})([\u4e00-\u9fa5]{2,8})([A-Z]{4,12})", flat)
    if m:
        rec["route"] = f"{m.group(1)}-{m.group(3)}"
    m = re.search(r"(\d{2,4})\s*$", "")
    m = re.search(r"BOARDINGTIME([\d:]{3,5})", flat) or re.search(r"登机时间([\d:]{3,5})", flat)
    if m:
        t = m.group(1)
        rec["trip_time"] = f"{t[:2]}:{t[2:]}" if ":" not in t and len(t) == 4 else t
    m = re.search(r"舱位CLASS([A-Z])", flat)
    if m:
        rec["trip_class"] = m.group(1)
    bits = list(filter(None, [
        rec["flight_no"], rec["route"].replace("-", "→") if rec["route"] else "",
        rec["trip_date"], rec["seat"], rec["trip_class"]]))
    rec["detail"] = "  ".join(bits)
    rec["warn"].append("登机牌只能证明行程，不是报销凭证（找出对应的机票发票或行程单）")
    return rec


def _month_day_to_iso(s):
    """'12AUG' → 'XXXX-08-12'（没有年份，先按当前年补）"""
    mon = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
           "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}
    m = re.match(r"^(\d{1,2})([A-Z]{3})$", (s or "").strip().upper())
    if not m:
        return ""
    d, mm = int(m.group(1)), mon.get(m.group(2))
    if not mm:
        return ""
    from datetime import date
    return f"{date.today().year:04d}-{mm:02d}-{d:02d}"


# ---------------------------------------------------------------- 数电票的「购方 / 销方」
# 版式上购销方是左右并排的两列，抽成文本后变成「购方名 销方名」「购方税号 销方税号」
# 两行并排 —— 名字和税号都会**挤在同一行**里。所以必须按行切、按空格劈，
# 而且要把「标签:值」那种行整行排除，否则会出两种错（都实测见过）：
#   ① 两个名字粘成一个：购方显示成「A公司 B公司」；
#   ② 名字行下面那行「购方开户银行:-;银行账号:-;」也含「银行」二字，
#      被当成公司名 → 销方显示成「购方开户银行:-;银行账号:-;」。
_LABEL_WORDS = ("开户银行", "银行账号", "账号", "地址", "电话", "纳税人识别号",
                "信用代码", "识别号", "名称", "监制", "发票", "备注", "开票人",
                "下载次数", "合计", "金额", "单价", "税率",
                "税务总局", "税务局", "全国统一")
_NAME_WORDS = ("公司", "中心", "局", "店", "厂", "所", "部", "社", "银行", "事务所",
               "酒店", "宾馆", "超市", "医院", "学校", "大学", "集团", "企业",
               "站")
# 纳税人识别号：15~20 位数字/大写字母，两侧不能还粘着字母数字（否则会把长串切一半）
INV_TAX_RE = re.compile(r"(?<![0-9A-Z])[0-9A-Z]{15,20}(?![0-9A-Z])")


def _like_unit_name(s: str) -> bool:
    """这一段文字像不像一个单位名（公司 / 中心 / 事务所…）"""
    t = _despace(s)
    if not (3 <= len(t) <= 40):
        return False
    if any(w in t for w in _LABEL_WORDS):
        return False
    if re.search(r"\d{4}年|\d+\.\d{2}|[¥￥%*]", t):
        return False
    return any(w in t for w in _NAME_WORDS)


def _tax_ids_of(ln: str) -> list:
    """
    这一行里像不像**税号行**（可能并排写着购方、销方两个税号）。
    20 位纯数字是发票号码，不算；明细行、带冒号的标签行一律不算。

    ⚠️ 先按**没去空格**的行找：两个税号是靠空格分开的，去了空格会粘成 36 位的一串，
    结果一个都匹配不上（实测过）。去空格的版本只在原行里找不到时兜底。
    """
    if ":" in ln or "：" in ln:
        return []
    if any(c in ln for c in "*¥￥%") or re.search(r"\d+\.\d{2}", ln):
        return []

    def pick(s):
        return [g for g in INV_TAX_RE.findall(s)
                if re.search(r"\d", g) and not re.fullmatch(r"\d{20}", g)]

    return pick(ln) or pick(_despace(ln))


def _pick_parties(rlines, lines) -> tuple:
    """
    从数电票文本层里认「购买方 / 销售方」的名称与税号，返回 (购方, 购方税号, 销方, 销方税号)。

    版式千变万化，实测过三种：两个名字并排一行 + 两个税号并排一行；名字和税号一行一对；
    名字和号码各占一行且顺序还被内容流打乱。所以不认「第几行」，改用**税号定位**：
    购销方的名称一定紧挨着税号（前后几行内），先围着税号行找，找不齐再全页兜底。
    """
    rl = [x.strip() for x in (rlines or []) if x.strip()]
    if not rl:
        rl = [x for x in (lines or []) if x]

    def scan(seq, names, ids):
        for ln in seq:
            if ":" in ln or "：" in ln:              # 「标签:值」行不是名字
                continue
            rest = ln
            got = _tax_ids_of(ln)
            if got:
                for g in got:
                    if g not in ids:
                        ids.append(g)
                    rest = rest.replace(g, " ")      # 把税号挖掉，剩下的才是名字
            if len(_despace(rest)) > 44:             # 太长的一整段不是名字
                continue
            for piece in re.split(r"\s{1,}", rest):
                if _like_unit_name(piece) \
                        and _despace(piece) not in [_despace(x) for x in names]:
                    names.append(piece)
            if len(names) >= 2 and len(ids) >= 2:
                return True
        return len(names) >= 2 and len(ids) >= 2

    names, ids = [], []
    id_at = next((i for i, ln in enumerate(rl) if _tax_ids_of(ln)), None)
    if id_at is not None:                            # 围着税号行找（最靠得住）
        scan(rl[max(0, id_at - 6): id_at + 7], names, ids)
    if not (len(names) >= 2 and len(ids) >= 2):      # 找不齐再整页扫一遍补
        scan(rl, names, ids)
    return (names[0] if names else "", ids[0] if ids else "",
            names[1] if len(names) > 1 else "", ids[1] if len(ids) > 1 else "")


# ---------------------------------------------------------------- PDF：数电票
def parse_einvoice_pdf(lines, rlines, raw, path: Path) -> dict:
    """
    按数电票版式从 PDF 文本里取字段。
    文本层是「标签在上、值在下」的竖排顺序，所以靠位置关系取：
    号码行后面的几行依次是 日期、购方名称、购方税号、销方名称、销方税号。
    """
    warn = []
    rec = _blank(str(path), "pdf", T_INVOICE)
    rec["warn"] = warn
    total_at = None

    # 票种：标题行通常在最上面，但有的版式前面压着一堆「购买方信息」之类的竖排标签
    # （地铁票就是），所以整页扫；标题后面还常粘着号码
    # （`电子发票（普通发票）26327000001463385744`），所以只取「XX发票（YY）」这一小段。
    # 另外要 NFKC 一下：「电⼦发票」里的 ⼦ 是康熙部首 U+2F26，不是「子」。
    for ln in lines:
        if "发票" not in ln:
            continue
        m = re.search(r"([\u4e00-\u9fa5]{2,8}发票)\s*[（(]([^）)]{2,12})[）)]", ln)
        if m:
            rec["kind"] = f"{m.group(1)}（{m.group(2)}）"
            break
    if not rec["kind"]:
        for ln in lines:
            m = re.search(r"([\u4e00-\u9fa5]{2,8}发票)", ln)
            if m:
                rec["kind"] = m.group(1)
                break
    if rec["kind"]:
        # NFKC 是为了把「电⼦发票」的康熙部首 ⼦ 归回「子」，但它顺手会把全角括号变半角，
        # 所以再翻回来 —— 票面写法还是全角好看。
        k = unicodedata.normalize("NFKC", rec["kind"])
        rec["kind"] = k.replace("(", "（").replace(")", "）")

    # 号码 + 紧跟的日期 / 购销方
    i0 = next((i for i, ln in enumerate(lines) if re.fullmatch(r"\d{20}", ln)), None)
    if i0 is None:
        # 「发票号码:26379...」这种带前缀的写法也要认
        m = re.search(r"发票号码[：:]\s*(\d{20})", "".join(lines))
        if m:
            rec["no"] = m.group(1)
        else:
            m = INV_NO_RE.search(raw)
            if m:
                rec["no"] = m.group(0)
        warn.append("PDF 里没找到独立的发票号码行，号码可能取自其他位置")
    else:
        rec["no"] = lines[i0]
    if i0 is not None:
        tail = lines[i0 + 1: i0 + 8]
        dat = next((x for x in tail if re.search(r"\d{4}年\d{1,2}月\d{1,2}日", x)), "")
        rec["date"] = norm_date(dat)
        # 购销方：按行切、按空格劈两半（两个名字/两个税号挤在同一行，详见 _pick_parties）
        rec["buyer"], rec["buyer_id"], rec["seller"], rec["seller_id"] = \
            _pick_parties(rlines, lines)
    if not rec["date"]:
        m = re.search(r"开票日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)", "".join(lines))
        if m:
            rec["date"] = norm_date(m.group(1))
    if not rec["no"]:
        m = re.search(r"发票号码[：:]\s*(\d{20})", "".join(lines))
        if m:
            rec["no"] = m.group(1)

    # 金额：带两个 ¥ 的那行 = 不含税金额 + 税额；
    #      带中文大写数字的那行 = 价税合计（大写 + 小写）
    # 注意：别用「一行里有几个数字」来判断 —— 明细行「…1239.62 74.381239.62…」也会
    #      凑出好几个数字，容易误取。认 ¥ 符号和中文大写最稳。
    for idx, ln in enumerate(lines):
        money = NUM_RE.findall(ln.replace("¥", " ").replace("￥", " "))
        if CN_NUM_RE.search(ln):
            rec["total_cn"] = _clean(re.sub(r"[¥￥]?\s*-?[\d,]+\.\d{2}", "", ln)) or rec["total_cn"]
            if money:
                rec["total"] = _f(money[-1])
                total_at = idx
            else:
                # 大写金额单独占一行、小写跑到别的行去了（如「65.67¥」，¥ 写在数字后面）。
                # 中文大写是票面上最权威的金额，直接把它算回来。
                cn = cn_to_amount(ln)
                if cn is not None:
                    rec["total"] = cn
                    total_at = idx
            continue
        if rec["amount"] is None and (ln.count("¥") + ln.count("￥")) >= 2 and len(money) >= 2:
            rec["amount"], rec["tax"] = _f(money[0]), _f(money[1])
    # 金额散在相邻几行里的版式（如「价税合计（大写） （小写）」换行后才有 ¥ 和数字）
    if rec["total"] is None:
        for idx, ln in enumerate(lines):
            if not re.search(r"价税合计", ln):
                continue
            for j in range(idx, min(idx + 8, len(lines))):
                if CN_NUM_RE.search(lines[j]) and NUM_RE.search(lines[j]):
                    rec["total"] = _f(NUM_RE.findall(lines[j])[-1])
                    rec["total_cn"] = _clean(re.sub(r"[¥￥]?\s*-?[\d,]+\.\d{2}", "", lines[j]))
                    total_at = j
                    break
            if rec["total"] is not None:
                break
    if rec["amount"] is None and rec["total"] is not None:
        rec["amount"] = rec["total"]

    # 开票人：紧跟在价税合计那行后面的人名（版式里写两遍，取第一个）
    if total_at is not None:
        for ln in lines[total_at + 1: total_at + 4]:
            if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", ln):
                rec["drawer"] = ln
                break
    # 备注＝项目名（在「销方开户银行」那一行之后）
    #
    # ⚠️ 坑：票上备注栏空着的时候，「销方开户银行」后面紧跟的是**明细行**
    # （`*交通运输服务*客运服务费 15.68 3% 0.47`），照单全收的话「项目/事由」就成了
    # 一串乱码 —— 用户在报销单上看到的「事由 / 摘要」就是它。所以明细行必须剔掉。
    k = next((i for i, ln in enumerate(lines) if re.search(r"销[售]?方开户银行", ln)), None)
    if k is not None:
        cut = []
        for ln in lines[k + 1:]:
            if re.search(r"开票人|销售方地址|购买方地址|购买方信息|销售方信息", ln):
                break
            cut.append(ln)
        cut = [x for x in cut if not any(c in x for c in "*¥￥%")]
        rec["project"], rec["project_code"] = split_project(" ".join(cut))

    # 开票人：PDF 文本是「先列所有标签、再列所有值」，所以「开票人：」下一行往往是
    # 发票号码而不是人名 —— 必须限定成人名（2~4 个汉字）才收，否则会串成号码。
    for i, ln in enumerate(lines):
        if re.search(r"开票人[：:]?\s*$", ln) and i + 1 < len(lines):
            nxt = _clean(lines[i + 1])
            if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", nxt):
                rec["drawer"] = nxt
            break

    # 货物服务名称（明细行里 *类别*名称）。
    # 字符集里不能放数字和空白 —— 明细行后面紧跟的是数量和金额（`客运服务费115.67…`），
    # 放进去就会把数字一起吞成「客运服务费115」。
    m = re.search(r"\*[^*]{2,10}\*([\u4e00-\u9fa5A-Za-z（）()·]{2,30})", "".join(lines))
    if m:
        rec["item"] = _clean(m.group(1))

    # 出行明细行（旅客运输服务版式自带），表头是
    #   「出行人 有效身份证件号 出行日期 出发地 到达地 等级 交通工具类型」
    #   未去空格时形如「张三 110101********1234 2026-05-11 武汉站东广场 园博园北 无 公共交通」
    #   也有「赵大勇 110101********1234 2026-08-19 某某科技发展(某某 烟台站-西进站口 出租车」
    # ⚠️ 不能要求固定 7 段：**「等级」常常是空的**（出租车/网约车不印等级），旧写法就整行认不出，
    #    结果票面上明明有出行人却没抓到。现在只要有「姓名 + 掩码证件号 + 日期」，
    #    后面剩几段就认几段（出发地/到达地用前两段，交通工具类型取最后一段）。
    for i, ln in enumerate(rlines or []):
        m = re.match(r"^([\u4e00-\u9fa5]{2,4})\s+([\dXx]*\*[\dXx*]{4,14})\s+"
                     r"(\d{4}-\d{2}-\d{2})\s+(\S.*)$", ln.strip())
        if not m:
            continue
        rest = m.group(4).split()
        # 出发地/到达地被折到下一行（地址里有括号时常见）→ 拼回来
        if i + 1 < len(rlines or []) and rest:
            nxt = (rlines[i + 1] or "").strip()
            if nxt.startswith((")", "）")):
                rest[0] += nxt
        rec["traveler"] = m.group(1)
        rec["tv_from"] = "票面"
        rec["trip_date"] = norm_date(m.group(3))
        if len(rest) >= 2:
            rec["route"] = f"{rest[0]}→{rest[1]}"
        rec["detail"] = "  ".join(
            [rec["trip_date"], rec["route"], rest[-1] if len(rest) >= 3 else ""]).strip()
        if len(rest) >= 3:
            rec["trip_class"] = rest[-1]
        break

    # 备注里藏着的行程（如携程代订机票的「(1) 2025/7/30 济南-兰州 MU6499」）
    if not rec["detail"]:
        m = re.search(r"备注[：:]?\s*(.{4,200})", "".join(lines))
        if m:
            rec["detail"] = _trip_from_remark(m.group(1))
    if not rec["detail"]:
        block = " ".join(lines)
        m = re.search(r"(携程订单[：:].{0,200})", block)
        if m:
            rec["detail"] = _trip_from_remark(m.group(1))

    # 红字票：金额为负，或票面印着红色提示。红字/作废是税务口径的标记，直接影响
    # 能不能报销，所以除了 is_red 这个布尔值，还要在「票面标记」里留个能看的字。
    rec["red"] = pdf_red_text(path)[:120]
    if rec["is_red"] or (rec["total"] is not None and rec["total"] < 0) or "红字" in rec["red"]:
        rec["is_red"] = True
        rec["stamp"] = rec["stamp"] or "红字"
    elif "作废" in rec["red"]:
        rec["stamp"] = rec["stamp"] or "作废"
    # 兜底一：整页去空白后再搜一次 20 位号码。有的 PDF 每个字符都是独立定位的，
    # 按行看就断成好几截（浏览器打印出来的通行费发票就是这样）。
    if not rec["no"]:
        m = INV_NO_RE.search(_despace(raw))
        if m:
            rec["no"] = m.group(0)
    # 兜底二：发票正票上没有行程，但文件名常常写着
    # （高德给发票起的名字就是「【起点-终点-金额元-N个行程】」）
    if not rec["detail"]:
        nm = _route_from_name(path.stem)
        if nm:
            rec["route"] = nm.replace("-", "→")
            rec["detail"] = nm
    # 兜底：连价税合计都没读到时，用文件名里的金额（高德给发票起的名字就带金额）
    if not rec["total"]:
        rec["total"] = _amount_from_name(path.stem)
    if rec["total"] is not None and rec["amount"] is None:
        rec["amount"] = rec["total"]
    for f_, cn in (("no", "发票号码"), ("date", "开票日期"), ("total", "价税合计")):
        if not rec[f_]:
            warn.append(f"没取到{cn}")
    if not rec["buyer"] and not rec["seller"]:
        warn.append("购销方读不出来：这份 PDF 的文字层编码是坏的（网页打印件常见），"
                    "需要票面信息的话请用税务平台下载的 OFD / XML 原件")
    return rec


def _trip_from_remark(text):
    """从发票备注里挖行程，如
    `携程订单:1128141155563375,(1) 2025/7/30 济南-兰州 MU6499 李四 (2) 2025/7/30 兰州-嘉峪关 MU2360 李四`
    → `07-30 济南-兰州 MU6499；07-30 兰州-嘉峪关 MU2360`
    """
    text = _clean(text)
    hits = []
    for m in re.finditer(r"(\d{4}/\d{1,2}/\d{1,2})\s*([\u4e00-\u9fa5]{2,8})[-–—]([\u4e00-\u9fa5]{2,8})\s*"
                         r"([A-Z]{2}\d{3,4}[A-Z]?|[\u4e00-\u9fa5]?\d{3,4}次)", text):
        d = norm_date(m.group(1))
        hits.append(f"{d[5:]} {m.group(2)}-{m.group(3)} {m.group(4)}")
    if hits:
        return "；".join(hits)
    # 铁路/代理：`2025年7月30日 济南 兰州 G123`
    for m in re.finditer(r"(\d{4}年\d{1,2}月\d{1,2}日)\s*([\u4e00-\u9fa5]{2,8})\s+([\u4e00-\u9fa5]{2,8})", text):
        d = norm_date(m.group(1))
        hits.append(f"{d[5:]} {m.group(2)}-{m.group(3)}")
    if hits:
        return "；".join(hits)
    return ""


# ---------------------------------------------------------------- PDF 总入口
def parse_pdf(path: Path) -> dict:
    raw = _pdf_text(path)
    if not raw.strip():
        rec = _blank(str(path), "pdf-img", T_IMAGE)
        rec["kind"] = "图片型 PDF"
        rec["warn"].append("PDF 没有文字层（是扫描件或截图），金额等信息要人工补")
        return rec
    rlines = [x.strip() for x in raw.splitlines() if x.strip()]
    lines = [_despace(x) for x in raw.splitlines()]
    lines = [x for x in lines if x]

    kind = _sniff(lines, raw)
    if kind == "train":
        return parse_train_pdf(lines, raw, path)
    if kind == "flight":
        return parse_flight_pdf(lines, rlines, raw, path)
    if kind == "trip":
        return parse_trip_pdf(lines, raw, path)
    if kind == "boarding":
        return parse_boarding_pdf(lines, raw, path)
    rec = parse_einvoice_pdf(lines, rlines, raw, path)
    # 连「发票」字样都没有的，多半是混进来的附件、合同、报告
    flat = "".join(lines)
    if not rec["no"] and not rec["kind"] and "发票" not in flat:
        rec["ctype"] = T_OTHER
        rec["warn"].append("没认出来这是什么票（可能是附件/合同/报告，不是报销凭证）")
    return rec


# ---------------------------------------------------------------- OFD
def _ofd_xbrl(z) -> dict:
    """铁路电子客票：Doc_0/Attachs/rai_issuer_*.xml 是 XBRL，字段最全"""
    for n in z.namelist():
        if "Attachs/" in n and n.lower().endswith(".xml"):
            try:
                d = z.read(n)
            except Exception:
                continue
            if b"xbrl" in d[:600] or b"rai:" in d[:2000]:
                try:
                    return _xbrl_to_dict(d)
                except Exception:
                    pass
    return {}


def _ofd_textmap(z) -> dict:
    """{TextObject ID: 文字}：把每页 Content.xml 里的文字对象都收进来"""
    m = {}
    for n in z.namelist():
        if not re.search(r"Pages/.*Content\.xml$", n):
            continue
        try:
            root = ET.fromstring(z.read(n))
        except Exception:
            continue
        for el in root.iter():
            name = el.tag.split("}")[-1] if isinstance(el.tag, str) else ""
            if name != "TextObject":
                continue
            oid = el.attrib.get("ID")
            if not oid:
                continue
            txt = "".join(t.text or "" for t in el.iter()
                          if isinstance(t.tag, str) and t.tag.split("}")[-1] == "TextCode")
            if txt.strip():
                m.setdefault(oid, txt.strip())
    return m


def _ofd_tags(z) -> dict:
    """数电票：Tags/CustomTag.xml 是「字段名 → TextObject ID」的映射"""
    out = {}
    for n in z.namelist():
        if not re.search(r"CustomTag\.xml$", n):
            continue
        try:
            root = ET.fromstring(z.read(n))
        except Exception:
            continue
        for el in root.iter():
            name = el.tag.split("}")[-1] if isinstance(el.tag, str) else ""
            if name in ("eInvoice", "CustomTag", "CustomTags"):
                continue
            for c in list(el):
                cn = c.tag.split("}")[-1] if isinstance(c.tag, str) else ""
                if cn == "ObjectRef" and (c.text or "").strip():
                    out.setdefault(name, c.text.strip())
                    break
    return out


_RAI_MAP = {
    "TypeOfVoucher": "kind",
    "ElectronicInvoiceRailwayETicketNumber": "no",
    "DateOfIssue": "date",
    "TypeOfBusiness": "biz",
    "DepartureStation": "dep",
    "DestinationStation": "arr",
    "TrainNumber": "train_no",
    "TravelDate": "trip_date",
    "DepartureTime": "trip_time",
    "SeatLevel": "trip_class",
    "Carriage": "carriage",
    "Seat": "seat_no",
    "Fare": "fare",
    "AmountRefunded": "refund",
    "Name": "traveler",
    "IdNumber": "id_no",
    "TotalAmountExcludingTax": "amount",
    "TaxRate": "tax_rate",
    "TaxAmount": "tax",
    "NameOfPurchaser": "buyer",
    "UnifiedSocialCreditCodeOfPurchaser": "buyer_id",
    "Remarks": "remark",
}

_OFD_FIELD_MAP = {
    "InvoiceNo": "no", "IssueDate": "date",
    "BuyerName": "buyer", "BuyerTaxID": "buyer_id",
    "SellerName": "seller", "SellerTaxID": "seller_id",
    "TaxExclusiveTotalAmount": "amount", "TaxTotalAmount": "tax",
    "TaxInclusiveTotalAmount": "total", "InvoiceClerk": "drawer",
    "Item": "item",
}


def _parse_ofd_rail(g, path: Path) -> dict:
    rec = _blank(str(path), "ofd-rail", T_TRAIN)
    rec["kind"] = g.get("TypeOfVoucher") or "电子发票（铁路电子客票）"
    rec["no"] = g.get("ElectronicInvoiceRailwayETicketNumber", "")
    rec["date"] = norm_date(g.get("DateOfIssue", ""))
    rec["trip_date"] = norm_date(g.get("TravelDate", ""))
    rec["trip_time"] = g.get("DepartureTime", "")
    rec["train_no"] = g.get("TrainNumber", "")
    rec["trip_class"] = g.get("SeatLevel", "")
    dep, arr = _strip_station(g.get("DepartureStation", "")), _strip_station(g.get("DestinationStation", ""))
    rec["route"] = f"{dep}-{arr}" if dep and arr else (dep or arr)
    car, seat = g.get("Carriage", ""), g.get("Seat", "")
    if car or seat:
        rec["seat"] = f"{car}车{seat}号" if car else seat
    rec["traveler"] = g.get("Name", "")
    rec["buyer"] = g.get("NameOfPurchaser", "")
    rec["buyer_id"] = g.get("UnifiedSocialCreditCodeOfPurchaser", "")
    rec["seller"] = "中国铁路"
    rec["item"] = "铁路旅客运输"
    fare = _f(g.get("Fare")) or _f(g.get("AmountRefunded"))
    rec["total"] = fare
    rec["amount"] = _f(g.get("TotalAmountExcludingTax")) or fare
    rec["tax"] = _f(g.get("TaxAmount"))
    if not g.get("Fare") and g.get("AmountRefunded"):
        rec["warn"].append("这是退票票，票面金额是退票费")
    biz = g.get("TypeOfBusiness") or ""
    if biz and biz != "售":
        rec["warn"].append(f"票面业务类型：{biz}")
    bits = list(filter(None, [rec["train_no"], rec["route"],
                              f"{rec['trip_date']} {rec['trip_time']}".strip(),
                              rec["trip_class"], rec["seat"]]))
    rec["detail"] = "  ".join(bits)
    return rec


def _parse_ofd_einvoice(z, path: Path) -> dict:
    """数电票 OFD：CustomTag 给字段→ID，Content.xml 给 ID→文字；再退回 OFD.xml 的 CustomDatas"""
    rec = _blank(str(path), "ofd", T_INVOICE)
    tags, tmap = _ofd_tags(z), _ofd_textmap(z)
    got = {}
    for k, oid in tags.items():
        val = tmap.get(oid)
        if val:
            got[k] = val
    for k, v in got.items():
        field = _OFD_FIELD_MAP.get(k)
        if not field:
            continue
        if field in ("amount", "tax", "total"):
            rec[field] = _f(v)
        elif field == "date":
            rec["date"] = norm_date(v)
        else:
            rec[field] = _clean(v)
    # 兜底：OFD.xml 的 CustomDatas
    if not rec["no"]:
        for n in z.namelist():
            if n.lower().endswith("ofd.xml"):
                try:
                    root = ET.fromstring(z.read(n))
                except Exception:
                    break
                for el in root.iter():
                    if el.tag.split("}")[-1] != "CustomData":
                        continue
                    name, val = el.attrib.get("Name", ""), (el.text or "").strip()
                    if not name or not val:
                        continue
                    if name == "发票号码":
                        rec["no"] = val
                    elif name == "开票日期":
                        rec["date"] = rec["date"] or norm_date(val)
                    elif name == "合计金额":
                        rec["amount"] = rec["amount"] or _f(val)
                    elif name == "合计税额":
                        rec["tax"] = rec["tax"] or _f(val)
                    elif name == "销售方纳税人识别号":
                        rec["seller_id"] = rec["seller_id"] or val
                    elif name == "购买方纳税人识别号":
                        rec["buyer_id"] = rec["buyer_id"] or val
                break
    if rec["total"] is None and rec["amount"] is not None:
        rec["total"] = round(rec["amount"] + (rec["tax"] or 0), 2)
    # 票种：CustomTag 里没有这一项，直接从页面上的文字里捞「电子发票（普通发票）」
    if not rec["kind"]:
        for v in list(got.values()) + list(tmap.values()):
            m = re.search(r"([\u4e00-\u9fa5]{2,8}发票)\s*[（(]([^）)]{2,12})[）)]", v)
            if m:
                rec["kind"] = f"{m.group(1)}（{m.group(2)}）"
                break
    if rec["kind"]:
        rec["kind"] = unicodedata.normalize("NFKC", rec["kind"]).replace("(", "（").replace(")", "）")
    return rec


def parse_ofd(path: Path) -> dict:
    """
    OFD 是 zip 容器，里面藏数据有三条路；三条都没读全就找同目录同名 PDF 补齐。
    （用户下载的 OFD 旁边几乎都配着一份同名 PDF，那条路最省事。）
    """
    rec = None
    try:
        with zipfile.ZipFile(path) as z:
            g = _ofd_xbrl(z)
            if g.get("ElectronicInvoiceRailwayETicketNumber") or "铁路" in str(g.get("TypeOfVoucher", "")):
                rec = _parse_ofd_rail(g, path)
            else:
                rec = _parse_ofd_einvoice(z, path)
    except Exception:
        rec = None
    if rec and rec.get("no") and rec.get("total") is not None and rec.get("seller"):
        return rec
    same = path.with_suffix(".pdf")
    if same.exists():
        try:
            sub = parse_pdf(same)
        except Exception:
            sub = None
        if sub:
            if rec is None:
                out = dict(sub)
                out["src"] = str(path)
                out["src_kind"] = "ofd→同名pdf"
                out["warn"] = list(sub.get("warn") or []) + ["数据取自同目录同名 PDF"]
                return out
            merged = dict(rec)
            added = []
            for k, v in sub.items():
                if k in ("warn", "src", "src_kind"):
                    continue
                if merged.get(k) in (None, "", 0) and v not in (None, "", 0):
                    merged[k] = v
                    added.append(k)
            merged["src_kind"] = "ofd+同名pdf"
            if added:
                merged["warn"] = list(rec.get("warn") or []) + \
                    ["OFD 里信息不全，其余字段取自同目录同名 PDF"]
            return merged
    if rec:
        return rec
    r = parse_from_name(path, note="OFD 里没找到可解析的发票数据")
    r["ctype"] = r["ctype"] or T_OTHER
    return r


# ---------------------------------------------------------------- ZIP / 图片 / 文件名兜底
def parse_zip(path: Path) -> dict:
    """
    zip：优先找里面的发票 XML；没有 XML 就把里面的 pdf/ofd 解出来读。
    （税局「网上购票系统」下发的 zip 就是一层壳，里面装的是同号的 ofd + pdf，
      能读出号码才能跟目录里那份 PDF 归并成同一条，不然台账会多出一条。）
    """
    import tempfile
    names = []
    try:
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.lower().endswith(".xml")]
            for n in names:
                try:
                    data = z.read(n)
                except Exception:
                    continue
                if b"SellerName" in data or b"EIid" in data:
                    rec = parse_xml_bytes(data, path, "zip")
                    rec["warn"].append("数据取自 zip 内的 XML")
                    return rec
            inner = [n for n in z.namelist()
                     if Path(n).suffix.lower() in (".pdf", ".ofd", ".jpg", ".png", ".jpeg")]
            if inner:
                with tempfile.TemporaryDirectory(prefix="invzip_") as td:
                    z.extractall(td)
                    for f in sorted(Path(td).rglob("*")):
                        if not f.is_file() or f.suffix.lower() not in (".pdf", ".ofd", ".jpg", ".png", ".jpeg"):
                            continue
                        try:
                            r = parse_file(f)
                        except Exception:
                            continue
                        if r and r.get("no"):
                            out = dict(r)
                            out["src"] = str(path)
                            out["src_kind"] = "zip内" + str(r.get("src_kind", ""))
                            out["warn"] = list(r.get("warn") or []) + [f"数据取自 zip 内的 {f.name}"]
                            return out
    except Exception as e:
        return parse_from_name(path, note=f"zip 打不开（{type(e).__name__}）")
    return parse_from_name(path, note=f"zip 里没有发票 XML（{len(names)} 个 xml 都不像数电票）")


def parse_image(path: Path) -> dict:
    """.jpg/.png 之类：内容读不出来，但要登记，提醒人工补"""
    rec = _blank(str(path), "image", T_IMAGE)
    rec["kind"] = "图片凭证"
    # 文件名往往是唯一的线索：「武汉公交发票100元.jpg」
    m = re.search(r"([\d,]+(?:\.\d{1,2})?)\s*元", path.stem)
    if m:
        rec["total"] = _f(m.group(1))
        rec["amount"] = rec["total"]
    m = re.search(r"发票|收据|车票|票", path.stem)
    if m:
        rec["kind"] = path.stem[:20]
    rec["detail"] = path.stem
    return rec


def parse_from_name(path: Path, note: str = "") -> dict:
    """最后兜底：只从文件名里抠号码 / 日期 / 金额 / 起终点"""
    stem = path.stem
    rec = _blank(str(path), "name", T_OTHER)
    rec["kind"] = path.stem[:24]
    m = INV_NO_RE.search(stem)
    if m:
        rec["no"] = m.group(0)
        rec["ctype"] = T_INVOICE
    # 先把 20 位号码挖掉再找日期，否则 26349119423005870096 会被抠出 2634-91-19
    stripped = INV_NO_RE.sub(" ", stem)
    m2 = re.search(r"(20\d{2})\s*[-_.年]\s*(\d{1,2})\s*[-_.月]\s*(\d{1,2})", stripped)
    if m2 and _valid_ymd(*m2.groups()):
        rec["date"] = f"{int(m2.group(1)):04d}-{int(m2.group(2)):02d}-{int(m2.group(3)):02d}"
    else:
        m2 = re.search(r"(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)", stripped)
        if m2 and _valid_ymd(*m2.groups()):
            rec["date"] = f"{m2.group(1)}-{m2.group(2)}-{m2.group(3)}"
    # 金额：「…204.pdf」「…289.50.pdf」「打车13-…」
    m3 = re.search(r"(\d{1,5}\.\d{2})\s*(?:元)?", stripped)
    if m3:
        rec["total"] = _f(m3.group(1))
        rec["amount"] = rec["total"]
    # 区间：「汉口至恩施174元」「济南东-烟台 李强 10」→ 起终点
    m4 = re.search(r"([\u4e00-\u9fa5]{2,8})\s*[-–—至到]\s*([\u4e00-\u9fa5]{2,8})", stripped)
    if m4:
        rec["route"] = f"{m4.group(1)}-{m4.group(2)}"
        rec["detail"] = rec["route"]
        if rec["ctype"] == T_OTHER:
            rec["ctype"] = T_TRAIN if "票" in stem else T_OTHER
    # 出行人：文件名里的 2~4 字中文人名（在括号或分隔符之后）
    m5 = re.search(r"[-—\s（(]([\u4e00-\u9fa5]{2,3})\s*\d*\s*[）)]?\s*$", stripped)
    if m5:
        rec["traveler"] = m5.group(1)
        rec["tv_from"] = "文件名"
    if note:
        rec["warn"].append(note)
    rec["warn"].append("只从文件名认出了这些信息，金额/号码可能不全，建议人工核对")
    return rec


# ---------------------------------------------------------------- 统一入口
def parse_file(path: Path) -> dict:
    ext = path.suffix.lower()
    if ext in IMAGE_EXTS:
        rec = parse_image(path)
    elif ext == ".xml":
        try:
            rec = parse_xml_bytes(path.read_bytes(), path, "xml")
        except Exception as e:
            rec = parse_from_name(path, note=f"XML 解析失败（{type(e).__name__}）")
    elif ext == ".zip":
        rec = parse_zip(path)
    elif ext == ".ofd":
        rec = parse_ofd(path)
    elif ext == ".pdf":
        rec = parse_pdf(path)
    else:
        return None
    # 不管哪个解析器，只要出行人是票面里读出来的（XML 的 TravelerName、OFD 的 Name…），
    # 一律标注来源 = 票面；没标的说明是文件名兜底之类，保持原样。
    if rec and rec.get("traveler") and not rec.get("tv_from"):
        rec["tv_from"] = "票面"
    return rec


def collect_files(root, recursive: bool = True) -> list:
    """列出待处理的票据文件（按后缀过滤）"""
    root = Path(root)
    it = root.rglob("*") if recursive else root.glob("*")
    out = []
    for f in it:
        try:
            if not f.is_file() or f.suffix.lower() not in SUPPORTED_EXTS:
                continue
            # 跳过 WPS/Office 开着的临时文件
            if f.name.startswith("~$") or f.name.startswith("."):
                continue
            # 归档/输出目录不是票据来源，别扫进去
            if f.name.lower() in ("merged.pdf",):
                continue
        except OSError:
            continue
        out.append(f)
    # 同一个发票号常常「pdf + ofd + xml」三份都在，按文件名排一下让 xml/zip 先被读到
    order = {".xml": 0, ".zip": 1, ".pdf": 2, ".ofd": 3}
    out.sort(key=lambda p: (str(p.parent), order.get(p.suffix.lower(), 9), p.name))
    return out


# ---------------------------------------------------------------- 解析缓存
# 解析靠 PDF 抽文本，整目录跑一遍要几秒；同一批文件反复扫描时没必要重复解析。
# 缓存键 = 路径 + 大小 + 修改时间，文件一变就自动重解析。
def cache_path():
    try:
        from app_paths import WORK_DIR
    except Exception:
        return None
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    return WORK_DIR / "parse_cache.json"


def load_cache() -> dict:
    p = cache_path()
    if not p or not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_cache(cache: dict):
    p = cache_path()
    if not p:
        return
    try:
        p.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def _src_ver() -> str:
    """本文件的修改时间：解析规则一改，缓存的键就全变，老缓存自动作废"""
    try:
        return str(int(Path(__file__).stat().st_mtime))
    except Exception:
        return "0"


_SRC_VER = _src_ver()


def _ckey(f: Path) -> str:
    try:
        st = f.stat()
        return f"{_SRC_VER}|{f}|{st.st_size}|{int(st.st_mtime)}"
    except OSError:
        return f"{_SRC_VER}|{f}"


def parse_all(files: list, cache: dict = None, progress=None) -> list:
    """
    批量解析（带缓存）。progress(已完成, 总数, 当前文件名) 用于界面显示进度。
    返回 [{rec, key}, ...]，key 是缓存键，供调用方回写。
    """
    cache = {} if cache is None else cache
    out = []
    total = len(files)
    for i, f in enumerate(files, 1):
        k = _ckey(f)
        rec = cache.get(k)
        if rec is None:
            try:
                rec = parse_file(f)
            except Exception as e:
                rec = _blank(str(f), "error", T_OTHER)
                rec["warn"] = [f"解析异常 {type(e).__name__}: {e}"]
            if rec:
                cache[k] = rec
        if rec:
            rec = dict(rec)
            rec.setdefault("src", str(f))
            out.append({"rec": rec, "key": k})
        if progress:
            progress(i, total, f.name)
    return out


# ---------------------------------------------------------------- 去重合并
def _score(r: dict) -> int:
    return (sum(1 for k in ("no", "date", "total", "buyer", "seller", "project",
                            "detail", "route", "traveler")
                if r.get(k))
            + (3 if r.get("src_kind") in ("xml", "zip", "ofd-rail") else 0)
            + (2 if r.get("src_kind") in ("ofd", "ofd→同名pdf") else 0))


def _money_str(v) -> str:
    """
    金额统一成两位小数字符串再参与指纹。
    坑：rec 里的金额是 float 748.0，写进 Excel 再读回来是 int 748，
    str() 出来一个是 '748.0' 一个是 '748' —— 同一份凭证第二次扫描就会被当成新的。
    """
    if v in (None, ""):
        return ""
    try:
        return f"{float(str(v).replace(',', '')):.2f}"
    except (TypeError, ValueError):
        return str(v)


def fingerprint(r: dict) -> str:
    """
    没有发票号码的凭证（行程单/登机牌/图片）靠内容指纹查重，不能靠路径。
    ledger.row_fingerprint() 直接调用本函数，保证收发两端算的是同一个值。
    """
    parts = [r.get("ctype") or "", r.get("date") or r.get("trip_date") or "",
             _money_str(r.get("total")), r.get("detail") or ""]
    if not any(parts[1:]):
        return "path:" + str(r.get("src", "")).lower()
    return "fp:" + hashlib.md5("|".join(parts).encode("utf-8")).hexdigest()[:14]


def dedupe_same_invoice(recs: list) -> list:
    """
    同一个发票号有多个文件（pdf + ofd + zip 里都有 XML）时只留信息最全的那个，
    但把「同号码出现几次」记下来，交给查重逻辑判断是不是重复报销。
    """
    best = {}
    for r in recs:
        if not r:
            continue
        k = r.get("no") or fingerprint(r)
        cur = best.get(k)
        if cur is None:
            r["files"] = [r["src"]]
            r["dup_files"] = 0
            best[k] = r
            continue
        cur["files"].append(r["src"])
        cur["dup_files"] += 1
        if _score(r) > _score(cur):
            files, n = cur["files"], cur["dup_files"]
            r["files"], r["dup_files"] = files, n
            best[k] = r
    return list(best.values())


def dup_key(r: dict) -> str:
    """查重键：发票号码全局唯一；没有号码的凭证用内容指纹"""
    if r.get("no"):
        return f"{r['no']}|{r.get('date','')}|{r.get('total','')}"
    return fingerprint(r)


if __name__ == "__main__":
    import sys
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r"E:\桌\湖北报销发票整理")
    files = collect_files(target)
    print(f"共 {len(files)} 个文件")
    recs = [parse_file(f) for f in files]
    recs = dedupe_same_invoice(recs)
    print(f"去重合并后 {len(recs)} 条\n")
    for r in sorted(recs, key=lambda x: (x.get("ctype", ""), x.get("date", ""))):
        print(f"[{r.get('ctype',''):<6}] {r['no'] or '—':<22} {r.get('date') or '—':<11} "
              f"{str(r.get('total')):>9}  {(r.get('kind') or '')[:16]:<18} "
              f"{(r.get('route') or '')[:22]:<24} {(r.get('traveler') or '')[:6]}"
              + (f"  ⚠{r['warn'][:1]}" if r.get("warn") else ""))
