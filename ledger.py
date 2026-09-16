# -*- coding: utf-8 -*-
"""
发票台账 · 数据访问层
=====================
对外接口和以前一模一样（还是 load / save / add_records / update_rows / …），
但「谁是真身」换了：

    SQLite（发票台账.db）  =  唯一写入源，所有写操作在事务里做
    发票台账.xlsx          =  导出快照，双击就能看、能发给别人、能手工核对

为什么换（原来是「读整张 Excel → 改内存 → 整张写回」）：
  两个人在同一分钟里各改一次，后保存的会把前一个人整片覆盖掉，而且**悄无声息**。
  这是多人使用最致命的问题，也是这份改造清单里的第一条。

现在的写入路径（add_records / update_rows / patch_rows / delete_rows / clear_ledger）
都是「一个事务里读判断 + 改」，SQLite 的写锁会把并发写串行化 ——
第二个事务读到的是第一个已提交的结果，于是「两个人同时入库」不会再丢数据。

关于 Excel 快照：
  · 每次写操作后自动重导一份，保证你打开看到的和库里是一致的；
  · 但**如果你自己在 Excel 里改过**，重导前会把你的版本另存成
    「台账备份/发票台账_手改_<时间>.xlsx」，绝不静默覆盖掉你手改的东西；
  · 想让手改的内容回到库里，用界面上的「从 Excel 导入」。

查重口径（没变）：
  发票号码是全局唯一的（数电票 20 位），号码已在台账里 = 重复；
  行程单/登机牌/图片没有号码，按「类型+日期+金额+明细」的内容指纹；
  再兜一层「文件路径相同 = 同一条记录」。
"""
import hashlib
import shutil
import sqlite3
import threading
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import app_paths
import db
import invparse

SHEET = "发票台账"

# 「凭证类型」和「行程/明细」是后加的，专门用来装发票上没有、只有行程单上才有的信息
# （打车从哪到哪、坐的哪趟车、几车几座），台账里一眼能看全，不用再翻原件。
COLUMNS = [
    ("序号", 6), ("状态", 8), ("凭证类型", 10), ("票面标记", 12), ("出行人", 10),
    ("发票号码", 23),
    ("开票日期", 12),
    ("票种", 24), ("行程/明细", 46),
    ("购方名称", 34), ("销方名称", 34), ("不含税金额", 13), ("税额", 11), ("价税合计", 13),
    ("项目/事由", 44), ("项目编号", 14), ("费用类别", 10), ("报销人", 9),
    ("报销批次", 18), ("入库时间", 17), ("文件路径", 44), ("提示", 30),
]
FIELDS = [c[0] for c in COLUMNS]
MONEY_FIELDS = ("不含税金额", "税额", "价税合计")

# 列定义是唯一的真相：db.py 里的中英对照必须和这里对上，否则界面看到 A、库里存 B
assert set(FIELDS) == set(db.FIELDS), "ledger.COLUMNS 与 db.FIELDS 不一致"

_HDR_FILL = PatternFill("solid", fgColor="E8EEF7")
_DONE_FILL = PatternFill("solid", fgColor="EAF6EA")     # 已报销：淡绿
_DUP_FILL = PatternFill("solid", fgColor="FDECEC")      # 重复：淡红
_THIN = Side(style="thin", color="D0D0D0")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


class LedgerBusy(RuntimeError):
    """台账正被别的写入占着（数据库锁没抢到）"""


# 导 Excel 是个「读全表 + 写文件」的活，两个请求同时干会互相踩（甚至撞同一个文件）。
# 串行化它；真正的写入早就提交到数据库了，这里排队只影响快照生成的速度。
_export_lock = threading.Lock()


def _ledger_file() -> Path:
    """Excel 快照的位置（每次现取，见 db.db_path() 里的说明）"""
    return Path(app_paths.LEDGER_FILE)


def _bak_dir() -> Path:
    return Path(app_paths.LEDGER_BAK)


def _blank_row() -> dict:
    return {f: "" for f in FIELDS}


def key_of(row: dict) -> str:
    """查重键：发票号码 + 开票日期 + 价税合计"""
    return f"{row.get('发票号码','')}|{row.get('开票日期','')}|{row.get('价税合计','')}"


def row_fingerprint(row: dict) -> str:
    """
    没有发票号码的凭证（打车行程单、机票行程单、图片）的行指纹。
    直接复用 invparse.fingerprint()，只维护一套算法 —— 否则同一份行程单
    每次扫描都会被当成新记录，台账里越积越多。
    """
    return db._fp_of(row)


# ------------------------------------------------------------------ 读
def load() -> list:
    """读台账（全部未删除的行，按入库顺序编号 1..N）"""
    return db.all_rows()


def count() -> int:
    return db.count()


# ------------------------------------------------------------------ 备份
def _uniq(dir_: Path, stem: str, suffix: str) -> Path:
    """同一秒里连做两次备份会撞名，撞了就直接覆盖上一份 —— 那就等于没备份。
    所以加序号退避。"""
    dst = dir_ / f"{stem}{suffix}"
    n = 1
    while dst.exists() and n < 100:
        dst = dir_ / f"{stem}_{n}{suffix}"
        n += 1
    return dst


def backup() -> str:
    """把台账留一份到「台账备份」：数据库快照 + Excel 快照，返回主备份路径。

    数据库用 sqlite 自己的 backup API 拷，不是直接复制文件 —— WAL 模式下还有一部分
    最近提交的内容可能还在 -wal 里，直接拷 .db 会拷到一份不完整的。
    """
    bak = _bak_dir()
    bak.mkdir(parents=True, exist_ok=True)
    stem = f"发票台账_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    out = ""

    p = db.db_path()
    if p.exists():
        dst = _uniq(bak, stem, ".db")
        try:
            src = sqlite3.connect(str(p))
            tgt = sqlite3.connect(str(dst))
            with tgt:
                src.backup(tgt)
            src.close()
            tgt.close()
            out = str(dst)
        except Exception:
            try:
                if dst.exists():
                    dst.unlink()
            except OSError:
                pass

    lf = _ledger_file()
    if lf.exists():
        dst2 = _uniq(bak, stem, ".xlsx")
        try:
            shutil.copy2(lf, dst2)
            out = out or str(dst2)
        except OSError:
            pass

    # 只留最近 20 份（两类各算各的，免得一边多把另一边挤没了）
    for pat in ("发票台账_*.db", "发票台账_*.xlsx"):
        olds = sorted(bak.glob(pat))
        for q in olds[:-20]:
            try:
                q.unlink()
            except OSError:
                pass
    return out


# ------------------------------------------------------------------ Excel 快照
def _file_hash(p) -> str:
    try:
        return hashlib.md5(Path(p).read_bytes()).hexdigest()
    except OSError:
        return ""


def _keep_handedited() -> str:
    """把「用户手改过的那份 Excel」另存一份，别让程序把它盖掉"""
    lf = _ledger_file()
    if not lf.exists():
        return ""
    bak = _bak_dir()
    bak.mkdir(parents=True, exist_ok=True)
    dst = _uniq(bak, f"发票台账_手改_{datetime.now():%Y%m%d_%H%M%S}", ".xlsx")
    try:
        shutil.copy2(lf, dst)
        return str(dst)
    except OSError:
        return ""


def _write_excel(rows: list):
    """把行写成 Excel 快照（表头样式、金额格式、已报销/重复的行底色都在这里）"""
    lf = _ledger_file()
    lf.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET

    for i, (name, width) in enumerate(COLUMNS, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.font = Font(bold=True, size=10)
        c.fill = _HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _BORDER
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 22

    for ri, row in enumerate(rows, start=2):
        for ci, name in enumerate(FIELDS, start=1):
            v = row.get(name, "")
            if name in MONEY_FIELDS:
                v = _num(v)
            c = ws.cell(row=ri, column=ci, value=v)
            c.font = Font(size=10)
            c.border = _BORDER
            if name in MONEY_FIELDS:
                c.number_format = "#,##0.00"
                c.alignment = Alignment(horizontal="right")
            elif name in ("序号", "状态", "凭证类型", "开票日期", "费用类别", "报销人"):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(vertical="center", wrap_text=False)
        tip = str(row.get("提示", ""))
        if row.get("状态") == "已报销":
            fill = _DONE_FILL
        elif "重复" in tip or "已报销" in tip:
            fill = _DUP_FILL
        else:
            fill = None
        if fill:
            for ci in range(1, len(FIELDS) + 1):
                ws.cell(row=ri, column=ci).fill = fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(FIELDS))}{max(len(rows) + 1, 2)}"
    try:
        wb.save(str(lf))
    except PermissionError as e:
        raise LedgerBusy(f"台账文件正被占用（多半是 Excel 开着），请先关掉它再试：{lf}") from e


def export_excel(check_handedit: bool = True) -> dict:
    """重导 Excel 快照 —— 永远重新读一次最新数据，不拿调用方手里的旧快照。

    为什么要重读：两个请求几乎同时入库时，A 先在事务里提交、B 后提交；如果 A 拿自己
    那一刻的行去写文件，就会把 B 刚加的那行从 Excel 里抹掉 —— 数据库是对的、快照是错的，
    这是最难查的一种不一致。

    check_handedit：导之前先看一眼「现在磁盘上这份是不是在 Excel 里被改过」——
    是的话先留一份「_手改_」备份，再覆盖。宁可多留一个文件，也不能让人手工改的东西
    悄没声地没了。
    """
    with _export_lock:
        rows = load()
        info = {"backup": "", "rows": len(rows)}
        if check_handedit:
            last = db.meta_get("excel_export_hash")
            cur = _file_hash(_ledger_file())
            if last and cur and cur != last:
                info["backup"] = _keep_handedited()
                if info["backup"]:
                    db.audit("检出 Excel 被外部修改", "ledger", "",
                             before={"hash": cur}, after={"kept": info["backup"]})
        _write_excel(rows)
        db.meta_set("excel_export_hash", _file_hash(_ledger_file()))
        return info


def save(rows: list, do_backup: bool = True):
    """整表写回。

    老接口，语义是「这些行就是台账的全部内容」——给「从 Excel 导入」和诊断脚本用。
    日常的增删改走 add_records / update_rows / …（那些是增量的，不会整表覆盖）。
    """
    if do_backup:
        backup()
    db.replace_all(rows)
    export_excel(check_handedit=False)


def _num(v):
    if v in (None, ""):
        return None
    try:
        return round(float(str(v).replace(",", "")), 2)
    except (TypeError, ValueError):
        return v


# ------------------------------------------------------------------ 查重 / 追加
def index_by_no(rows: list) -> dict:
    """{发票号码: [行...]}，同一号码可能有多行（历史重复入过）"""
    idx = {}
    for r in rows:
        no = str(r.get("发票号码", "")).strip()
        if no:
            idx.setdefault(no, []).append(r)
    return idx


def reindex(rows: list) -> list:
    """给每一行重排序号（界面上靠序号定位行，必须连续且唯一）"""
    for i, r in enumerate(rows, start=1):
        r["序号"] = i
    return rows


# 老台账升级时要补的字段：rec 里的值 → 台账列
_BACKFILL_FIELDS = (
    ("凭证类型", ("ctype",)),
    ("票面标记", ("stamp",)),
    ("发票号码", ("no",)),
    ("开票日期", ("date", "trip_date")),
    ("票种", ("kind",)),
    ("行程/明细", ("detail",)),
    ("购方名称", ("buyer",)),
    ("销方名称", ("seller",)),
    ("不含税金额", ("amount",)),
    ("税额", ("tax",)),
    ("价税合计", ("total",)),
    ("项目/事由", ("project",)),
    ("项目编号", ("project_code",)),
    # 出行人 = 票面上这个人是谁（乘车人 / 旅客 / 打车出行人）。
    # ⚠️ 跟「报销人」是两个概念：报销人是谁提交的这张单子，由入库时手填。
    ("出行人", ("traveler",)),
)


def _backfill(conn, row, rec: dict) -> bool:
    """
    给历史行补上新字段（凭证类型 / 行程明细…）。
    只填空值 —— 用户已经看过、改过的内容一律不动，防止升级时把人工修正冲掉。
    """
    fields = {}
    for col, keys in _BACKFILL_FIELDS:
        if str(row[db.COL2DB[col]] or "").strip():
            continue
        for k in keys:
            v = rec.get(k)
            if v not in (None, ""):
                fields[col] = v
                break
    if not fields:
        return False
    db.update_row_by_id(conn, row["id"], fields)
    return True


def add_records(recs: list, category: str = "", person: str = "", batch: str = "",
                source_name: str = "") -> dict:
    """
    把解析出来的凭证登记进台账。
    返回 {"added":[...], "in_ledger":[...], "dup":[...], "manual":[...]}
      added     新登记的行
      in_ledger 台账里已经有（没重复登记，只提示）
      dup       台账里已有且状态是「已报销」→ 重复报销风险，标红
      manual    连号码带金额都读不出来的文件，需要人工看

    ⚠️ 整批在一个事务里做完：以前是「先 load 全表、在内存里加、再整表写回」，
    两个人同时入库时后写的会把先写的覆盖掉。现在每条都是「事务内查一次库 → 插/补」，
    并发入库不会互相丢数据。
    """
    added, in_ledger, dup, manual = [], [], [], []
    patched = 0
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    changed = False

    try:
        with db.write_tx() as conn:
            for rec in recs:
                no = (rec.get("no") or "").strip()
                ctype = (rec.get("ctype") or "").strip()
                date = rec.get("date") or rec.get("trip_date") or ""
                total = rec.get("total")
                detail = rec.get("detail") or ""
                # 号码和内容两条线索一条都没有 → 真的认不出来，交给人工
                if not no and total is None and not detail:
                    manual.append(rec)
                    continue
                fp = "" if no else invparse.fingerprint(rec)
                exists = db.find_existing(conn, no=no, fp=fp, src=rec.get("src"))
                if exists:
                    done = [r for r in exists if r["status"] == "已报销"]
                    target = done[0] if done else exists[0]
                    if _backfill(conn, target, rec):
                        patched += 1
                        changed = True
                    tip = ("⚠ 重复：台账中已存在且状态为【已报销】，同一张票不能报两次"
                           if done else "台账中已有此凭证（未报销），未重复登记")
                    (dup if done else in_ledger).append({
                        "rec": rec, "tip": tip,
                        "row": {col: target[dbc] for col, dbc in db.COL2DB.items()}})
                    continue

                row = _blank_row()
                row.update({
                    "状态": "未报销",
                    "凭证类型": ctype,
                    "票面标记": rec.get("stamp", ""),
                    "出行人": rec.get("traveler", ""),
                    "发票号码": no,
                    "开票日期": date,
                    "票种": rec.get("kind", ""),
                    "行程/明细": detail,
                    "购方名称": rec.get("buyer", ""),
                    "销方名称": rec.get("seller", ""),
                    "不含税金额": rec.get("amount"),
                    "税额": rec.get("tax"),
                    "价税合计": total,
                    "项目/事由": rec.get("project", ""),
                    "项目编号": rec.get("project_code", ""),
                    "费用类别": category or guess_category(rec),
                    # 报销人 = 谁提交的这张单子，只有入库时手填的那份算数。
                    # ⚠️ 不要拿 rec 里的 traveler 顶上 —— 那是「出行人」，两个概念
                    "报销人": person,
                    "报销批次": batch,
                    "入库时间": now,
                    "文件路径": rec.get("src", ""),
                    "提示": "；".join([*rec.get("warn", []),
                                      *(["红字发票（负数）"] if rec.get("is_red") else [])]),
                })
                try:
                    rid = db.insert_row(conn, row, fp=fp,
                                        path_key=db.norm_path(rec.get("src")))
                except sqlite3.IntegrityError:
                    # 极端情况：同一瞬间另一个请求也把这张票插了进来，唯一索引把它挡下了。
                    # 这不是错误 —— 重新查一次库，按「台账里已有」处理就好。
                    again = db.find_existing(conn, no=no, fp=fp, src=rec.get("src"))
                    if again:
                        tgt = again[0]
                        done_t = tgt["status"] == "已报销"
                        (dup if done_t else in_ledger).append({
                            "rec": rec,
                            "tip": ("⚠ 重复：台账中已存在且状态为【已报销】，同一张票不能报两次"
                                    if done_t else "台账中已有此凭证（未报销），未重复登记"),
                            "row": {c: tgt[d] for c, d in db.COL2DB.items()}})
                    continue
                row["_id"] = rid
                added.append(row)
                changed = True
    except sqlite3.OperationalError as e:
        raise LedgerBusy(f"台账数据库正忙（{e}），请稍后重试") from e

    allrows = db.all_rows()
    if changed:
        export_excel()
        _fill_seq(added, allrows)
        db.audit("入库", "ledger", "",
                 after={"新增": len(added), "已有": len(in_ledger), "重复风险": len(dup),
                        "需人工": len(manual), "补全老行": patched,
                        "新增号码": [r.get("发票号码") or "(无号码)"
                                     for r in added][:200]})
    return {"added": added, "in_ledger": in_ledger, "dup": dup, "manual": manual,
            "patched": patched, "all": allrows}


def _fill_seq(added: list, allrows: list):
    """给刚加进去的行补上「序号」（界面/日志里会用到）"""
    idx = {db.norm_path(r.get("文件路径")): r.get("序号") for r in allrows}
    for r in added:
        s = idx.get(db.norm_path(r.get("文件路径")))
        if s is not None:
            r["序号"] = s


CATEGORY_HINTS = [
    ("交通费", ("运输", "客运", "车票", "机票", "行程单", "通行费", "出租车", "加油", "燃油")),
    ("住宿费", ("住宿", "客房", "酒店", "旅馆")),
    ("餐饮费", ("餐饮", "餐费", "食品", "酒")),
    ("办公费", ("办公", "文具", "耗材", "打印", "纸")),
    ("会议费", ("会议", "会务")),
    ("培训费", ("培训", "教育")),
    ("咨询审计费", ("咨询", "审计", "鉴证", "服务费", "会计")),
    ("通信费", ("通信", "话费", "电信", "移动", "联通")),
]


def guess_category(rec: dict) -> str:
    """按凭证类型 / 货物服务名称猜个费用类别（只是初值，用户可以在台账里改）"""
    ctype = rec.get("ctype") or ""
    if ctype in ("火车票", "飞机行程单", "打车行程单", "登机牌"):
        return "交通费"
    text = f"{rec.get('item','')} {rec.get('seller','')} {rec.get('detail','')}"
    for cat, words in CATEGORY_HINTS:
        for w in words:
            if w in text:
                return cat
    return "其他"


# ------------------------------------------------------------------ 改状态
def update_rows(seqs, **fields) -> int:
    """按「序号」批量改字段（标已报销、改费用类别、填报销人/批次…）"""
    fields = {k: v for k, v in fields.items() if k in FIELDS}
    if not fields:
        return 0
    changes = []
    try:
        with db.write_tx() as conn:
            ids = db.seq_to_id(conn, seqs)
            n = 0
            for rid in ids:
                cur = conn.execute("SELECT * FROM invoices WHERE id = ?", (rid,)).fetchone()
                if cur is None:
                    continue
                f = dict(fields)
                if f.get("状态") == "已报销" and not f.get("报销批次"):
                    if not str(cur["batch"] or "").strip():
                        f["报销批次"] = _today_batch()
                # 审计只留「这次动了哪几个字段」的改前值，不整行抄一遍
                before = {k: cur[db.COL2DB[k]] for k in f if k in db.COL2DB}
                if db.update_row_by_id(conn, rid, f):
                    changes.append({"before": before, "after": f})
                    n += 1
    except sqlite3.OperationalError as e:
        raise LedgerBusy(f"台账数据库正忙（{e}），请稍后重试") from e
    if n:
        export_excel()
        db.audit("修改台账", "invoice", "",
                 before=[c["before"] for c in changes],
                 after={"fields": fields, "rows": n})
    return n


def _today_batch() -> str:
    return datetime.now().strftime("%Y-%m-%d") + " 报销单"


def patch_rows(mapping: dict) -> int:
    """
    一次改多行、而且每行改的字段还不一样 —— 给「重新识别」用。
    mapping = {序号: {字段: 值}}
    """
    if not mapping:
        return 0
    try:
        with db.write_tx() as conn:
            # 序号 = 「按 id 排序后的第几行」，和 all_rows()/get_ledger 用的是同一个顺序
            all_seqs = [r["id"] for r in conn.execute(
                "SELECT id FROM invoices WHERE is_deleted = 0 ORDER BY id")]
            n = 0
            for seq, f in mapping.items():
                if not str(seq).strip().isdigit():
                    continue
                i = int(seq)
                if not (0 < i <= len(all_seqs)):
                    continue
                clean = {k: v for k, v in (f or {}).items() if k in FIELDS}
                if clean and db.update_row_by_id(conn, all_seqs[i - 1], clean):
                    n += 1
    except sqlite3.OperationalError as e:
        raise LedgerBusy(f"台账数据库正忙（{e}），请稍后重试") from e
    if n:
        export_excel()
        db.audit("重新识别", "ledger", "",
                 after={"rows": n,
                        "fields": sorted({k for f in mapping.values() if f for k in f})})
    return n


# ------------------------------------------------------------------ 删除 / 清空
# 用户要求：「凭证台账没有清除或者删除的功能，要是都堆在一块，下次增加新的呢」
# 所以这里给两条路：① 按序号删掉选中的行；② 一把清空整张台账。
# 两个都会先走 backup()，在「台账备份」里留一份，删错了还能捞回来。
# 删除是「软删除」：行还在库里（审计追得回来），只是界面和查重都看不见它了。
def delete_rows(seqs) -> int:
    """按「序号」批量删行，返回删掉的条数"""
    seqs = [s for s in (seqs or [])]
    if not seqs:
        return 0
    backup()
    before = []
    try:
        with db.write_tx() as conn:
            ids = db.seq_to_id(conn, seqs)
            if not ids:
                return 0
            for rid in ids:
                r = conn.execute("SELECT * FROM invoices WHERE id = ?", (rid,)).fetchone()
                if r is not None and not r["is_deleted"]:
                    before.append({k: v for k, v in dict(r).items()})
            db.soft_delete(conn, ids)
    except sqlite3.OperationalError as e:
        raise LedgerBusy(f"台账数据库正忙（{e}），请稍后重试") from e
    n = len(before)
    if n:
        export_excel()
        db.audit("删除台账行", "invoice", "", before=before)
    return n


def clear_ledger() -> int:
    """清空台账（只留表头）。返回清掉的条数"""
    n = db.count()
    if not n:
        return 0
    backup()
    try:
        with db.write_tx() as conn:
            db.clear_all(conn)
    except sqlite3.OperationalError as e:
        raise LedgerBusy(f"台账数据库正忙（{e}），请稍后重试") from e
    export_excel()
    db.audit("清空台账", "ledger", "", before={"count": n})
    return n


def distinct(field: str) -> list:
    """台账里某个字段出现过的所有取值（排序后），给界面的下拉用"""
    vals = {str(r.get(field) or "").strip() for r in load()}
    vals.discard("")
    return sorted(vals, reverse=True)


# ------------------------------------------------------------------ 汇总
def aggregate(rows: list) -> dict:
    """按 月份 / 费用类别 / 销方 / 项目 汇总（默认只统计未报销的，可再筛）"""
    def blank():
        return {"张数": 0, "不含税金额": 0.0, "税额": 0.0, "价税合计": 0.0}

    def add(d, k):
        if not k:
            k = "(未填)"
        c = d.setdefault(k, blank())
        c["张数"] += 1
        for f in MONEY_FIELDS:
            c[f] += float(_num(r.get(f)) or 0)
        return c

    out = {"按月": {}, "按凭证类型": {}, "按费用类别": {}, "按销方": {}, "按项目": {}, "按状态": {}}
    for r in rows:
        d = str(r.get("开票日期", ""))
        add(out["按月"], d[:7] if len(d) >= 7 else "(无日期)")
        add(out["按凭证类型"], str(r.get("凭证类型", "")))
        add(out["按费用类别"], str(r.get("费用类别", "")))
        add(out["按销方"], str(r.get("销方名称", "")))
        add(out["按项目"], str(r.get("项目/事由", ""))[:40])
        add(out["按状态"], str(r.get("状态", "")))
    return out


def export_summary(path, rows: list):
    """把汇总导成一个 Excel（多张 sheet），台账本体不动"""
    agg = aggregate(rows)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "总览"

    def money(ws, r, c, v, fmt="#,##0.00"):
        cell = ws.cell(row=r, column=c, value=round(float(v or 0), 2))
        cell.number_format = fmt
        return cell

    ws.append(["发票台账汇总"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([])
    pending = [r for r in rows if r.get("状态") != "已报销"]
    ws.append(["台账张数", len(rows), "待报销张数", len(pending),
               "已报销张数", len(rows) - len(pending)])
    ws.append(["待报销价税合计", round(sum(float(_num(r.get("价税合计")) or 0) for r in pending), 2)])
    for i, w in zip("ABCDE", (16, 12, 14, 12, 14)):
        ws.column_dimensions[i].width = w

    for name, data in (("按月", agg["按月"]), ("按凭证类型", agg["按凭证类型"]),
                       ("按费用类别", agg["按费用类别"]), ("按销方", agg["按销方"]),
                       ("按项目", agg["按项目"]), ("按状态", agg["按状态"])):
        s = wb.create_sheet(name)
        s.append([name[2:] if name.startswith("按") else name, "张数", "不含税金额", "税额", "价税合计"])
        for c in s[1]:
            c.font = Font(bold=True)
            c.fill = _HDR_FILL
        for k, v in sorted(data.items(), key=lambda kv: -kv[1]["价税合计"]):
            s.append([k, v["张数"], None, None, None])
            money(s, s.max_row, 3, v["不含税金额"])
            money(s, s.max_row, 4, v["税额"])
            money(s, s.max_row, 5, v["价税合计"])
        s.column_dimensions["A"].width = 44
        for c in "BCDE":
            s.column_dimensions[c].width = 14
        s.freeze_panes = "A2"
    wb.save(str(path))
    return str(path)


if __name__ == "__main__":
    rs = load()
    print(f"台账现有 {len(rs)} 行")
    if rs:
        a = aggregate(rs)
        print("\n按状态：")
        for k, v in a["按状态"].items():
            print(f"  {k}: {v['张数']} 张，价税合计 {v['价税合计']:,.2f}")
        print("\n按月：")
        for k, v in sorted(a["按月"].items()):
            print(f"  {k}: {v['张数']} 张，价税合计 {v['价税合计']:,.2f}")
