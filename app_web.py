# -*- coding: utf-8 -*-
"""
发票报销工具 · 界面主程序
==========================
架构：Python 后端（本地 HTTP 服务，只监听 127.0.0.1） + 系统默认浏览器打开界面
      bat 启动 → 起服务 → 浏览器打开 http://127.0.0.1:<端口>/index.html
      → 页面经 gui/bridge.js 调后端接口

四件事：
  1. 入库建账：递归遍历你指定的「凭证文件夹」，解析下面这些凭证的要素，登记到一个
     Excel 台账里，同一张票自动查重（查重键：发票号码；没号码的行程单走内容指纹）：
       · 发票：PDF / OFD / XML（税局下发的 zip 包里的 XML 也能读）
       · 火车票（铁路电子客票，OFD 里带 XBRL 的最全）、飞机行程单、打车行程单、登机牌
       · 图片凭证（.jpg/.png 等，整张图当一张凭证登记）
  2. 台账管理：在界面上筛选、勾选、批量改状态（未报销→已报销）、改费用类别。
  3. 出单据：把选中的发票生成《费用报销单》/《差旅费报销单》PDF；
     还能把发票 PDF 合并成「一张 A4 上下两张、中间裁开」方便打印。
  4. 出统计：按月 / 费用类别 / 销方 / 项目 汇总，导出 Excel。

为什么不用数据库：台账就是一个人人可看可改的 Excel 文件，单机自用不需要更重的东西。
为什么用浏览器而不是套壳：见 发票管理 项目踩过的坑（WebView2 在本机不稳定）。
"""
import functools
import http.server
import inspect
import json
import os
import re
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from collections import deque
from datetime import datetime
from pathlib import Path

from app_paths import (APP_DIR, CONFIG_FILE, DATA_DIR, DEFAULTS, EDGE, GUI_DIR,
                       HOST, LAUNCH_LOG, LEDGER_FILE, OUTPUT_DIR, PASSWORD, PORT,
                       RECYCLE_DIR, SERVER_MODE, UI_PORT_FILE)

import attachment
import db

UI_PORT = PORT                      # 与「发票管理」错开，两个程序可以同时开；容器里由 FB_PORT 指定
IDLE_TIMEOUT = 300                  # 界面 5 分钟没请求就自动退出，不留后台进程（服务器模式不启用）
API_METHODS = {"poll", "whoami", "get_paths", "get_config", "set_config", "pick_folder",
               "check_source", "upload_info", "start_import", "get_ledger", "mark_rows",
               "update_rows", "refresh_rows", "delete_rows", "clear_ledger",
               "get_phone_map", "set_phone_name",
               "traveler_review", "save_traveler_review",
               "make_report", "report_persons", "merge_invoices", "export_summary",
               "list_outputs", "delete_output",
               "ledger_db_info", "import_from_excel", "audit_list", "list_jobs", "gen_files",
               "list_users", "create_user", "set_user_password", "set_user_role",
               "set_user_disabled", "force_logout",
               "open_folder", "open_file", "reveal", "shutdown",
               "job_retry", "job_cancel"}


def _allowed_methods(role: str) -> list:
    """某个角色**能调的接口全集** —— 界面靠它灰掉自己点不了的按钮。

    为什么整表下发而不是逐个问 can['xxx']：接口一多，前端漏写一个就变成
    「按钮亮着、点了报错」（用户遇到的就是这个）。而 db.ROLE_ALLOW 本来就是
    后端卡口用的那份白名单 —— role 传进来，允许集只有一份，不会两边对不上。
    admin 在白名单里是 None（= 全允许），这里展开成全部接口；未知角色给空表。
    """
    allow = db.ROLE_ALLOW.get(role or "", set())
    if allow is None:
        return sorted(API_METHODS)
    return sorted(allow)

# 耗时写操作：解析原件、写台账、打印 PDF —— 全都不能两个人同时跑。
# 以前只有「入库」有 busy 挡着，别的接口能被同时点：一个人在清空台账、另一个人在入库，
# 两边都不报错，结果谁也说不清。这里把它们统一收进「重活锁」。
HEAVY_JOBS = {"start_import", "refresh_rows", "make_report", "merge_invoices",
              "export_summary", "import_from_excel", "clear_ledger"}
# 真正要很久、必须异步跑（否则浏览器会等到超时）的：只有入库
HEAVY_SPAWN = {"start_import"}
# kind → 真正干活的函数名（没列出来的就是同名：接口方法本身就是干活的）。
# 入库是个例外：start_import 只管「接活 + 起线程」，动手的是 _import_job。
RETRY_IMPL = {"start_import": "_import_job"}

# 任务名（界面上的中文说法）。kind 就是接口名，重试时才能原样找回方法。
JOB_KINDS = {
    "start_import": "入库建账",
    "refresh_rows": "重新识别",
    "make_report": "生成报销单",
    "merge_invoices": "合并 PDF",
    "export_summary": "导出统计",
    "import_from_excel": "从 Excel 导入",
    "clear_ledger": "清空台账",
}

# ======================================================================
# 日志
# ----------------------------------------------------------------------
# 一条日志要有三个去处，少一个排障就断链：
#   ① launch.log 文件       —— 事后翻，重启也还在
#   ② stdout               —— 容器里就是 `docker logs`
#   ③ 内存环形缓冲          —— 界面「运行记录」页实时看（poll 取走）
#
# ⚠️ 以前是三个各自为政：_log_line 只写 ①、Api._log 只进 ③+②、而 do_POST 里
#    接口抛的异常**哪儿都不进**。于是「删除报错」在界面上只显示一句笼统的
#    失败、docker logs 一片空白 —— 用户反馈的「我这看容器日志啥也没有」就是这个。
#    现在统一走 _log_line，一个出口写全三处。
# ======================================================================
LOG_LEVELS = {"DEBUG": 10, "INFO": 20, "WARN": 30, "ERROR": 40}
_LOG_LEVEL = (os.environ.get("FB_LOG_LEVEL") or "INFO").strip().upper()
if _LOG_LEVEL not in LOG_LEVELS:
    _LOG_LEVEL = "INFO"


def _force_utf8_streams():
    """把 stdout/stderr 强制成 UTF-8。

    容器里 locale 常常是 C/POSIX，Python 的 stdout 编码就可能是 ASCII ——
    这时 print 中文直接抛 UnicodeEncodeError。以前的 _safe_print 把它 except 掉了，
    结果是「日志什么都没打出来」，比报错还难查。启动时先改掉编码，从根上避免。
    """
    for nm in ("stdout", "stderr"):
        s = getattr(sys, nm, None)
        if s is None:
            continue
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:                                       # noqa: BLE001
            pass


_force_utf8_streams()

_log_lock = threading.Lock()
_LOGS = deque(maxlen=600)          # 界面「运行记录」读的就是这个
_LOG_SEQ = 0


def _push_log(msg: str, level: str = "INFO") -> int:
    """进内存环形缓冲，返回这条的序号（界面按序号做增量）"""
    global _LOG_SEQ
    with _log_lock:
        _LOG_SEQ += 1
        _LOGS.append({"i": _LOG_SEQ, "t": datetime.now().strftime("%H:%M:%S"),
                      "m": msg, "lv": level})
        return _LOG_SEQ


def _safe_print(msg: str):
    try:
        print(msg, flush=True)
    except UnicodeEncodeError:
        # reconfigure 没生效（被别的库换回了 ASCII）时的兜底：编码成 UTF-8 再写
        try:
            buf = getattr(sys.stdout, "buffer", None)
            if buf is not None:
                buf.write((msg + "\n").encode("utf-8", "replace"))
                buf.flush()
        except Exception:                                       # noqa: BLE001
            pass
    except Exception:                                           # noqa: BLE001
        pass


def _log_line(msg: str, level: str = "INFO", ui: bool = None):
    """统一日志出口：文件 + stdout(docker logs) + 界面「运行记录」。

    ui：要不要也进界面的运行记录。默认「除 DEBUG 外都进」——DEBUG 是逐接口的流水，
        全塞给界面会把用户真正要看的东西冲掉。
    """
    if LOG_LEVELS.get(level, 20) < LOG_LEVELS.get(_LOG_LEVEL, 20):
        return
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [{level}] {msg}"
    try:
        with open(LAUNCH_LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:                                           # noqa: BLE001
        pass
    if ui is None:
        ui = level != "DEBUG"
    if ui:
        _push_log(msg, level)
    _safe_print(line)


# ======================================================================
# 配置
# ======================================================================
def load_config() -> dict:
    cfg = dict(DEFAULTS)
    if CONFIG_FILE.exists():
        try:
            cfg.update(json.loads(CONFIG_FILE.read_text(encoding="utf-8")))
        except Exception:
            pass
    return cfg


def save_config(cfg: dict):
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception as e:
        _log_line(f"配置写入失败：{e}")


def same_path(a, b) -> bool:
    """两个路径是不是同一个目录（Windows 大小写不敏感、正反斜杠都算）"""
    def norm(p):
        s = str(p or "").strip().strip('"').replace("\\", "/").rstrip("/")
        return os.path.normcase(s)
    na, nb = norm(a), norm(b)
    return bool(na) and na == nb


# ======================================================================
# 后端接口
# ======================================================================
# 关键词搜索覆盖的字段（发票要素 + 行程要素 + 归属），按顺序即优先级
KW_FIELDS = ("发票号码", "凭证类型", "销方名称", "购方名称", "项目/事由", "行程/明细",
             "费用类别", "报销人", "出行人", "报销批次", "项目编号", "票面标记")


def filter_rows(rows: list, f: dict) -> list:
    """台账筛选：状态 / 凭证类型 / 费用类别 / 报销批次 / 报销人 / 月份 / 金额区间 / 关键词。
    get_ledger（界面看）和 export_summary（导出 Excel）共用同一套口径，
    免得界面筛完了导出却是全量、对不上。"""
    f = f or {}
    out = rows
    for key, col in (("status", "状态"), ("ctype", "凭证类型"),
                     ("category", "费用类别"), ("batch", "报销批次"), ("person", "报销人")):
        v = str(f.get(key) or "").strip()
        if v:
            out = [r for r in out if str(r.get(col, "") or "").strip() == v]
    if f.get("month"):
        # 行程单类凭证常常没有「开票日期」，退回用「入库时间」的月份，免得一筛就空
        out = [r for r in out
               if str(r.get("开票日期", "") or r.get("入库时间", "")).startswith(str(f["month"]))]
    try:
        lo = float(str(f.get("min") or "").replace(",", ""))
    except ValueError:
        lo = None
    try:
        hi = float(str(f.get("max") or "").replace(",", ""))
    except ValueError:
        hi = None
    if lo is not None or hi is not None:
        def in_range(r):
            try:
                v = float(str(r.get("价税合计") or 0).replace(",", "") or 0)
            except ValueError:
                return False
            return (lo is None or v >= lo) and (hi is None or v <= hi)
        out = [r for r in out if in_range(r)]
    kw = str(f.get("kw", "")).strip()
    if kw:
        out = [r for r in out
               if any(kw in str(r.get(k, "")) for k in KW_FIELDS)]
    return out


# ======================================================================
# 生成文件历史（报销单 / 合并 PDF / 统计 Excel）
# ----------------------------------------------------------------------
# 用户要求「报销单也要有删除功能」——所以生成的东西都记一笔到
# 输出/生成记录.json，界面重新进来也能看到历史上生成过什么、随手删掉。
# 删除走**回收站**（SHFileOperationW + FOF_ALLOWUNDO），不是永久删除，删错能捞。
# ======================================================================
def _history_file() -> Path:
    return Path(OUTPUT_DIR) / "生成记录.json"


def _load_history() -> list:
    try:
        data = json.loads(_history_file().read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_history(items: list):
    try:
        _history_file().write_text(
            json.dumps(items[-300:], ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception as e:
        _log_line(f"生成记录写入失败：{e}")


def _to_recycle_bin(path) -> bool:
    """把文件丢进回收站（不是永久删除）——用户删错了还能捞回来

    本机：走 Windows 回收站（SHFileOperationW + FOF_ALLOWUNDO）。
    容器/NAS：没有 Windows 回收站，改成挪进「数据目录/回收站」，语义一样是「删错能捞」。
    """
    p = Path(path)
    if SERVER_MODE or os.name != "nt":
        import shutil
        try:
            RECYCLE_DIR.mkdir(parents=True, exist_ok=True)
            dst = Path(RECYCLE_DIR) / p.name
            if dst.exists():
                dst = Path(RECYCLE_DIR) / f"{p.stem}_{datetime.now():%Y%m%d_%H%M%S}{p.suffix}"
            shutil.move(str(p), str(dst))
            return True
        except Exception as e:
            _log_line(f"移入回收站目录失败：{type(e).__name__}: {e}")
            return False

    import ctypes
    from ctypes import wintypes

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [("hwnd", wintypes.HWND), ("wFunc", wintypes.UINT),
                    ("pFrom", wintypes.LPCWSTR), ("pTo", wintypes.LPCWSTR),
                    ("fFlags", ctypes.c_uint), ("fAnyOperationsAborted", wintypes.BOOL),
                    ("hNameMappings", ctypes.c_void_p), ("lpszProgressTitle", wintypes.LPCWSTR)]

    try:
        op = SHFILEOPSTRUCTW()
        op.wFunc = 3                              # FO_DELETE
        op.pFrom = str(Path(path).resolve()) + "\0"
        op.fFlags = 0x0040 | 0x0004 | 0x0010      # 允许撤销 | 不再确认 | 不弹错误框
        return ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op)) == 0
    except Exception:
        return False


# ======================================================================
# 文件下发（服务器模式）
# ----------------------------------------------------------------------
# 本机模式下「打开文件 / 打开文件夹」是交给 Windows 资源管理器去干的；
# 容器里没有资源管理器，改成让浏览器直接打开或下载 —— 所以得把这几个目录
# 通过 /files/<别名>/... 挂出去。
#
# 别名是固定的（output / data / source），不是任意路径 —— 免得把容器里
# 整个文件系统暴露给浏览器。路径穿越（../）在 Handler 里另行拦掉。
# ======================================================================
# ---------------------------------------------------------------- 什么不往外发
# /files/ 是「容器里没有资源管理器」的替代品，但它以前把整个数据目录挂了出去 ——
# 那里面有 config.json（口令就在里面）、手机号姓名对照表、日志、台账备份和数据库本体。
# 局域网里谁都能拉到这些东西，等于把口令明着贴在墙上。做成 NAS 版之后这必须先堵上。
# 现在按「目录名 + 文件名 + 后缀」拦一道，只放行真正该给人看的：
# 票据原件、报销单 PDF、台账 Excel、汇总表。
_FS_DENY_DIRS = {"台账备份", "缓存", "browser_debug", "回收站", "__pycache__", ".git"}
_FS_DENY_NAMES = {
    "config.json", "launch.log", "run.log", "ui_port.txt",
    "手机号姓名.json", "出行人核对.json", "生成记录.json", "parse_cache.json",
}
_FS_DENY_EXT = {".db", ".db-wal", ".db-shm", ".py", ".pyc", ".pyo", ".bat", ".ps1",
                ".log", ".json", ".ini", ".toml", ".yml", ".yaml", ".env", ".key", ".pem"}


def _fs_blocked(rel: str) -> bool:
    """这个相对路径该不该被挡在 /files/ 外面（路径里任意一段命中就挡）"""
    for p in (rel or "").replace("\\", "/").split("/"):
        if not p:
            continue
        if p in _FS_DENY_DIRS or p in _FS_DENY_NAMES:
            return True
        if Path(p).suffix.lower() in _FS_DENY_EXT:
            return True
    return False


def _fs_roots() -> dict:
    """可下发的目录：输出目录、数据目录、当前票据文件夹"""
    roots = {"output": Path(OUTPUT_DIR), "data": Path(DATA_DIR)}
    src = str(load_config().get("source") or "").strip()
    if src:
        roots["source"] = Path(src)
    # 长路径优先匹配：输出目录在数据目录里面，先认出更具体的那个
    return dict(sorted(roots.items(), key=lambda kv: -len(str(kv[1]))))


def _url_of(path) -> str:
    """绝对路径 → /files/<别名>/<相对路径>；不在可下发目录里就返回空串"""
    try:
        rp = Path(path).resolve()
    except Exception:
        return ""
    for alias, root in _fs_roots().items():
        try:
            rr = Path(root).resolve()
        except Exception:
            continue
        if rp == rr:
            return f"/files/{alias}/"
        if rr in rp.parents:
            relp = rp.relative_to(rr).as_posix()
            if _fs_blocked(relp):
                return ""      # 被拦下的东西不给 URL，免得生成一个点开就 403 的死链
            return f"/files/{alias}/{relp}"
    return ""


def _resolve_under(root: Path, rel: str) -> Path:
    """把 /files/ 后面的相对路径拼到 root 上；越界（..、绝对路径）返回 None"""
    rel = (rel or "").strip().lstrip("/")
    try:
        target = (Path(root) / rel).resolve()
        rr = Path(root).resolve()
    except Exception:
        return None
    if target != rr and rr not in target.parents:
        return None
    return target


# ======================================================================
# 从本机上传票据（先传到 NAS，再在 NAS 上入库）
# ----------------------------------------------------------------------
# 为什么需要这条通道：程序跑在 NAS 的容器里，看不见你电脑的磁盘 —— 容器里那句
# 「选择文件夹」只能选到 NAS 自己的目录。所以另开一条上传通道，把**你电脑上**的
# 文件夹原样搬进 NAS 的票据目录：
#
#     票据目录/上传/<你选的那个文件夹名>/<里面的子目录…/文件>
#
# 目录结构原样保留是有意为之：发票和它的行程单必须在同一个文件夹里才能配上对
# （见 attachment.py），照抄结构才不会把配对关系拆散。
# ======================================================================
UPLOAD_SUBDIR = "上传"
_UPLOAD_CHUNK = 1024 * 1024
MAX_UPLOAD_BYTES = 512 * 1024 * 1024      # 单个文件上限（发票不可能这么大，防手滑）


def upload_base() -> Path:
    """上传落点的父目录 = 当前票据文件夹。

    注意后半段：界面上传完会把路径框切到「上传」目录下（这样一眼只看传上来的这批），
    而那个路径会同步进配置 —— 所以要在这里把结尾的「上传」剥掉，否则下次上传就落在
    「上传/上传」里，一次次传下去越挖越深。
    """
    src = str(load_config().get("source") or "").strip()
    base = Path(src) if src else (Path(DATA_DIR) / "票据")
    if base.name == UPLOAD_SUBDIR:
        base = base.parent
    return base


def upload_root(make: bool = True) -> Path:
    """上传落点「票据目录/上传」。建不出来就抛异常 —— 宁可明说，
    也别悄悄存到别的地方去（界面显示的落点必须就是程序真正写的地方）。"""
    root = upload_base() / UPLOAD_SUBDIR
    if make:
        root.mkdir(parents=True, exist_ok=True)
    return root


def _safe_rel_parts(rel: str) -> list:
    """把浏览器给的相对路径拆成安全的一段段。
    正常浏览器给的路径（webkitRelativePath）绝不会含 .. —— 一出现就整条判非法。"""
    out = []
    for p in (rel or "").replace("\\", "/").split("/"):
        p = p.replace("\x00", "").strip()
        if p in ("", "."):
            continue
        if p == ".." or p.startswith("..") or ":" in p:
            return []
        out.append(p)
    return out


def _unique_path(path: Path, limit: int = 300) -> Path:
    """同名但内容不一样 → 加 -2 / -3 后缀，两份都留着（覆盖比留两份危险）"""
    if not path.exists():
        return path
    for i in range(2, limit):
        cand = path.with_name(f"{path.stem}-{i}{path.suffix}")
        if not cand.exists():
            return cand
    return path.with_name(f"{path.stem}-{datetime.now():%Y%m%d%H%M%S}{path.suffix}")


# ---------------------------------------------------------------- 按内容认票
# 为什么要这套东西：同一张票常常同时有 OFD 和 PDF，但用户手存的 PDF 文件名
# 是人话（「张三-汉口-郑州东.pdf」），里面根本没有 20 位发票号 —— 只按文件名
# 找 PDF 会找不到，合并出来就少几张（用户实测：选 4 张只合进去 1 张）。
# 所以文件名对不上时，把同目录 PDF 的文字层读出来，按票面里的发票号认。
_PDF_IDX_CACHE = {}                 # {(文件夹, 签名): {发票号: PDF路径}}
_PDF_IDX_MAX = 8                    # 最多缓存 8 个文件夹的索引，够用又不涨内存
_PDF_IDX_LIMIT = 400                # 一个文件夹 PDF 太多就别扫了，别把界面卡住


def _pdf_numbers(p: Path) -> set:
    """读一个 PDF 的文字层，把里面的 20 位发票号全挖出来（读不了就返回空集）"""
    try:
        import invparse
        txt = invparse._pdf_text(p)
    except Exception:
        return set()
    if not txt:
        return set()
    return set(re.findall(r"(?<!\d)\d{20}(?!\d)", txt))


def _folder_pdf_index(folder: Path) -> dict:
    """
    把 folder 下的 PDF 按「票面里的 20 位发票号」建一份索引。

    缓存键带上「PDF 个数 + 最新修改时间」，用户新往文件夹里放了票，下次合并会自动重扫，
    不用重启程序。文件夹太大（>400 个 PDF）就直接放弃，避免为了合并卡住半天。
    """
    try:
        if not Path(folder).is_dir():
            return {}
        pdfs = [p for p in Path(folder).glob("*.pdf")
                if not p.name.startswith(("发票合并", "发票台账汇总", "_"))]
    except OSError:
        return {}
    if not pdfs or len(pdfs) > _PDF_IDX_LIMIT:
        return {}
    try:
        sig = (len(pdfs), round(max(p.stat().st_mtime for p in pdfs), 1))
    except OSError:
        return {}
    key = (str(folder).lower(), sig)
    if key in _PDF_IDX_CACHE:
        return _PDF_IDX_CACHE[key]
    idx = {}
    for p in pdfs:
        for no in _pdf_numbers(p):
            idx.setdefault(no, p)
    if len(_PDF_IDX_CACHE) >= _PDF_IDX_MAX:
        _PDF_IDX_CACHE.clear()
    _PDF_IDX_CACHE[key] = idx
    return idx


class Api:
    def __init__(self):
        self.busy = False
        self.exit_flag = threading.Event()
        self.last_active = time.time()
        # 日志缓冲在模块级（_LOGS/_LOG_SEQ）—— 因为登录、越权这些发生在 UiHandler 里，
        # 拿不到 Api 实例，但一样要出现在界面的「运行记录」里。
        self._progress = {"on": False, "cur": 0, "total": 0, "label": ""}
        self._status = {"text": "就绪", "kind": "ready"}
        self._result = None            # 最近一次入库结果
        self._result_id = 0            # 结果编号：界面靠它判断「这条结果我看过了没」
        self._cfg = load_config()
        self._last_output = []         # 最近生成的文件
        self.running = None
        self._heavy = None             # 当前占着「重活锁」的任务（dict）；None ＝ 没人占用
        self._heavy_lock = threading.Lock()
        self._boot_data()

    # ---------- 启动时的数据层准备 ----------
    def _boot_data(self):
        """建表 → 老台账 Excel 迁进数据库 → 收拾上次没跑完的任务。

        为什么要在这里做：以前台账就是一个 Excel，直接就是数据；现在真身是 SQLite，
        第一次跑新版得把老的 Excel 搬进去，不然用户会看到「台账怎么空了」。
        迁移是幂等的（迁过一次就有标记），所以每次启动调一下很便宜。
        """
        try:
            db.init()
            res = db.migrate_from_excel(LEDGER_FILE)
            if res.get("migrated"):
                self._log(f"📦 老台账已搬进数据库：{res['migrated']} 行"
                          f"（原 Excel 备份为 {Path(res['backup']).name}）")
            elif res.get("skipped") and not db.meta_get("migrated_from_excel"):
                self._log(f"· 台账数据库已就绪（{db.count()} 行）")
            # 第一次跑：设了访问口令但库里还没有任何账号 —— 就用这个口令把管理员建出来。
            # 不这么做的话，用户部署完 NAS 会发现「登录页要用户名，可我没有账号」。
            if db.user_count() == 0 and _cur_password():
                try:
                    db.user_create("admin", _cur_password(), role="admin",
                                   display_name="管理员")
                    self._log("👤 已用访问口令创建管理员账号 admin（登录名 admin）")
                except Exception as e:                          # noqa: BLE001
                    _log_line(f"建管理员账号失败：{e}")
            n = db.job_orphans_fix()
            if n:
                self._log(f"🧹 上次没跑完的任务收了 {n} 个")
        except Exception as e:                                  # noqa: BLE001
            _log_line(f"数据层准备失败：{type(e).__name__}: {e}")
            self._log(f"⚠ 数据层准备失败：{type(e).__name__}: {e}")

    # ---------- 基础设施 ----------
    def touch(self):
        self.last_active = time.time()

    def _set_result(self, data):
        """记录最近一次入库结果。**不清空**：这样界面刷新一下也能把上次结果显示出来，
        重复推给同一个界面由 _result_id 去重。"""
        self._result = data
        self._result_id += 1

    def _log(self, msg: str):
        """操作日志：文件 + stdout(docker logs) + 界面「运行记录」，三处一起写。"""
        _log_line(msg)

    def _set_status(self, text, kind="ready"):
        self._status = {"text": text, "kind": kind}

    def _set_progress(self, on, cur=0, total=0, label=""):
        self._progress = {"on": bool(on), "cur": cur, "total": total, "label": label}
        if on:
            self._job_progress(cur, total)

    # ---------- 后台任务：重活互斥 + 任务留痕（P0-4）----------
    def _job_take(self, name, label, payload=None):
        """抢「重活锁」并建一条任务。抢到返回 (任务号, "")；抢不到返回 (None, 为什么)。"""
        with self._heavy_lock:
            if self._heavy:
                h = self._heavy
                who = h.get("user") or "?"
                return None, (f"「{h.get('label')}」正在执行（{who}），"
                              f"请等它跑完再试")
            jid = ""
            try:
                jid = db.job_create(name, label, payload=payload)
            except Exception as e:                              # noqa: BLE001
                # 任务表写不进去不该把功能整个卡死：照常跑，只是这次没留痕
                _log_line(f"建任务失败：{type(e).__name__}: {e}")
            self._heavy = {"id": jid, "kind": name, "label": label,
                           "user": db.actor(), "at": time.time()}
            self._job_id = jid
        self.busy = True                # 界面靠它禁用按钮，所以同步接口也算「忙」
        if jid:
            try:
                db.job_update(jid, state=db.JOB_RUNNING, started_at=db._now())
            except Exception as e:                              # noqa: BLE001
                _log_line(f"任务转运行中失败：{type(e).__name__}: {e}")
        return jid, ""

    def _job_release(self, jid):
        with self._heavy_lock:
            self._heavy = None
        if getattr(self, "_job_id", "") == jid:
            self._job_id = ""
        self.busy = False

    def _job_finish(self, jid, out, err=""):
        """收尾：成功 / 失败都写回任务（含失败原因，失败的任务点「重试」能再来一次）"""
        if not jid:
            return
        ok = not err and not (isinstance(out, dict) and out.get("error"))
        if not ok and not err:
            err = str(out.get("error") or "执行失败")
        # 只留个摘要：入库结果里的明细清单可能有几百条，整份塞进任务表没意义
        brief = None
        if ok and isinstance(out, dict):
            brief = {k: v for k, v in out.items()
                     if k not in ("manual_list", "dup_list", "rows", "files", "all")}
        try:
            db.job_update(jid, state=db.JOB_DONE if ok else db.JOB_FAILED,
                          ended_at=db._now(), error="" if ok else str(err)[:800],
                          result_json=brief)
        except Exception as e:                                  # noqa: BLE001
            _log_line(f"写任务结果失败：{type(e).__name__}: {e}")

    def _job_call(self, name, fn, args=(), kwargs=None, payload=None):
        """耗时写操作的统一入口：抢锁 → 记任务 → 跑 → 收尾。

        返回什么：同步接口直接返回它本来的结果；抢不到锁返回 {"error": ...}；
        异步接口（入库）立刻返回 {"job": 任务号}。
        """
        label = JOB_KINDS.get(name, name)
        if payload is None:
            payload = self._job_payload(fn, args, kwargs)
        if name in HEAVY_SPAWN:
            return self._job_spawn(name, label, payload, fn, args=args, kwargs=kwargs)
        jid, err = self._job_take(name, label, payload)
        if err:
            self._log(f"⏳ {err}")
            return {"error": err, "busy": True}
        try:
            out = fn(*args, **(kwargs or {}))
        except Exception as e:
            self._job_finish(jid, None, f"{type(e).__name__}: {e}")
            self._job_release(jid)
            self._log(f"❌ {label} 失败：{type(e).__name__}: {e}")
            raise
        self._job_finish(jid, out)
        self._job_release(jid)
        return out

    @staticmethod
    def _job_payload(fn, args, kwargs):
        """把这次调用的参数记下来 —— 失败后点「重试」要能原样再跑一遍"""
        try:
            bound = inspect.signature(fn).bind_partial(*(args or ()), **(kwargs or {}))
            bound.apply_defaults()
            return dict(bound.arguments)
        except (TypeError, ValueError):
            return {"args": list(args or []), "kwargs": dict(kwargs or {})}

    def _job_spawn(self, name, label, payload, fn, args=(), kwargs=None):
        """异步版（入库这种要跑几分钟的）。

        ⚠️ 参数一定要在这里转交出去：早先写成 `fn()` 直接调用，等于把入库目录、
        递归开关全丢了 —— 程序会拿配置文件里的旧默认值去扫，界面上写着 A 目录、
        扫的却是 B 目录（「界面看见的必须等于程序在做的」正是这么破的）。

        ⚠️ 还要先把「谁在操作」抓下来再起线程：后台线程里 contextvars 是**空的**
        （实测 `db.actor()` 返回 ''），不显式带过去的话，入库写出来的每一行
        created_by、审计里的 username 全是空白 —— 多人用了反而查不出「谁入的库」。
        """
        jid, err = self._job_take(name, label, payload)
        if err:
            self._log(f"⏳ {err}")
            return {"error": err, "busy": True}
        who = (db.actor(), db.client_ip(), db.role())
        kw = dict(kwargs or {})

        def runner():
            db.set_actor(*who)
            try:
                out = fn(*args, **kw)
            except Exception as e:                              # noqa: BLE001
                self._job_finish(jid, None, f"{type(e).__name__}: {e}")
                self._log(f"❌ {label} 失败：{type(e).__name__}: {e}")
            else:
                self._job_finish(jid, out)
            finally:
                self._job_release(jid)

        threading.Thread(target=runner, daemon=True).start()
        return {"job": jid, "started": True}

    def _job_progress(self, cur, total):
        """把进度同步进当前任务（供「任务」页看，重开界面也看得到跑到哪了）"""
        jid = getattr(self, "_job_id", "")
        if not jid:
            return
        try:
            db.job_update(jid, done=int(cur or 0), total=int(total or 0))
        except Exception:                                       # noqa: BLE001
            pass

    def job_retry(self, jid=""):
        """重跑一个失败的任务（参数是当初记下来的那份）"""
        j = db.job_get(jid)
        if not j:
            return {"error": "找不到这个任务"}
        if j.get("state") in (db.JOB_QUEUED, db.JOB_RUNNING):
            return {"error": "这个任务还在跑，不用重试"}
        name = j.get("kind") or ""
        if name not in HEAVY_JOBS:
            return {"error": f"这个任务（{name or '未知'}）不支持重试"}
        # 重试 ＝ 原样再跑一遍那个接口，所以权限要按**那个接口**再查一次：
        # 只把 job_retry 这个名字放开是不够的，否则业务员能借重试去清空台账。
        if not db.can(db.actor(), db.role(), name):
            return {"error": f"当前身份不能重跑「{JOB_KINDS.get(name, name)}」"}
        # ⚠️ 要拿**真正干活的那个函数**，不能拿接口外壳：外壳自己也要走任务层抢同一把锁，
        #    于是抢不到、悄悄返回，任务被记成「成功」—— 界面看着好好的，其实什么都没做。
        fn = getattr(self, RETRY_IMPL.get(name, name), None)
        if not callable(fn):
            return {"error": f"接口已不存在：{name}"}
        payload = db.job_retry_payload(jid)
        if payload is None:
            return {"error": "这个任务没留下参数，重试不了（旧版本建的任务）"}
        self._log(f"↻ 重试任务「{JOB_KINDS.get(name, name)}」（原任务 {jid}）")
        if set(payload) == {"args", "kwargs"}:      # 参数签名拿不到时的兜底
            return self._job_call(name, fn, args=payload["args"],
                                  kwargs=payload["kwargs"], payload=payload)
        return self._job_call(name, fn, kwargs=payload, payload=payload)

    def job_cancel(self, jid=""):
        """标记取消。

        真正的取消没法把一个正在解析 PDF 的线程安全掐掉（硬杀会留下半截的台账写入），
        所以这里只把「还在排队」的任务标掉；已经跑起来的老实等它跑完。
        """
        j = db.job_get(jid)
        if not j:
            return {"error": "找不到这个任务"}
        if j.get("state") == db.JOB_QUEUED:
            db.job_update(jid, state=db.JOB_CANCELLED, ended_at=db._now(),
                          error="已取消")
            self._log(f"⛔ 任务已取消：{j.get('label')}")
            return {"ok": True, "cancelled": True}
        if j.get("state") == db.JOB_RUNNING:
            return {"error": "任务已经在跑了，取消不了（等它跑完）"}
        return {"error": "这个任务已经结束了"}

    # ---------- 前端轮询 ----------
    def poll(self, since=0):
        """前端每 0.4s 调一次；日志按 since 增量返回，结果带 result_id 供前端去重"""
        self.touch()
        logs = [x for x in list(_LOGS) if x["i"] > int(since or 0)]
        return {"logs": logs, "seq": _LOG_SEQ, "progress": self._progress,
                "status": self._status, "busy": self.busy,
                "result": self._result, "result_id": self._result_id}

    # ---------- 路径 / 配置 ----------
    def get_paths(self):
        cfg = load_config()
        return {"source": cfg.get("source", ""), "ledger": str(LEDGER_FILE),
                "output": str(OUTPUT_DIR), "app": str(APP_DIR),
                "ledger_exists": LEDGER_FILE.exists()}

    def get_config(self):
        cfg = load_config()
        cfg["ledger_exists"] = LEDGER_FILE.exists()
        cfg["output"] = str(OUTPUT_DIR)
        return cfg

    def set_config(self, patch: dict = None):
        cfg = load_config()
        cfg.update(patch or {})
        save_config(cfg)
        self._cfg = cfg
        return cfg

    def pick_folder(self, initial: str = ""):
        """弹系统目录选择框（用户点了按钮才弹，属于用户主动操作）

        容器/NAS 里没有图形界面，也弹不出来 —— 直接告诉用户手输路径（界面上本来就有输入框）。
        """
        if SERVER_MODE:
            self._log("ℹ️ 服务器模式下没有目录选择框，请直接在路径框里填票据文件夹"
                      "（容器内路径，例如 /票据）")
            return {"server": True,
                    "hint": "服务器模式下弹不出目录选择框，请直接在路径框里填票据文件夹"
                            "（容器内路径，例如 /票据）"}
        import tkinter as tk
        from tkinter import filedialog
        self._log("打开文件夹选择框…")
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            path = filedialog.askdirectory(initialdir=initial or str(Path.home()),
                                           title="选择发票所在文件夹")
        finally:
            root.destroy()
        return path or ""

    def check_source(self, path: str = ""):
        """统计这个文件夹里有多少个可识别的凭证文件（发票 + 行程单 + 图片）"""
        p = Path(path or load_config().get("source", ""))
        if not p.exists() or not p.is_dir():
            return {"ok": False, "msg": "路径不存在或不是文件夹", "files": 0}
        import invparse
        files = invparse.collect_files(p, load_config().get("recursive", True))
        kinds = {}
        for f in files:
            kinds[f.suffix.lower()] = kinds.get(f.suffix.lower(), 0) + 1
        return {"ok": True, "files": len(files), "kinds": kinds,
                "msg": f"找到 {len(files)} 个凭证文件" if files else "没找到可识别的凭证文件"}

    def upload_info(self):
        """上传通道的状态：服务器模式才用得上（本机模式直接指路径就行，不用搬）"""
        base = upload_base()
        try:
            root = upload_root()
            ok, msg = True, ""
        except Exception as e:
            root, ok, msg = base / UPLOAD_SUBDIR, False, f"{type(e).__name__}: {e}"
        return {"server": SERVER_MODE, "root": str(root), "base": str(base),
                "ok": ok, "msg": msg, "max_size": MAX_UPLOAD_BYTES,
                "max_text": _human_size(MAX_UPLOAD_BYTES)}

    # ---------- 入库建账 ----------
    def start_import(self, recursive=None, category="", person="", batch="", source=""):
        """source = 界面上那个路径框里的目录。**必须传**，见 _import_worker 里的说明。

        返回 True/False 是界面早就依赖的契约（False ＝ 没接上活），所以这里保持不动，
        只是把真正干活的那一步交给任务层：抢锁、留痕、失败能重试。
        """
        if self.busy:
            return False
        payload = {"recursive": recursive, "category": category, "person": person,
                   "batch": batch, "source": source}
        res = self._job_call("start_import", self._import_job, kwargs=payload,
                             payload=payload)
        if isinstance(res, dict) and res.get("error"):
            self._log(f"⏳ {res['error']}")
            return False
        return True

    def _import_job(self, recursive=None, category="", person="", batch="", source=""):
        """入库任务的同步外壳：把 _import_worker 写回的结果转成任务结果。

        为什么要这层：_import_worker 是把结果写进 self._result 给界面看的，它自己不
        返回东西 —— 任务层就没法判断这次到底成功还是失败（失败会被记成「成功」，
        那「失败可重试」就成了摆设）。
        """
        self._import_worker(recursive, category, person, batch, source)
        r = self._result or {}
        if r.get("error"):
            return {"error": r["error"]}
        return dict(r)

    def _import_worker(self, recursive, category, person, batch, source=""):
        try:
            import invparse
            import ledger
            cfg = load_config()
            # ⚠️ 入库目录一律以「界面上填的那个」为准，不能用配置文件里的旧值兜底。
            #    曾经这里只读 cfg["source"]：用户把路径框改成另一个目录、点入库，
            #    实际扫的却还是老目录 —— 弹窗和日志都写着新路径，台账里进来的却是
            #    老目录那 80 多条，用户会以为"删了又全回来了 / 新的一个没入"。
            src = str(source or cfg.get("source", "")).strip().strip('"').strip()
            root = Path(src)
            if recursive is None:
                recursive = bool(cfg.get("recursive", True))
            if not root.exists():
                self._log(f"❌ 凭证文件夹不存在：{root}")
                self._set_status("文件夹不存在", "err")
                self._set_result({"error": f"文件夹不存在：{root}"})
                return
            # 用户明确指定了这个目录，就记进配置，免得「打开凭证文件夹」等地方对不上
            if src and not same_path(cfg.get("source", ""), src):
                cfg["source"] = src
                save_config(cfg)
                self._log(f"📌 已把凭证文件夹记为：{root}")
            self._set_status("读取文件…", "running")
            self._log(f"开始遍历：{root}" + ("（含子文件夹）" if recursive else "（只当前层）"))
            files = invparse.collect_files(root, recursive)
            if not files:
                self._log("ℹ️ 这个文件夹里没有 PDF / OFD / XML / zip 文件")
                self._set_status("没找到凭证文件", "ready")
                self._set_result({"found": 0})
                return
            self._log(f"找到 {len(files)} 个文件，开始解析…")

            cache = invparse.load_cache()
            cached_hit = 0

            def prog(i, total, name):
                self._set_progress(True, i, total, f"解析中… {name[:32]}")
                if i % 10 == 0 or i == total:
                    self._log(f"  解析进度 {i}/{total}")

            parsed = invparse.parse_all(files, cache, prog)
            invparse.save_cache(cache)
            recs = invparse.dedupe_same_invoice([x["rec"] for x in parsed])
            self._log(f"解析完成，得到 {len(recs)} 条凭证（同一发票号的多个文件已合并）")

            # 出行人认领：票面 → 手机号对照表 → 文件夹名，认不出就留空让人手填。
            # 顺手从「同文件夹 + 日期接近」的票面姓名里学「手机号 → 姓名」，
            # 下次同一部手机开的行程单就能自动认人（打车单票面没有姓名，只能这样）。
            import traveler as tv
            stat = tv.apply(recs)
            for rec in recs:
                if not rec.get("traveler"):
                    rec["warn"].append("没认出行人，可在台账里手填")
                elif rec.get("tv_from") == tv.SRC_FOLDER:
                    rec["warn"].append(f"出行人「{rec['traveler']}」按文件夹名推断，请核对")
                elif rec.get("tv_from") == tv.SRC_PHONE:
                    rec["warn"].append(f"出行人「{rec['traveler']}」由手机号对照表译出，请核对")
            self._log("👤 " + tv.summary_text(stat))
            for ph, nm in (stat.get("learned") or {}).items():
                self._log(f"   学到手机号：{ph} → {nm}（同一部手机开的行程单下次自动认人）")

            self._set_progress(True, 0, 0, "登记台账…")
            res = ledger.add_records(recs, category=category, person=person, batch=batch)
            added = len(res["added"])
            patched = int(res.get("patched") or 0)
            self._log(f"✅ 台账新增 {added} 条；已有 {len(res['in_ledger'])} 条；"
                      f"重复报销风险 {len(res['dup'])} 条；需人工 {len(res['manual'])} 条")
            if patched:
                self._log(f"🔧 顺带给 {patched} 条老记录补上了「凭证类型 / 行程明细」"
                          f"（老台账表头升级）")

            # 按凭证类型拆开报一遍，用户一眼能看出这批里几张发票、几张行程单
            by_type = {}
            for r in res["added"]:
                t = str(r.get("凭证类型") or "其他凭证")
                d = by_type.setdefault(t, {"n": 0, "sum": 0.0})
                d["n"] += 1
                d["sum"] += float(str(r.get("价税合计") or 0).replace(",", "") or 0)
            if by_type:
                self._log("　本次新增明细：" + "　".join(
                    f"{t} {v['n']}张/¥{v['sum']:,.2f}"
                    for t, v in sorted(by_type.items(), key=lambda kv: -kv[1]["n"])))
            for d in res["dup"]:
                self._log(f"  ⚠ 重复：{d['rec'].get('no') or d['rec'].get('ctype')} {d['tip'][:40]}")
            for m in res["manual"]:
                self._log(f"  ❓ 认不出的凭证：{Path(m.get('src','')).name[:52]}"
                          + (f"（{m['warn'][0][:40]}）" if m.get("warn") else ""))
            total_sum = sum(float(str(r.get("价税合计") or 0).replace(",", "") or 0)
                            for r in res["all"] if r.get("状态") != "已报销")
            self._log(f"台账当前 {len(res['all'])} 行，其中待报销合计 ¥{total_sum:,.2f}")
            self._set_progress(False)
            self._set_status(f"入库完成（新增 {added} 条）", "ok")
            self._set_result({"added": added, "in_ledger": len(res["in_ledger"]),
                            "dup": len(res["dup"]), "manual": len(res["manual"]),
                            "patched": patched, "source": str(root),
                            "total": len(res["all"]), "files": len(files),
                            "pending_amount": round(total_sum, 2),
                            "by_type": {t: v["n"] for t, v in
                                        sorted(by_type.items(), key=lambda kv: -kv[1]["n"])},
                            # 明细一起带回去，界面直接列表格，不用去猜日志文本
                            "manual_list": [{"name": Path(m.get("src", "")).name,
                                             "why": (m.get("warn") or [""])[0] or "没找到 20 位发票号码"}
                                            for m in res["manual"]],
                            "dup_list": [{"no": d["rec"].get("no") or d["rec"].get("ctype", ""),
                                          "tip": d["tip"]}
                                         for d in res["dup"]]})
        except ledger.LedgerBusy as e:
            self._log(f"❌ {e}")
            self._set_status("台账被占用", "err")
            self._set_result({"error": str(e)})
        except Exception as e:
            self._log(f"❌ 入库失败：{type(e).__name__}: {e}")
            self._set_status("入库失败", "err")
            self._set_result({"error": f"{type(e).__name__}: {e}"})
        finally:
            self._set_progress(False)
            # self.busy 交给任务层收（_job_release）：在这里提前放开的话，
            # 界面上按钮会先变成可点，其实锁还在人手上。

    # ---------- 台账 ----------
    def get_ledger(self, filt: dict = None):
        import ledger as lg
        try:
            allrows = lg.load()
        except Exception as e:
            return {"error": f"{type(e).__name__}: {e}", "rows": [], "total": 0}
        rows = filter_rows(lg.reindex(list(allrows)), filt)
        # 附件归位：打车行程单排到它对应发票的**紧后面**（用户要求「要跟在那个发票之后」）。
        # 只挪显示顺序，「序号」是行的身份，跟着行走、不重编号。
        att = attachment.analyze(rows)
        rows = att["rows"]
        # ⭐ 金额口径必须和生成的单据**一模一样**：打车行程单这类「发票的证明附件」
        #   金额已经含在它对应的那张发票里，再算一遍就是重复计钱。
        #   曾致台账页显示 4361.00、生成的单据只有 4284.90 —— 差的正是那张 76.10 的行程单。
        attach_seqs = att["attach_seqs"]
        m_bill = [self._amt(r) for r in rows if str(r.get("序号")) not in attach_seqs]
        m_att = [self._amt(r) for r in rows if str(r.get("序号")) in attach_seqs]
        pend = [self._amt(r) for r in rows
                if str(r.get("序号")) not in attach_seqs and r.get("状态") != "已报销"]
        # ⭐「分开报」：按「报销批次」分组，每组给张数/金额/起止日期/涉及的人，
        #   界面上点一下就只看这一批 —— 两批票就能各自独立统计、各自出单。
        #   分组时**忽略当前的批次筛选**，否则筛了某一批就只剩一个分组，反而没得挑。
        by_batch = self._by_batch(
            filter_rows(lg.reindex(list(allrows)),
                        {k: v for k, v in (filt or {}).items() if k != "batch"}),
            attach_seqs)
        return {
            "rows": rows,
            "total": len(rows),
            "sum": round(sum(m_bill), 2),                     # 计费合计（附件不重复计）
            "sum_raw": round(sum(m_bill) + sum(m_att), 2),    # 全部行简单相加（旧口径，留个对照）
            "attach_sum": round(sum(m_att), 2),               # 被排除掉的附件金额
            "pending_count": sum(1 for r in rows
                                 if str(r.get("序号")) not in attach_seqs
                                 and r.get("状态") != "已报销"),
            "pending_sum": round(sum(pend), 2),
            "all_count": len(allrows),
            "by_batch": by_batch,
            # 哪几行是「跟在发票后面的附件」→ 界面标「↳ 附件」并显示它是谁的附件
            "attach_links": att["link_of"],
            "attach_seqs": sorted(attach_seqs),
            "attach_count": att["n_attach"],
            # 「出行人核对」按钮上的角标：还空着的 + 提示里写着「请核对」的
            "tv_todo": sum(1 for r in allrows
                           if not str(r.get("出行人") or "").strip()
                           or "请核对" in str(r.get("提示") or "")),
            "categories": self._all_values("费用类别", rows=allrows),
            "ctypes": self._all_values("凭证类型", rows=allrows),
            "batches": self._all_values("报销批次", rows=allrows),
            "persons": self._all_values("报销人", rows=allrows),
            "months": self._all_values("开票日期", 7, rows=allrows),
            "aggregates": lg.aggregate(rows),
        }

    @staticmethod
    def _amt(r) -> float:
        """一行的价税合计（字符串带千分位/空值都能吃）"""
        return float(str(r.get("价税合计") or 0).replace(",", "") or 0)

    @staticmethod
    def _by_batch(rows, attach_seqs=None):
        """
        「按批次分开报」：把行按「报销批次」分组。
        每组给：张数 / 计费金额（附件不重复计）/ 附件张数 / 待报张数 / 已报张数 /
                开票日期区间 / 涉及的人（出行人，没有就用报销人）。
        次序：最近开票的批次在前，「没填批次」的排最后。
        """
        attach_seqs = {str(s) for s in (attach_seqs or ())}
        g = {}
        for r in rows:
            k = str(r.get("报销批次") or "").strip()
            d = g.get(k)
            if d is None:
                d = g[k] = {"batch": k, "n": 0, "n_attach": 0, "sum": 0.0,
                            "pending": 0, "done": 0, "first": "", "last": "",
                            "persons": []}
            d["n"] += 1
            if str(r.get("序号")) in attach_seqs:
                d["n_attach"] += 1
            else:
                d["sum"] += Api._amt(r)
            if r.get("状态") == "已报销":
                d["done"] += 1
            else:
                d["pending"] += 1
            who = str(r.get("出行人") or r.get("报销人") or "").strip()
            if who and who not in d["persons"]:
                d["persons"].append(who)
            dt = str(r.get("开票日期") or "").strip()[:10]
            if dt:
                d["first"] = dt if not d["first"] else min(d["first"], dt)
                d["last"] = dt if not d["last"] else max(d["last"], dt)
        # 最近开票的批次在前；没填日期（如手工录入的影印件）用全零占位，倒序后落到最后
        out = sorted(g.values(),
                     key=lambda x: (x["last"] or "0000-00-00", x["batch"]), reverse=True)
        for d in out:
            d["sum"] = round(d["sum"], 2)
        return out

    def _all_values(self, field, cut=0, rows=None):
        import ledger as lg
        if rows is None:
            rows = lg.load()
        vals = set()
        for r in rows:
            v = str(r.get(field, "") or "")
            if cut:
                v = v[:cut]
            if v:
                vals.add(v)
        return sorted(vals, reverse=True)

    def mark_rows(self, seqs, status="已报销", batch="", person=""):
        import ledger as lg
        fields = {"状态": status}
        if batch:
            fields["报销批次"] = batch
        elif status == "已报销":
            fields["报销批次"] = datetime.now().strftime("%Y-%m-%d") + " 报销单"
        if person:
            fields["报销人"] = person
        n = lg.update_rows(seqs or [], **fields)
        self._log(f"已把 {n} 张发票标记为「{status}」")
        return {"updated": n}

    def update_rows(self, seqs, fields: dict = None):
        import ledger as lg
        fields = {k: v for k, v in (fields or {}).items() if k in lg.FIELDS}
        n = lg.update_rows(seqs or [], **fields)
        self._log(f"已更新 {n} 张发票：" + "、".join(f"{k}={v}" for k, v in fields.items()))
        return {"updated": n}

    # ---------------------------------------------------------------- 出行人
    def get_phone_map(self):
        """
        「出行人手机号 → 姓名」对照表。打车行程单票面没有姓名、只有行程人手机号，
        靠这张表翻成人名（自动学 + 人工改）。
        """
        import traveler as tv
        import ledger as lg
        m = tv.load_map()
        # 台账里出现过的手机号也列出来（有的可能是后来加进台账、表里没有的）
        used = {}
        for r in lg.load():
            ph = tv.re.sub(r"\D", "", str(r.get("出行人") or ""))
            if len(ph) == 11:
                used.setdefault(ph, {"phone": ph, "name": "", "rows": 0})
                used[ph]["rows"] += 1
        rows = [{"phone": k, "name": v} for k, v in sorted(m.items())]
        known = {x["phone"] for x in rows}
        for k, v in used.items():
            if k not in known:
                rows.append({"phone": k, "name": "", "unsaved": True})
        return {"map": rows}

    def set_phone_name(self, phone, name=""):
        """记一条「这个手机号是谁」；name 给空 = 删掉这条"""
        import traveler as tv
        m = tv.set_name(phone, name)
        if str(name or "").strip():
            self._log(f"👤 已记住：手机号 {phone} → {name}")
        else:
            self._log(f"👤 已忘掉手机号 {phone} 的姓名对照")
        return {"map": [{"phone": k, "name": v} for k, v in sorted(m.items())]}

    # ---------------------------------------------------------------- 人工核对
    # 「三层兜底 + 人工核对」的后半截。
    # 自动认人终究是**猜**：手机号译出来的、文件夹名推出来的都可能错，票面读不到的
    # 干脆是空的。台账里只有一个「出行人」格子，看不出这个名字是票面白纸黑字写着的
    # （可信），还是手机号/文件夹猜的（得看一眼）。核对清单就是把每一行重新算一遍
    # 「现在能认成谁、依据是哪一层」，再跟台账里的值摆在一起让人当场改。
    @staticmethod
    def _strip_tv_tip(tip) -> str:
        """核对完，把「提示」里跟出行人有关的那句摘掉；重复报销、红字这些原样留着"""
        parts = [s.strip() for s in str(tip or "").split("；") if s.strip()]
        return "；".join(s for s in parts if "出行人" not in s)

    def traveler_review(self, seqs=None):
        import re as _re
        import invparse as ip
        import ledger as lg
        import traveler as tv
        try:
            rows = self._rows_by_seq(seqs) if seqs else lg.reindex(list(lg.load()))
        except Exception as e:
            return {"error": f"{type(e).__name__}: {e}"}

        cache = ip.load_cache()
        pm = tv.load_map()
        checked = tv.load_checked()
        budget = 400                 # 缓存里没有的最多现解析这么多张，别让界面干等
        out = []
        stat = {"total": 0, "need": 0, "done": 0, "blank": 0, "diff": 0,
                tv.SRC_FACE: 0, tv.SRC_FILENAME: 0, tv.SRC_PHONE: 0, tv.SRC_FOLDER: 0, "": 0}

        for r in rows:
            p = str(r.get("文件路径") or "")
            rec = None
            if p:
                try:
                    rec = cache.get(ip._ckey(Path(p)))       # 借解析缓存，命中就不用读文件
                except Exception:
                    rec = None
                if rec is None and budget > 0:
                    fp = Path(p)
                    if fp.exists():
                        budget -= 1
                        try:
                            rec = ip.parse_file(fp)
                        except Exception:
                            rec = None
            rec = dict(rec or {})
            if p:
                rec.setdefault("src", p)
            auto, src = tv.resolve(rec, phone_map=pm)
            cur = str(r.get("出行人") or "").strip()
            done = tv.is_checked(p, cur, checked)

            if not cur and not auto:
                state = "blank"          # 三层都没认出来 → 必须手填
            elif not cur:
                state = "todo"           # 认出来了，但台账这格还空着
            elif auto and cur != auto:
                state = "diff"           # 台账里的跟现在自动认的不一样 → 看一眼
            else:
                state = "ok"
            # 该不该人工看一眼：没认出来的必看；手机号/文件夹猜出来的要看；
            # 票面、文件名里白纸黑字写着姓名的，默认不必核（想全过一遍可以打开「显示全部」）
            weak = (state in ("blank", "todo", "diff")
                    or src in (tv.SRC_PHONE, tv.SRC_FOLDER))
            out.append({
                "seq": str(r.get("序号")), "ctype": r.get("凭证类型"),
                "no": r.get("发票号码"), "date": r.get("开票日期"),
                "total": r.get("价税合计"), "detail": r.get("行程/明细"),
                "person": r.get("报销人"), "cur": cur, "auto": auto, "src": src,
                "state": state, "need": bool((not done) and weak), "done": done,
                "file": p, "folder": tv.folder_of(p),
                "phone": _re.sub(r"\D", "", str(rec.get("phone") or "")),
            })
            stat["total"] += 1
            stat[src] = stat.get(src, 0) + 1
            if done:
                stat["done"] += 1
            if state == "blank":
                stat["blank"] += 1
            if state == "diff":
                stat["diff"] += 1
            if out[-1]["need"]:
                stat["need"] += 1

        # 要核对的排前面：认不出的 → 跟自动值打架的 → 文件夹推的 → 手机号译的 → 其他
        rank = {"blank": 0, "diff": 1}

        def _key(x):
            s = rank.get(x["state"])
            if s is None:
                s = 2 if x["src"] == tv.SRC_FOLDER else 3 if x["src"] == tv.SRC_PHONE else 4
            return (0 if x["need"] else 1, s, str(x["date"] or ""), str(x["seq"]))

        out.sort(key=_key)

        # 候选名字：对照表里的、台账里出现过的、配置里填的 —— 界面用 datalist 给补全
        names = set()
        for v in pm.values():
            if tv.is_name(v):
                names.add(v)
        try:
            for r in lg.load():
                for col in ("出行人", "报销人"):
                    v = str(r.get(col) or "").strip()
                    if tv.is_name(v):
                        names.add(v)
        except Exception:
            pass
        for v in _re.split(r"[、,，;；/\s]+", str(load_config().get("people") or "")):
            if tv.is_name(v):
                names.add(v)
        return {"rows": out, "stat": stat, "people": sorted(names)}

    def _face_names(self, seqs, by_seq):
        """
        查这几行「票面上白纸黑字写着的出行人」是谁 → {序号: 姓名}；票面没写的进不了这张表。

        为什么要它：批量「设置出行人」是勾一片、填一个名字，很容易把**票面本来读对**的
        姓名一起盖掉（邵自杰/陆兰玲 的高铁票曾被全写成 孙振强）。改之前先来这儿问一句。
        借解析缓存（命中就不读原件），缓存没有的当场解析；最多解析 200 张，免得界面干等。
        """
        import invparse as ip
        import traveler as tv
        out = {}
        try:
            cache = ip.load_cache()
        except Exception:
            cache = {}
        pm = tv.load_map()
        budget = 200
        for seq in seqs:
            r = by_seq.get(seq)
            if r is None:
                continue
            p = str(r.get("文件路径") or "").strip()
            if not p:
                continue
            rec = None
            try:
                rec = cache.get(ip._ckey(Path(p)))
            except Exception:
                rec = None
            if rec is None and budget > 0:
                fp = Path(p)
                if fp.exists():
                    budget -= 1
                    try:
                        rec = ip.parse_file(fp)
                    except Exception:
                        rec = None
            rec = dict(rec or {})
            rec.setdefault("src", p)
            name, src = tv.resolve(rec, phone_map=pm)
            if src == tv.SRC_FACE and name:
                out[seq] = name
        return out

    def save_traveler_review(self, items, remember=True, protect=True):
        """
        落下核对结果。items = [{"seq": 序号, "name": 出行人, "phone": 票面手机号}]
          · 写台账「出行人」（名字留空 = 清掉这格）；
          · 顺手把「提示」里那几句「请核对」摘掉，别的提示（重复、红字）留着；
          · remember=True 就把「手机号 → 姓名」记进对照表，下次同一部手机开的行程单自动认人；
          · 记一笔「这几张人工核过了」——名字以后再被改掉，核对标记自动失效。

        ⭐ protect=True（默认）：**票面上白纸黑字写着姓名的行，不许被别的名字盖掉**。
           批量「设置出行人」就是杀器 —— 勾一片、填一个名字，邵自杰/陆兰玲的票也写成了
           孙振强。这类行进 protected 名单、原样不动，让界面问一句「这几张也要覆盖吗」。
           用户在「出行人核对」面板里**逐行看着「来源」那一列**改的，传 protect=False。
        """
        import re as _re
        import ledger as lg
        import traveler as tv
        items = [x for x in (items or []) if isinstance(x, dict)]
        if not items:
            return {"updated": 0, "remembered": 0, "cleared": 0, "protected": []}
        try:
            by_seq = {str(r.get("序号")): r for r in lg.load()}
        except Exception as e:
            return {"error": f"{type(e).__name__}: {e}"}

        face = self._face_names([str(x.get("seq") or "").strip() for x in items], by_seq) \
            if protect else {}

        mapping, marks, learn, protected = {}, {}, {}, []
        for it in items:
            seq = str(it.get("seq") or "").strip()
            r = by_seq.get(seq)
            if r is None:
                continue
            name = str(it.get("name") or "").strip()
            fa = face.get(seq)
            if fa and name and name != fa:
                # 票面写着 fa，却要填成 name → 拦住，让界面来确认
                protected.append({"seq": seq, "face": fa, "new": name,
                                  "cur": str(r.get("出行人") or "").strip(),
                                  "no": str(r.get("发票号码") or ""),
                                  "detail": str(r.get("行程/明细") or ""),
                                  "file": str(r.get("文件路径") or "")})
                continue
            f = {"出行人": name}
            new_tip = self._strip_tv_tip(r.get("提示"))
            if new_tip != str(r.get("提示") or ""):
                f["提示"] = new_tip
            mapping[seq] = f
            p = str(r.get("文件路径") or "")
            if p:
                marks[p] = name                     # 空名字 = 撤掉这条核对标记
            ph = _re.sub(r"\D", "", str(it.get("phone") or ""))
            if remember and ph and name:
                learn[ph] = name

        try:
            n = lg.patch_rows(mapping)              # 只 load/save 一次，不会刷一屏备份
        except lg.LedgerBusy as e:
            return {"error": str(e)}
        tv.mark_checked(marks)
        if learn:
            m = tv.load_map()
            m.update(learn)
            tv.save_map(m)
        cleared = sum(1 for v in mapping.values() if not v["出行人"])
        self._log(f"👤 出行人核对完成：确认 {n} 张"
                  + (f"，其中 {cleared} 张留空" if cleared else "")
                  + (f"；记住 {len(learn)} 个手机号" if learn else ""))
        if protected:
            self._log(f"🛡️ 拦下 {len(protected)} 张：票面上写着别的人名，没被覆盖 → "
                      + "；".join(f"序号{d['seq']} 票面「{d['face']}」≠ 要填的「{d['new']}」"
                                  for d in protected[:8])
                      + ("…" if len(protected) > 8 else ""))
        # Excel 只是快照，写不出去（多半是 Excel 开着）不该让用户以为没改成 —— 说清楚
        stale = lg.excel_stale()
        if stale:
            self._log("⚠️ Excel 快照没刷新（文件被占用）：台账库已经改好了，"
                      "关掉 Excel 后随便再点一次修改就会自动补上")
        return {"updated": n, "remembered": len(learn), "cleared": cleared,
                "protected": protected, "excel_stale": stale}

    # 重新识别时只覆盖「票面上读来的」字段；人工填的一律不动
    _REFRESH_FIELDS = (("凭证类型", "ctype"), ("票面标记", "stamp"), ("发票号码", "no"),
                       ("开票日期", "date"), ("票种", "kind"), ("行程/明细", "detail"),
                       ("购方名称", "buyer"), ("销方名称", "seller"), ("不含税金额", "amount"),
                       ("税额", "tax"), ("价税合计", "total"), ("项目编号", "project_code"))

    def refresh_rows(self, seqs=None):
        """
        按**当前**的解析规则把选中的凭证重新读一遍原件，刷新票面字段。

        为什么要有它：解析规则是会改的（识别错了、票面版式变了），可台账里已经躺着的
        旧值是错的。删掉重新入库当然也行，但那样连「已报销状态 / 报销批次」一起丢。
        这里只覆盖票面上读来的字段，状态、费用类别、报销人、报销批次、入库时间都不动。
        """
        import ledger as lg
        import invparse as ip
        import traveler as tv
        rows = self._rows_by_seq(seqs)
        if not rows:
            return {"error": "没有选中的凭证"}
        patch, miss, fail = {}, [], []
        items = []                      # [(台账行, 新解析结果)]
        self._set_progress(True, 0, len(rows), "重新识别…")
        self._log(f"重新识别 {len(rows)} 张凭证的原件…")
        try:
            for i, r in enumerate(rows, start=1):
                self._set_progress(True, i, len(rows), f"重新识别 {i}/{len(rows)}")
                seq = str(r.get("序号"))
                p = Path(str(r.get("文件路径") or ""))
                if not p.exists():
                    miss.append({"row": seq, "file": p.name or "（台账里没记路径）"})
                    continue
                try:
                    rec = ip.parse_file(p)
                except Exception as e:
                    fail.append({"row": seq, "file": p.name,
                                 "why": f"{type(e).__name__}: {e}"})
                    continue
                items.append((r, rec))

            # 先整批认一遍出行人（顺带学「手机号 → 姓名」，选中的这批就是学习素材，
            # 这样老台账直接点「重新识别」也能把打车单认出来）
            if items:
                stat = tv.apply([rec for _, rec in items])
                self._log("👤 " + tv.summary_text(stat))
                for ph, nm in (stat.get("learned") or {}).items():
                    self._log(f"   学到手机号：{ph} → {nm}")

            for r, rec in items:
                seq = str(r.get("序号"))
                f = {}
                for col, key in self._REFRESH_FIELDS:
                    v = rec.get(key)
                    # 读不出来（空）就保留台账原值，免得把好数据冲成空白
                    if v not in (None, "", []):
                        f[col] = v
                # 出行人：只往「空着的」或「原来不是人名（比如以前误存成手机号）」的格子里填。
                # 旧值已经是正常姓名 → 可能是用户手填的，保留不动。
                name, src = str(rec.get("traveler") or ""), str(rec.get("tv_from") or "")
                old_tv = str(r.get("出行人") or "").strip()
                wrote_tv = False
                if name and (not old_tv or not tv.is_name(old_tv)):
                    f["出行人"] = name
                    wrote_tv = True
                if not name and not old_tv:
                    rec["warn"].append("没认出行人，可在台账里手填")
                elif wrote_tv and src == tv.SRC_FOLDER:
                    # 只在**确实写进去了**的时候提示，否则提示里的名字跟格子里留着的老值对不上
                    rec["warn"].append(f"出行人「{name}」按文件夹名推断，请核对")
                elif wrote_tv and src == tv.SRC_PHONE:
                    rec["warn"].append(f"出行人「{name}」由手机号对照表译出，请核对")
                # 项目/事由：旧值「空着」或「明显是误取的明细行（含 *）」时才动 ——
                # 含 * 的那种即使新值是空的也要覆盖（把乱码清掉），
                # 而用户手写过的（不含 *）一律保留
                old_proj = str(r.get("项目/事由") or "").strip()
                if not old_proj or "*" in old_proj:
                    f["项目/事由"] = rec.get("project", "")
                tip = "；".join([*rec.get("warn", []),
                                 *(["红字发票（负数）"] if rec.get("is_red") else [])])
                if tip:
                    f["提示"] = tip
                patch[seq] = f
            n = lg.patch_rows(patch) if patch else 0
            self._log(f"✅ 重新识别完成：更新 {n} 条"
                      + (f"，{len(miss)} 条原件找不到" if miss else "")
                      + (f"，{len(fail)} 条读不了" if fail else ""))
            for m in miss:
                self._log(f"    · 原件不在：{m['file']}")
            for x in fail:
                self._log(f"    · 读不了：{x['file']} —— {x['why']}")
            self._set_status("重新识别完成", "ok" if not (miss or fail) else "warn")
            return {"updated": n, "miss": len(miss), "fail": len(fail),
                    "miss_list": miss, "fail_list": fail}
        finally:
            self._set_progress(False)

    def delete_rows(self, seqs=None):
        """删掉台账里选中的行（删前会自动备份一张台账到「台账备份」）"""
        import ledger as lg
        n = 0
        try:
            n = lg.delete_rows(seqs or [])
        except lg.LedgerBusy as e:
            self._log(f"❌ {e}")
            return {"error": str(e)}
        if n:
            self._log(f"🗑 已从台账删除 {n} 条记录（删除前已自动备份到「台账备份」）")
        return {"deleted": n}

    def clear_ledger(self):
        """清空台账（表头保留；清空前后自动备份，误删能从「台账备份」恢复）"""
        import ledger as lg
        try:
            n = lg.clear_ledger()
        except lg.LedgerBusy as e:
            self._log(f"❌ {e}")
            return {"error": str(e)}
        if n:
            self._log(f"🗑 台账已清空（删掉 {n} 条，清空前已备份到「台账备份」）")
            self._set_status("台账已清空", "ready")
        return {"cleared": n}

    # ---------- 生成文件（历史 / 删除）----------
    def _record_output(self, path, kind="", extra: dict = None) -> dict:
        p = Path(path)
        item = {"path": str(p), "name": p.name, "kind": kind or "文件",
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "size": p.stat().st_size if p.exists() else 0}
        item.update(extra or {})
        items = [x for x in _load_history() if str(x.get("path")) != item["path"]]
        items.append(item)
        _save_history(items)
        self._last_output.append(str(p))
        # 同时写进数据库（P0-1 的 generated_files 表）：那个 JSON 历史文件是给界面看的，
        # 数据库这份才查得动「这张单子当时是用哪些票生成的」。
        try:
            db.gen_file_add(str(p), kind=item["kind"], size=item["size"],
                            snapshot=extra or None, name=item["name"])
        except Exception as e:                                  # noqa: BLE001
            _log_line(f"生成文件入库失败：{type(e).__name__}: {e}")
        return item

    def list_outputs(self):
        """生成过的文件（最新在前）；已经被手动删掉的会带 exists=False"""
        out = []
        for it in reversed(_load_history()):
            p = Path(str(it.get("path") or ""))
            d = dict(it)
            d["name"] = p.name or it.get("name", "")
            d["exists"] = p.exists()
            if p.exists():
                d["size"] = p.stat().st_size
            out.append(d)
        return {"items": out, "output": str(OUTPUT_DIR)}

    def delete_output(self, path=""):
        """删掉一个生成的文件（送回收站，不是永久删除）"""
        p = Path(str(path or ""))
        root = Path(OUTPUT_DIR).resolve()
        try:
            rp_ = p.resolve()
            inside = rp_ == root or root in rp_.parents
        except Exception:
            inside = False
        if not inside:
            return {"error": "只允许删除「输出」目录里的生成文件"}
        removed = False
        if p.exists():
            removed = _to_recycle_bin(p)
            if not removed:
                try:
                    p.unlink()
                    removed = True
                except OSError as e:
                    return {"error": f"删除失败：{e}"}
        _save_history([x for x in _load_history() if str(x.get("path")) != str(p)])
        self._log(f"🗑 已删除生成文件：{p.name}（已送回收站，需要的话可以去回收站还原）")
        return {"ok": True, "removed": removed, "name": p.name}

    # ---------- 报销单 ----------
    def _rows_by_seq(self, seqs):
        import ledger as lg
        want = {str(s) for s in (seqs or [])}
        return [r for r in lg.load() if str(r.get("序号")) in want]

    def report_persons(self, seqs=None):
        """
        模板二版式的「出差补助（按人）」块要用的数：按出行人把票据合计算出来。

        口径**刻意**跟生成的单据一致（复用 report_data：附件（打车行程单）不重复计钱），
        否则界面上算出一个数、打印出来又是另一个数，用户会以为哪儿出错了。
        """
        import report_pdf as rp
        rows = self._rows_by_seq(seqs)
        if not rows:
            return {"people": [], "total": 0.0, "n": 0, "attach": 0}
        rows = attachment.order_rows(sorted(rows, key=attachment.row_sort_key))
        d = rp.report_data(rows, {"kind": rp.JY_KIND})
        return {"people": [{"name": p["name"], "bills": p["bills"]} for p in d["by_person"]],
                "total": d["total"], "n": len(rows), "attach": d["n_attach"]}

    def make_report(self, meta: dict = None, seqs=None, fmt: str = "pdf"):
        """生成报销单。fmt: pdf / xlsx / both（用户可在界面上选，两种格式随便挑）"""
        import report_pdf as rp
        meta = dict(meta or {})
        rows = self._rows_by_seq(seqs)
        if not rows:
            return {"error": "没有选中的凭证"}
        # 附件归位：行程单排到对应发票的紧后面（report_data 里还会排一次，是幂等的）
        rows = attachment.order_rows(sorted(rows, key=attachment.row_sort_key))
        kind = meta.get("kind") or "费用报销单"
        fmt = (fmt or "pdf").lower()
        label = {"pdf": "PDF", "xlsx": "Excel", "both": "PDF + Excel"}.get(fmt, fmt)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = re.sub(r"[\\/:*?\"<>|]", "",
                      f"{kind}_{meta.get('person', '') or '未填'}_{stamp}")
        base = OUTPUT_DIR / stem
        try:
            self._set_progress(True, 0, 0, "生成报销单…")
            self._log(f"生成《{kind}》（{label}）：{len(rows)} 张凭证 → {stem}.*")
            # 金额口径（附件不重复计 + 差旅费补助）统一由 report_data 给，免得两边算得不一样
            d = rp.report_data(rows, meta)
            if d["n_attach"]:
                self._log(f"🚕 附件归位：{d['n_attach']} 张行程单已跟在对应发票后面，"
                          f"金额已含在发票里，不重复计入合计")
            if d["alw"]:
                self._log("💰 差旅费补助：" + "；".join(
                    f"{a['name']} {a['people']}人×{a['days']}天×{a['rate']:,.2f}"
                    for a in d["alw"]) + f" = ¥{d['alw_total']:,.2f}")
            if d["by_person"]:            # 模板二版式：按人汇总（票据＋这个人的补助）
                self._log("👥 按人汇总：" + "；".join(
                    f"{p['name']} {p['bills']:,.2f}"
                    + (f"＋补助{p['alw']:,.2f}＝{p['real']:,.2f}" if p["alw"] else "")
                    for p in d["by_person"]))
            outs = rp.make_report(rows, meta, base, fmt=fmt)
            info = {"count": len(rows), "sum": d["grand"], "upper": d["grand_upper"],
                    "bill": d["total"], "alw": d["alw_total"], "attach": d["n_attach"]}
            for p in outs:
                self._log(f"✅ 已生成 {p.name}（报销合计 ¥{d['grand']:,.2f} = {info['upper']}）")
                self._record_output(p, kind, info)
            self._set_status("报销单已生成", "ok")
            return {"out": str(outs[0]), "outs": [str(p) for p in outs],
                    "name": outs[0].name, "fmt": fmt, **info}
        except Exception as e:
            self._log(f"❌ 生成报销单失败：{type(e).__name__}: {e}")
            self._set_status("生成失败", "err")
            return {"error": f"{type(e).__name__}: {e}"}
        finally:
            self._set_progress(False)

    def _resolve_pdf(self, row):
        """
        找出这张发票对应的 PDF 文件，找不到就返回 None。
        优先用记录里的文件；如果记的是 zip/xml/ofd（XML/OFD 原件），就到同目录找一个
        文件名里带同一发票号码的 PDF —— 归档目录里 PDF 和 zip 通常是成对出现的。
        """
        return self._find_pdf(row)[0]

    def _find_pdf(self, row):
        """
        找这张发票能拿去合并的 PDF，返回 (路径 或 None, 找不到的原因)。

        找的顺序（前一步命中就不往后走）：
          1. 台账里记的那个文件本身就是 PDF → 直接用
          2. 同目录里文件名带发票号码的 PDF
          3. 同目录里的同名不同后缀 PDF
          4. 上级目录里文件名带发票号码的 PDF（OFD 常放在以号码命名的子文件夹里）
          5. 【按内容认票】同目录 / 上级目录里的 PDF 逐个读票面文字，谁的票面里有这个
             发票号码就是它 —— 手存成「张三-汉口-郑州东.pdf」这种名字全靠这一步
          6. 输出目录里带号码的 PDF（以前生成的合并件等）
        """
        import invparse
        p = Path(str(row.get("文件路径") or ""))
        no = str(row.get("发票号码") or "")
        suffix = p.suffix.lower()
        if suffix == ".pdf" and p.exists():
            return p, ""
        # OFD 里常带同名的 PDF 一起放在同一个文件夹，先按文件名找
        folders, seen = [], set()
        for folder in (p.parent, p.parent.parent):
            if folder and folder.is_dir() and str(folder).lower() not in seen:
                seen.add(str(folder).lower())
                folders.append(folder)
        if no:
            for folder in folders:
                try:
                    cands = list(folder.glob("*.pdf"))
                except OSError:
                    cands = []
                for cand in cands:
                    if cand.name.startswith(("发票合并", "发票台账汇总", "_")):
                        continue
                    if no in cand.name or (p.stem and p.stem.replace(".", "") in cand.name):
                        return cand, ""
                for cand in cands:
                    if cand.stem == p.stem:
                        return cand, ""
            # 文件名对不上的（「张三-汉口-郑州东.pdf」这种），打开票面按发票号认
            for folder in folders:
                hit = _folder_pdf_index(folder).get(no)
                if hit:
                    return hit, ""
        for cand in OUTPUT_DIR.glob(f"*{no}*.pdf") if no else []:
            return cand, ""
        # 到这儿说明真没有 PDF
        if suffix == ".ofd":
            why = ("只有 OFD 版式文件，附近也没找到同一张票的 PDF"
                   "（把 PDF 版式放到同一层文件夹里，或改用「按内容认票」能扫到的位置）")
        elif suffix in (".xml", ".zip"):
            why = "只有 XML 原始数据，没有 PDF 版式文件（可在税局重新下载 PDF 版式）"
        elif suffix in (".jpg", ".jpeg", ".png", ".bmp", ".webp"):
            why = "是图片凭证，没有 PDF（图片可以自己插进文档里打印）"
        elif not p.exists():
            why = f"台账里记的原始文件已经不在了：{p.name or '（没记路径）'}"
        else:
            why = "找不到对应的 PDF 文件"
        return None, why

    def merge_invoices(self, seqs=None, mode="2up"):
        import report_pdf as rp
        rows = self._rows_by_seq(seqs)
        if not rows:
            return {"error": "没有选中的凭证"}
        # 附件归位：打车行程单（证明发票明细的附件）一律排到它对应发票的**紧后面**，
        # 并且在 2up 版式下尽量和发票落在同一张 A4 的上下两半（用户要求「要在一块」）。
        rows = attachment.order_rows(sorted(rows, key=attachment.row_sort_key))
        pairs = attachment.analyze(rows)["pairs"]        # [(附件行, 主票行, 金额)]
        att_of_main = {id(m): a for a, m, _amt in pairs}
        pdf_of = {}
        missing = []
        for r in rows:
            p, why = self._find_pdf(r)
            pdf_of[id(r)] = p
            if not p:
                missing.append({
                    "name": Path(str(r.get("文件路径") or "")).name or "（无文件）",
                    "no": str(r.get("发票号码") or "") or "（无号码）",
                    "why": why,
                    "row": str(r.get("序号") or ""),
                })
        # 分组：一个单元 = 一张发票 + 它的附件，单元内不会被裁切到两张纸上
        units, used = [], set()
        for r in rows:
            if id(r) in used:
                continue
            used.add(id(r))
            pdf = pdf_of.get(id(r))
            buddy = att_of_main.get(id(r))
            if buddy is not None and id(buddy) not in used and pdf_of.get(id(buddy)):
                used.add(id(buddy))
                units.append([x for x in (pdf, pdf_of.get(id(buddy))) if x])
            elif pdf:
                units.append([pdf])
        paths = [p for u in units for p in u]
        linked = sum(1 for u in units if len(u) > 1)
        if not paths:
            self._log("❌ 选中的凭证都找不到 PDF 版式文件，没法合并：")
            for m in missing[:8]:
                self._log(f"    · {m['name']} —— {m['why']}")
            return {"error": "找不到可合并的 PDF 版式文件",
                    "missing": len(missing), "missing_list": missing}
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        tag = "上下两张" if mode == "2up" else "每页一张"
        out = OUTPUT_DIR / f"发票合并_{tag}_{stamp}.pdf"
        try:
            self._set_progress(True, 0, 0, "合并 PDF…")
            self._log(f"合并 {len(paths)} 张发票 PDF（{tag}）…")
            if linked:
                where = "同一张 A4 的上下两半" if mode == "2up" else "相邻两页"
                self._log(f"🚕 附件归位：{linked} 组「发票 + 行程单」已排在一起（{where}）")
            info = rp.merge_pdfs(paths, out, mode=mode, units=units)
            self._log(f"✅ 已生成 {out.name}：{info['pages']} 页 A4")
            if info["skipped"]:
                for name, why in info["skipped"]:
                    self._log(f"  ⚠ 跳过 {name[:40]}：{why[:60]}")
            if missing:
                self._log(f"  ⚠ {len(missing)} 张没有可合并的 PDF，没进这份文件：")
                for m in missing[:8]:
                    self._log(f"      · {m['name']} —— {m['why']}")
            self._record_output(out, "发票合并",
                                {"pages": info["pages"], "count": info["files"],
                                 "missing": len(missing), "linked": linked})
            if missing:
                self._set_status(f"合并完成（少 {len(missing)} 张）", "warn")
            else:
                self._set_status("合并完成", "ok")
            return {"out": str(out), "name": out.name, "pages": info["pages"],
                    "files": info["files"], "missing": len(missing), "linked": linked,
                    "missing_list": missing, "asked": len(rows),
                    "skipped": [{"name": n, "why": w} for n, w in info["skipped"]]}
        except Exception as e:
            self._log(f"❌ 合并失败：{type(e).__name__}: {e}")
            self._set_status("合并失败", "err")
            return {"error": f"{type(e).__name__}: {e}"}
        finally:
            self._set_progress(False)

    # ---------- 统计 ----------
    def export_summary(self, filt: dict = None):
        import ledger as lg
        try:
            rows = filter_rows(lg.load(), filt)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out = OUTPUT_DIR / f"发票台账汇总_{stamp}.xlsx"
            lg.export_summary(out, rows)
            self._log(f"✅ 已导出统计：{out.name}（{len(rows)} 行）")
            self._record_output(out, "统计 Excel", {"count": len(rows)})
            return {"out": str(out), "name": out.name, "rows": len(rows)}
        except Exception as e:
            self._log(f"❌ 导出统计失败：{type(e).__name__}: {e}")
            return {"error": f"{type(e).__name__}: {e}"}

    # ---------- 数据层 / 审计 ----------
    def ledger_db_info(self):
        """台账数据库的健康状况（设置页显示，也可以当健康检查用）"""
        try:
            st = db.stats()
            lf = Path(LEDGER_FILE)
            st.update({
                "excel": str(lf), "excel_exists": lf.exists(),
                "migrated_at": db.meta_get("migrated_from_excel"),
                "unique_index_ok": not db.meta_get("unique_index_inv_no_failed")
                                   and not db.meta_get("unique_index_fp_failed"),
            })
            return {"ok": True, **st}
        except Exception as e:                                  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}

    def import_from_excel(self, path=""):
        """把一份台账 Excel 导回数据库（整表替换，导之前自动备份）。

        这是「Excel 从写入源降级成导入源」的另一半：在 Excel 里改完想并回系统就走这里。
        因为会整表覆盖，界面上要二次确认。
        """
        src = Path(path) if path else Path(LEDGER_FILE)
        if not src.exists():
            return {"error": f"找不到 Excel：{src}"}
        try:
            before = db.count()
            res = db.migrate_from_excel(src, force=True)
            self._log(f"📥 已从 {src.name} 导入 {res.get('migrated', 0)} 行"
                      f"（导入前 {before} 行；原文件备份为 {Path(res.get('backup') or '').name}）")
            return {"ok": True, "before": before, **res}
        except Exception as e:                                  # noqa: BLE001
            self._log(f"❌ 从 Excel 导入失败：{type(e).__name__}: {e}")
            return {"error": f"{type(e).__name__}: {e}"}

    def audit_list(self, limit=200, action="", username="", offset=0):
        """审计日志：谁、什么时候、从哪来、改了什么（改前改后都有）"""
        try:
            return {"rows": db.list_audit(limit=limit, action=action,
                                          username=username, offset=offset),
                    "total": db.count_audit()}
        except Exception as e:                                  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}

    def list_jobs(self, limit=50):
        """后台任务列表（排队/运行中/成功/失败）+ 此刻占着重活锁的那个任务。

        界面「任务」页就是照这个画的：谁在跑、跑到哪、失败的能不能重试。
        """
        try:
            rows = db.job_list(limit=limit)
            today = datetime.now().strftime("%Y-%m-%d")
            counts = {"running": 0, "queued": 0, "failed": 0, "done_today": 0}
            for r in rows:
                r["kind_cn"] = JOB_KINDS.get(r.get("kind") or "", r.get("kind") or "")
                r["state"] = r.get("state") or ""
                r["secs"] = round(_elapsed(r), 1)
                r["retryable"] = bool(r.get("kind") in HEAVY_JOBS
                                      and r["state"] == db.JOB_FAILED
                                      and (r.get("payload") or ""))
                r.pop("payload", None)       # 参数里可能有路径，不必发给界面
                r.pop("result_json", None)
                if r["state"] == db.JOB_RUNNING:
                    counts["running"] += 1
                elif r["state"] == db.JOB_QUEUED:
                    counts["queued"] += 1
                elif r["state"] == db.JOB_FAILED:
                    counts["failed"] += 1
                elif r["state"] == db.JOB_DONE and str(r.get("created_at", "")).startswith(today):
                    counts["done_today"] += 1
            cur = self._heavy or {}
            return {"rows": rows, "counts": counts,
                    "busy": bool(cur), "busy_label": cur.get("label", ""),
                    "busy_user": cur.get("user", ""),
                    "kinds": JOB_KINDS}
        except Exception as e:                                  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}

    def gen_files(self, limit=200):
        """生成过的报销单 / 合并文件（带快照，方便追溯某张单子用了哪些票）"""
        try:
            rows = db.gen_file_list(limit=limit)
            for r in rows:
                r.pop("snapshot_json", None)
            return {"rows": rows}
        except Exception as e:                                  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}

    # ---------- 用户与权限（只有管理员能调；ROLE_ALLOW 里只有 admin 有） ----------
    def whoami(self):
        """我是谁、什么角色、哪些事能办 —— 界面靠它决定显示哪些入口。

        ⚠️ 藏起来只是「不显眼」，不是「不能做」：真正的卡口在 do_POST 里，
           每个接口都会再查一次权限。两处都要有。
        """
        me, r = db.actor(), db.role() or "admin"
        return {"user": me, "role": r, "role_cn": db.ROLES.get(r, r),
                # auth=False 表示本机模式（没设访问口令，压根没有「登录」这回事）——
                # 界面靠它决定要不要显示「当前登录 / 退出登录」
                "auth": _auth_on(),
                "can": {k: db.can(me, r, k) for k in
                        ("clear_ledger", "delete_rows", "set_config", "import_from_excel",
                         "make_report", "start_import", "list_users", "get_ledger")},
                # 完整白名单：界面据此灰掉点不了的按钮（见 _allowed_methods）
                "allowed": _allowed_methods(r)}

    def list_users(self):
        try:
            rows = []
            for u in db.user_list():
                u.pop("pw_hash", None)          # 口令哈希绝不下发给浏览器（哪怕是自己人）
                u["role_cn"] = db.ROLES.get(u["role"], u["role"])
                rows.append(u)
            me = db.actor()
            my_role = ""
            for u in rows:
                if u["username"] == me:
                    my_role = u["role"]
            return {"rows": rows, "roles": db.ROLES, "me": me, "my_role": my_role}
        except Exception as e:                                  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}

    def create_user(self, username="", password="", role="biz",
                    display_name="", dept=""):
        username = str(username or "").strip()
        if not username or not password:
            return {"error": "用户名和密码都要填"}
        if len(str(password)) < 6:
            return {"error": "密码至少 6 位"}
        if db.user_get(username):
            return {"error": f"用户名「{username}」已经有人用了"}
        try:
            db.user_create(username, password, role=role,
                           display_name=display_name, dept=dept)
        except ValueError as e:
            return {"error": str(e)}
        except Exception as e:                                  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log(f"👤 新增用户：{username}（{db.ROLES.get(role, role)}）")
        db.audit("新增用户", "user", username, after={"role": role, "dept": dept})
        return {"ok": True, "rows": self.list_users().get("rows", [])}

    def set_user_password(self, username="", password=""):
        if not db.user_get(username):
            return {"error": f"没有这个用户：{username}"}
        if len(str(password or "")) < 6:
            return {"error": "密码至少 6 位"}
        try:
            db.user_set_password(username, password)
        except Exception as e:                                  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        n = _drop_sessions_of(username)     # 改完密码必须把旧会话踢掉，否则旧 Cookie 照用
        self._log(f"👤 已重置 {username} 的密码，踢掉 {n} 个会话")
        db.audit("重置密码", "user", username, after={"kicked": n})
        return {"ok": True, "kicked": n}

    def set_user_role(self, username="", role=""):
        if not db.user_get(username):
            return {"error": f"没有这个用户：{username}"}
        try:
            db.user_set_role(username, role)
        except ValueError as e:
            return {"error": str(e)}
        db.audit("改角色", "user", username, after={"role": role})
        self._log(f"👤 {username} 的角色改成「{db.ROLES.get(role, role)}」")
        return {"ok": True, "rows": self.list_users().get("rows", [])}

    def set_user_disabled(self, username="", disabled=True):
        if not db.user_get(username):
            return {"error": f"没有这个用户：{username}"}
        if username == "admin" and disabled:
            return {"error": "admin 是内置管理员，不能停用（否则谁都管不了用户了）"}
        db.user_set_disabled(username, bool(disabled))
        n = _drop_sessions_of(username) if disabled else 0
        db.audit("停用用户" if disabled else "启用用户", "user", username)
        self._log(f"👤 {username} 已{'停用' if disabled else '启用'}"
                  + (f"，踢掉 {n} 个会话" if n else ""))
        return {"ok": True, "kicked": n, "rows": self.list_users().get("rows", [])}

    def force_logout(self, username=""):
        n = _drop_sessions_of(username)
        db.audit("强制退出", "user", username, after={"kicked": n})
        self._log(f"👤 已让 {username} 退出（{n} 个会话）")
        return {"ok": True, "kicked": n}

    # ---------- 打开位置 ----------
    def open_folder(self, which="output"):
        places = {"output": OUTPUT_DIR, "ledger": LEDGER_FILE.parent,
                  "app": APP_DIR, "source": Path(load_config().get("source", APP_DIR))}
        if SERVER_MODE:
            # 容器里 /app 是只读镜像，里面没什么好看的；日志和数据都落在数据目录
            places["app"] = DATA_DIR
        target = Path(places.get(which, OUTPUT_DIR))
        if target.is_file():
            target = target.parent
        if not target.exists():
            return {"error": f"目录不存在：{target}"}
        if SERVER_MODE:
            # 容器里没有资源管理器 → 改成浏览器里列目录（前端 bridge 认 url 字段）
            url = _url_of(target)
            self._log(f"🌐 服务器模式：改为在浏览器里查看目录 {target}")
            return {"ok": True, "server": True, "path": str(target), "url": url,
                    "msg": "服务器模式没法打开系统文件夹，已改为在浏览器里查看"}
        os.startfile(str(target))
        self._log(f"已打开目录：{target}")
        return {"ok": True, "path": str(target)}

    def open_file(self, path=""):
        p = Path(path)
        if not p.exists():
            return {"error": f"文件不存在：{path}"}
        if SERVER_MODE:
            url = _url_of(p)
            if not url:
                return {"error": "服务器模式下只能打开「输出」「数据」「票据」目录里的文件"}
            self._log(f"🌐 服务器模式：改为浏览器打开 {p.name}")
            return {"ok": True, "server": True, "url": url, "name": p.name}
        os.startfile(str(p))
        return {"ok": True}

    def reveal(self, path=""):
        """在资源管理器里定位并选中文件（服务器模式没有这个概念 → 退化成下载）"""
        p = Path(path)
        if not p.exists():
            return {"error": f"文件不存在：{path}"}
        if SERVER_MODE:
            url = _url_of(p)
            if not url:
                return {"error": "服务器模式下只能打开「输出」「数据」「票据」目录里的文件"}
            self._log(f"🌐 服务器模式：没有「在文件夹中定位」这一步，已改为下载 {p.name}")
            return {"ok": True, "server": True, "url": url, "name": p.name}
        subprocess.Popen(["explorer", "/select,", str(p)])
        return {"ok": True}

    def shutdown(self):
        self._log("收到退出请求，正在关闭…")
        self.exit_flag.set()
        return True


# ======================================================================
# 访问口令
# ----------------------------------------------------------------------
# 只有设了口令才启用。本机双击 bat 跑的时候不设 —— 行为跟以前一模一样
# （打开就能用、不弹登录）。容器放到 NAS 上才需要，见 _cur_password()。
# ======================================================================
COOKIE_NAME = "fb_auth"

LOGIN_PAGE = """<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>发票报销工具 · 登录</title><style>
body{margin:0;height:100vh;display:flex;align-items:center;justify-content:center;
 background:#f4f5f7;font:14px/1.6 "Microsoft YaHei",system-ui,-apple-system,sans-serif;color:#1f2328}
.card{background:#fff;padding:26px 28px;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,.09);width:300px}
h1{font-size:17px;margin:0 0 4px}p.sub{margin:0 0 18px;color:#6b7280;font-size:12px}
input{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #d0d7de;border-radius:8px;font-size:14px;outline:none;margin-bottom:8px}
input:focus{border-color:#c0392b;box-shadow:0 0 0 3px rgba(192,57,43,.12)}
button{margin-top:14px;width:100%;padding:10px;border:0;border-radius:8px;background:#c0392b;
 color:#fff;font-size:14px;cursor:pointer}button:hover{background:#a93226}
.err{margin-top:10px;color:#c0392b;font-size:12px;min-height:16px}
</style></head><body>
<form class="card" method="post" action="/login">
  <h1>发票报销工具</h1>
  <p class="sub">有账号就填账号密码；只填访问口令＝管理员（用户名留空）</p>
  <input type="text" name="username" autocomplete="username" placeholder="用户名（留空＝用访问口令登录，是管理员）">
  <input type="password" name="password" autofocus autocomplete="current-password" placeholder="密码 / 访问口令">
  <button type="submit">进入</button>
  <div class="err">{err}</div>
</form></body></html>"""


def _login_page(err: str = "") -> str:
    return LOGIN_PAGE.replace("{err}", err)


def _cur_password() -> str:
    """当前口令：环境变量优先，其次配置文件里的 web_password（设置页可改）"""
    if PASSWORD:
        return PASSWORD
    return str(load_config().get("web_password") or "").strip()


def _auth_on() -> bool:
    return bool(_cur_password())


# 已登录的会话：令牌 → 是谁。只放内存 —— 容器重启后要重新登录（不用管过期）。
_SESSIONS = {}


def _new_session(user: str, role: str, ip: str = "") -> str:
    import secrets
    tok = secrets.token_hex(16)
    if len(_SESSIONS) > 500:                 # 自用场景，别无限涨
        _SESSIONS.clear()
    _SESSIONS[tok] = {"user": user or "", "role": role or "admin",
                      "ip": ip or "", "ts": time.time()}
    return tok


def _session_of(token: str) -> dict:
    return _SESSIONS.get(token or "") or {}


def _drop_sessions_of(user: str) -> int:
    """把某个人的会话全部踢掉（改密码 / 停用 / 强制退出时用）。

    ⚠️ 必须服务端踢：只让浏览器丢 Cookie 等于没退出（老 Cookie 还能继续用）。
    """
    gone = [t for t, s in _SESSIONS.items() if s.get("user") == user]
    for t in gone:
        _SESSIONS.pop(t, None)
    return len(gone)


def _human_size(n) -> str:
    n = float(n or 0)
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:,.0f} {unit}" if unit == "B" else f"{n:,.1f} {unit}"
        n /= 1024
    return f"{n:,.1f} GB"


# 这些后缀让浏览器直接看（PDF 内置阅读器、图片直接显示）；其余一律当下载
_INLINE_EXT = {".pdf", ".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".txt", ".log"}


def _elapsed(job: dict) -> float:
    """任务耗时（秒）。正在跑的就算到此刻，没开始的记 0。"""
    def _t(s):
        try:
            return datetime.strptime(str(s or ""), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
    a = _t(job.get("started_at")) or _t(job.get("created_at"))
    if not a:
        return 0.0
    b = _t(job.get("ended_at")) or datetime.now()
    return max(0.0, (b - a).total_seconds())


def _dir_page(title: str, rows: list, up: str = "") -> str:
    """给 /files/ 用的极简目录列表（rows = [(显示名, href, 备注)]）"""
    from urllib.parse import quote
    items = []
    if up:
        items.append(f'<li><a class="up" href="{up}">↩ 返回上一级</a></li>')
    for name, href, note in rows:
        items.append(f'<li><a href="{href}">{name}</a><span class="n">{note}</span></li>')
    return f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{title}</title><style>
body{{margin:0;padding:22px 26px;background:#f4f5f7;
 font:14px/1.7 "Microsoft YaHei",system-ui,-apple-system,sans-serif;color:#1f2328}}
h1{{font-size:15px;margin:0 0 12px;color:#57606a;font-weight:600;word-break:break-all}}
ul{{list-style:none;margin:0;padding:0;max-width:900px}}
li{{display:flex;justify-content:space-between;gap:12px;padding:8px 12px;background:#fff;
 border:1px solid #e6e8eb;border-radius:8px;margin-bottom:6px}}
a{{color:#1f2328;text-decoration:none;word-break:break-all}}a:hover{{color:#c0392b}}
a.up{{color:#57606a}}.n{{color:#8c959f;font-size:12px;white-space:nowrap}}
.empty{{color:#8c959f}}
</style></head><body><h1>{title}</h1><ul>{''.join(items) or '<li class="empty">（空目录）</li>'}</ul></body></html>"""


# ======================================================================
# 本地 HTTP 服务
# ======================================================================
class UiHandler(http.server.SimpleHTTPRequestHandler):
    api = None
    # 高频轮询接口不逐条记日志：poll 每 0.4s 一次、list_jobs 1.5s 一次，
    # 记了会把「运行记录」和 docker logs 刷满，真正要看的东西反而被冲走。
    QUIET_API = {"poll", "list_jobs"}

    def log_message(self, *args):
        pass

    # ---------- 登录相关 ----------
    def _cookie_token(self) -> str:
        for part in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = part.strip().partition("=")
            if k == COOKIE_NAME:
                return v
        return ""

    def _role(self) -> str:
        """当前会话的角色。免登录（本机模式）就是管理员 —— 跟以前一样什么都能干。"""
        if not _auth_on():
            return "admin"
        s = _session_of(self._cookie_token())
        return s.get("role") or ("admin" if s else "")

    def _username(self) -> str:
        """这次请求算谁做的（审计要用真名）"""
        if not _auth_on():
            return "本机"
        s = _session_of(self._cookie_token())
        if not s:
            return ""
        return s.get("user") or "口令用户"

    def _authed(self) -> bool:
        if not _auth_on():
            return True
        tok = self._cookie_token()
        return bool(tok) and tok in _SESSIONS

    def _deny(self):
        """没登录：页面请求跳登录页，接口请求返 401（前端 bridge 会当成错误抛出）"""
        if self.path.startswith("/api/"):
            _log_line(f"🚫 未登录就访问接口：{self.path} "
                      f"@{self.client_address[0] if self.client_address else '?'}", "WARN")
            body = json.dumps({"result": None, "error": "未登录：请先在页面里输入口令"},
                              ensure_ascii=False).encode("utf-8")
            self.send_response(401)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)
            return
        self.send_response(302)
        self.send_header("Location", "/login")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def _send_html(self, html: str, code: int = 200):
        body = html.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _try_login(self):
        """登录。

        两条路：
          ① 填了用户名 → 按数据库里的用户校验（多用户，能追责到人）；
          ② 没填用户名 / 库里还没这个账号 → 退回老的「访问口令」那条路，
             这样已经部署好的 NAS（只设了 FB_PASSWORD）不用改任何东西就能继续用。
        """
        import hmac
        from urllib.parse import parse_qs
        n = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(n).decode("utf-8", "replace") if n else ""
        form = parse_qs(raw)
        name = (form.get("username") or [""])[0].strip()
        pw = (form.get("password") or [""])[0]
        ip = self.client_address[0] if self.client_address else ""

        user = db.user_get(name) if name else {}
        if user:
            if user.get("disabled"):
                _log_line(f"🔑 登录被拒（账号已停用）：{name} @{ip}")
                db.audit("登录失败", "user", name, ok=False, error="账号已停用",
                         username=name, ip=ip)
                return self._send_html(_login_page("这个账号已被停用，找管理员开一下"))
            if db.verify_password(pw, user.get("pw_hash") or ""):
                return self._after_login(user["username"], user["role"], ip)
            _log_line(f"🔑 登录失败（密码不对）：{name} @{ip}")
            db.audit("登录失败", "user", name, ok=False, error="密码不对",
                     username=name, ip=ip)
            return self._send_html(_login_page("用户名或密码不对"))

        # 老路：访问口令。库里还没有任何账号时，顺手用这个口令把管理员建出来，
        # 免得用户部署完 NAS 却因为「没有账号」进不去。
        # ⚠️ compare_digest 不能拿含中文的 str 直接比（TypeError）→ 统一转成字节再比。
        #    否则用户在口令里打一个中文字，看到的就是「连接被重置」，最难查的那种。
        if _cur_password() and hmac.compare_digest(pw.encode("utf-8"),
                                                   _cur_password().encode("utf-8")):
            # 访问口令是部署者自己设的，等于「主人的钥匙」，等同于管理员。
            # 想区分到人、想限制权限，就给每个人开账号（设置页 → 用户管理）。
            if db.user_count() == 0:
                try:
                    db.user_create("admin", pw, role="admin", display_name="管理员")
                    _log_line("🔑 已用访问口令创建管理员账号：admin")
                except Exception as e:                          # noqa: BLE001
                    _log_line(f"建管理员账号失败：{e}")
                return self._after_login("admin", "admin", ip)
            return self._after_login("口令用户", "admin", ip)

        _log_line(f"🔑 登录失败 @{ip}")
        db.audit("登录失败", "user", name or "(口令)", ok=False, error="凭据不对",
                 username=name or "", ip=ip)
        self._send_html(_login_page("用户名或密码不对"))

    def _after_login(self, user: str, role: str, ip: str = ""):
        self.send_response(302)
        self.send_header("Location", "/index.html")
        self.send_header("Set-Cookie",
                         f"{COOKIE_NAME}={_new_session(user, role, ip)}; Path=/; "
                         f"Max-Age=2592000; HttpOnly; SameSite=Lax")
        self.send_header("Content-Length", "0")
        self.end_headers()
        try:
            db.user_touch_login(user)
        except Exception:                                       # noqa: BLE001
            pass
        db.audit("登录成功", "user", user, username=user, ip=ip)
        _log_line(f"🔑 登录成功：{user}（{db.ROLES.get(role, role)}） @{ip}")

    def _logout(self):
        _SESSIONS.pop(self._cookie_token(), None)     # 立刻失效，不只是让浏览器丢掉
        self.send_response(302)
        self.send_header("Location", "/login")
        self.send_header("Set-Cookie", f"{COOKIE_NAME}=; Path=/; Max-Age=0")
        self.send_header("Content-Length", "0")
        self.end_headers()

    # ---------- /files/ ：浏览器直接看容器里的文件 ----------
    def _serve_files(self, path: str):
        import mimetypes
        import shutil
        from urllib.parse import quote, unquote
        from urllib.parse import unquote as _uq

        parts = _uq(path[len("/files"):]).strip("/").split("/", 1)
        alias = parts[0] if parts and parts[0] else ""
        rel = parts[1] if len(parts) > 1 else ""
        roots = _fs_roots()

        if not alias:                        # /files/ → 有哪些可看的目录
            rows = []
            for a, root in roots.items():
                try:
                    n = len(list(Path(root).iterdir()))
                except OSError:
                    n = 0
                # ⚠️ 不再显示服务器上的绝对路径：容器里的目录结构对外没有意义，
                #    而真实路径本身就是一种信息泄露。
                rows.append((f"{a}/", f"/files/{a}/", f"{n} 项"))
            return self._send_html(_dir_page("可浏览的目录", rows))

        if alias not in roots:
            # ⚠️ send_error 的描述会拼进 HTTP 状态行，只能是 ASCII（中文会让
            #    wfile 用 latin-1 编码时报错、连接直接断掉，客户端看到的是「连接失败」）
            self.send_error(404, "unknown files alias")
            return
        if _fs_blocked(rel):                 # 配置 / 日志 / 缓存 / 备份 / 数据库：一律不给
            _log_line(f"🚫 拦下 /files/ 越权访问：{alias}/{rel}")
            self.send_error(403, "this path is not shared")
            return
        root = Path(roots[alias])
        target = _resolve_under(root, rel)
        if target is None:
            self.send_error(403, "path outside the shared roots")
            return

        if target.is_dir():
            try:
                entries = sorted(target.iterdir(), key=lambda p: (not p.is_dir(), p.name))
            except OSError as e:
                _log_line(f"列目录失败 {target}：{e}")
                self.send_error(500, "cannot list directory")
                return
            base_rel = target.relative_to(root).as_posix() if target != root else ""
            rows = []
            for p in entries:
                sub = f"{base_rel}/{p.name}".strip("/") if base_rel else p.name
                if _fs_blocked(sub):         # 列表里也别把它们列出来
                    continue
                note = "" if p.is_dir() else _human_size(p.stat().st_size if p.exists() else 0)
                rows.append((p.name + ("/" if p.is_dir() else ""),
                             f"/files/{alias}/{quote(sub)}" + ("/" if p.is_dir() else ""),
                             note))
            up = ""
            if target != root:
                parent = target.parent
                try:
                    if parent == root:
                        up = f"/files/{alias}/"
                    elif root in parent.parents:
                        up = f"/files/{alias}/{quote(parent.relative_to(root).as_posix())}/"
                except Exception:
                    up = ""
            title = f"/{base_rel}" if base_rel else f"{alias}"
            return self._send_html(_dir_page(title, rows, up))

        if not target.is_file():
            self.send_error(404, "not found")
            return

        ctype = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        size = target.stat().st_size
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(size))
        self.send_header("Cache-Control", "no-store")
        if target.suffix.lower() not in _INLINE_EXT:
            self.send_header("Content-Disposition",
                             "attachment; filename*=UTF-8''" + quote(target.name))
        self.end_headers()
        try:
            with open(target, "rb") as f:
                shutil.copyfileobj(f, self.wfile)
        except (BrokenPipeError, ConnectionResetError):
            pass

    # ---------- 上传：把本机文件搬进票据目录 ----------
    def _json(self, obj, status: int = 200):
        payload = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(payload)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json_err(self, msg, level: str = "ERROR"):
        """接口出错：既回给前端，也记进日志。

        ⚠️ 上传这条路是自己直接写 JSON 的（不走 do_POST 那套统一记录），
        以前出错只回给前端 —— 前端一吞，界面和 docker logs 就都没有线索了。
        """
        _log_line(f"❌ {self.path.split('?')[0]} 失败：{msg}", level)
        return self._json({"result": None, "error": msg})

    def _do_upload(self):
        """收一个文件。body 就是文件原始字节，相对路径放在 X-Upload-Path 头里（base64）。

        为什么不用 multipart/form-data：http.server 没有现成的解析，手写又要处理
        边界、编码一堆事；改成「一条请求一个文件 + 路径走头部」，两边都简单，也方便
        前端串行上传时一个个报进度。
        """
        import base64
        length = int(self.headers.get("Content-Length") or 0)
        try:
            rel = base64.b64decode(self.headers.get("X-Upload-Path") or "").decode("utf-8")
        except Exception:
            rel = ""
        parts = _safe_rel_parts(rel)
        if not parts:
            # 提前拒绝的这几种都直接把连接关掉：body 还没收，不能留在连接里
            self.close_connection = True
            return self._json_err("上传路径不合法（疑似越界）", "WARN")
        name = parts[-1]
        if length <= 0:
            return self._json_err(f"{name}：空文件", "WARN")
        if length > MAX_UPLOAD_BYTES:
            self.close_connection = True
            return self._json_err(f"{name}：{_human_size(length)} 超过单文件上限 "
                                  f"{_human_size(MAX_UPLOAD_BYTES)}", "WARN")
        try:
            root = upload_root()
        except Exception as e:
            self.close_connection = True
            return self._json_err(f"票据目录写不进去：{e}（去「设置」里改一下路径）")

        target = _resolve_under(root, "/".join(parts))       # 再确认一次没跑出 root
        if target is None:
            self.close_connection = True
            return self._json_err("上传路径越界", "WARN")
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.close_connection = True
            return self._json_err(f"建不了目录 {target.parent}：{e}")

        # 同名同大小 → NAS 上已经有这份了，别再写一遍（重复上传同一个文件夹很常见）
        if target.is_file() and target.stat().st_size == length:
            self.close_connection = True
            return self._json({"result": {"ok": True, "mode": "skip", "name": target.name,
                                          "rel": target.relative_to(root).as_posix(),
                                          "root": str(root), "size": length}})

        final = target if not target.exists() else _unique_path(target)
        tmp = final.with_name(final.name + ".part")          # 先写 .part，落地才算数
        got = 0
        try:
            with open(tmp, "wb") as f:
                while got < length:
                    chunk = self.rfile.read(min(_UPLOAD_CHUNK, length - got))
                    if not chunk:
                        break
                    f.write(chunk)
                    got += len(chunk)
            if got != length:
                raise IOError(f"只收到 {got}/{length} 字节")
            tmp.replace(final)                               # 半个文件不会留在票据目录里
        except Exception as e:
            self.close_connection = True
            try:
                tmp.unlink()
            except OSError:
                pass
            return self._json_err(f"{final.name} 写入失败：{e}")

        return self._json({"result": {"ok": True, "mode": "new", "name": final.name,
                                      "rel": final.relative_to(root).as_posix(),
                                      "root": str(root), "size": got}})

    # ---------- 路由 ----------
    def do_GET(self):
        path = self.path.split("?")[0]

        if path == "/login":
            self._send_html(_login_page())
            return
        if path == "/logout":
            self._logout()
            return
        # ping 免鉴权：main() 靠它判断「是不是已经有实例在跑」，不能被登录挡住
        if path.startswith("/api/ping"):
            if self.api:
                self.api.touch()
            body = b'{"ok":true}'
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)
            return
        if not self._authed():
            self._deny()
            return
        if path.startswith("/files"):
            self._serve_files(path)
            return
        super().do_GET()

    def do_POST(self):
        path = self.path.split("?")[0]
        if path == "/login":
            self._try_login()
            return
        if not path.startswith("/api/"):
            self.send_error(404)
            return
        if not self._authed():
            self._deny()
            return
        if self.api:
            self.api.touch()
        # 之后不管调哪个接口，写操作都会把「谁、什么角色、从哪来」记进审计
        db.set_actor(self._username(),
                     self.client_address[0] if self.client_address else "",
                     self._role())
        if path == "/api/upload_raw":        # 上传走原始字节，不套 JSON 那层
            t_up = time.time()
            self._do_upload()
            self._log_api_call("upload_raw", "", t_up)
            return
        name = path[len("/api/"):]
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b"{}"
        try:
            args = json.loads(raw.decode("utf-8")).get("args", [])
        except Exception:
            args = []

        result, error, tb, t0 = None, None, "", time.time()
        if name not in API_METHODS:
            error = f"未开放的方法: {name}"
        elif not db.can(self._username(), self._role(), name):
            # 每个写接口都在这里再过一遍权限：界面藏起按钮只是「不显眼」，不是「不能做」
            role_cn = db.ROLES.get(self._role(), self._role())
            error = f"权限不足：当前身份是「{role_cn}」，不能执行这一步"
            # 日志不在这里单独写：下面的 _log_api_call 会带接口名/身份/耗时记一行，
            # 比这里再补一行「越权尝试」更全（审计照旧单独留痕）。
            db.audit("越权尝试", "api", name, ok=False, error=error)
        else:
            fn = getattr(self.api, name, None)
            try:
                # 耗时写操作统一走任务层：抢重活锁（不让两个人同时改）、留痕、失败可重试。
                # 入库那种自己就会起后台线程的（HEAVY_SPAWN）不再包一层，否则叠两次任务。
                if callable(fn) and name in HEAVY_JOBS and name not in HEAVY_SPAWN:
                    result = self.api._job_call(name, fn, args=tuple(args))
                else:
                    result = fn(*args) if callable(fn) else None
            except Exception as e:
                error = f"{type(e).__name__}: {e}"
                # 完整堆栈进日志。以前这里只把错误字符串塞进返回值的 error 字段，
                # 前端又把它吞了 —— 于是「接口报错」在界面和 docker logs 两边都查不到线索。
                tb = traceback.format_exc()
        self._log_api_call(name, error, t0, tb)

        payload = json.dumps({"result": result, "error": error},
                             ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(payload)

    def _log_api_call(self, name, error, t0, tb=""):
        """每次接口调用都留一行：谁、什么角色、耗时、成没成、为什么败。

        这一行是排障的主线索 —— 用户报「点了没反应/报错」时，先看这里：
        失败原因、当时是什么身份，一眼就能分清「没权限」还是「程序出错」。
        """
        ms = (time.time() - t0) * 1000
        who = self._username() or "(未登录)"
        role_cn = db.ROLES.get(self._role(), self._role()) or "(无角色)"
        if error:
            _log_line(f"❌ {name} 失败（{ms:.0f}ms；{who}/{role_cn}）：{error}", "ERROR")
            if tb:
                _log_line(f"   堆栈（{name}）：\n{tb}", "ERROR")
        elif name not in self.QUIET_API:
            _log_line(f"· {name} ok（{ms:.0f}ms；{who}/{role_cn}）", "DEBUG")

    def end_headers(self):
        if not self.path.startswith("/api/"):
            self.send_header("Cache-Control", "no-store")
        super().end_headers()


def start_server(api, directory: Path, port: int = 0):
    class Handler(UiHandler):
        pass

    Handler.api = api
    handler = functools.partial(Handler, directory=str(directory))
    try:
        srv = http.server.ThreadingHTTPServer((HOST, port), handler)
    except OSError:
        if SERVER_MODE:
            # 容器里端口是映射好的（8766:8766），悄悄换随机端口 = 外面根本访问不到，
            # 所以这里直接抛出去让容器失败重启，比让人对着打不开的页面发懵强。
            raise
        _log_line(f"端口 {port} 被占用，改用随机端口")
        srv = http.server.ThreadingHTTPServer((HOST, 0), handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


def _idle_watchdog(api, stop, timeout=IDLE_TIMEOUT):
    while not stop.wait(5.0):
        if api.busy:
            api.touch()
            continue
        idle = time.time() - api.last_active
        if idle > timeout:
            _log_line(f"界面已 {idle:.0f}s 无请求（阈值 {timeout}s），自动退出")
            stop.set()
            return


def _alive_url():
    """已经有一个在跑吗？（读端口文件 + 探活）"""
    try:
        if not UI_PORT_FILE.exists():
            return ""
        port = int(UI_PORT_FILE.read_text(encoding="utf-8").strip())
    except Exception:
        return ""
    import urllib.request
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/ping", timeout=1.5) as r:
            if r.status == 200:
                return f"http://127.0.0.1:{port}/index.html"
    except Exception:
        return ""
    return ""


def _open_browser(url: str) -> bool:
    try:
        os.startfile(url)
        return True
    except Exception:
        pass
    try:
        return webbrowser.open(url)
    except Exception:
        return False


def main():
    _log_line("=" * 40)
    _log_line(f"启动：exe={getattr(sys, 'frozen', False)} 界面目录={GUI_DIR}")
    _log_line(f"运行模式：{'服务器（Docker/NAS）' if SERVER_MODE else '本机'}"
              f"　监听 {HOST}:{UI_PORT}　数据目录 {DATA_DIR}")
    _log_line(f"打印 PDF 用的浏览器：{EDGE}")
    if _auth_on():
        _log_line("访问口令：已启用（没登录的访问会被挡去登录页）")
    # 启动时把「程序在看哪些目录」说清楚 —— 报「删不掉/找不到文件」时，
    # 第一件事就是核对这里打印的路径跟实际挂载的对不对得上。
    _log_line(f"日志级别：{_LOG_LEVEL}（要看逐接口流水就设 FB_LOG_LEVEL=DEBUG）"
              f"　日志文件：{LAUNCH_LOG}")
    try:
        _log_line(f"票据文件夹：{load_config().get('source') or '(未设置)'}"
                  f"　台账库：{db.db_path()}　输出目录：{OUTPUT_DIR}"
                  f"　回收站：{RECYCLE_DIR}")
    except Exception as e:                                      # noqa: BLE001
        _log_line(f"启动自检读路径失败：{type(e).__name__}: {e}", "WARN")

    if not SERVER_MODE:
        alive = _alive_url()
        if alive:
            _log_line(f"检测到已有实例：{alive} → 直接打开它")
            _open_browser(alive)
            _safe_print(f"[提示] 程序已在运行，已为你打开界面：{alive}")
            return 0

    index = GUI_DIR / "index.html"
    if not index.exists():
        _log_line(f"❌ 界面文件缺失：{index}")
        _safe_print(f"界面文件缺失：{index}")
        return 1

    api = Api()
    try:
        srv = start_server(api, GUI_DIR, UI_PORT)
    except OSError as e:
        msg = f"❌ 端口 {HOST}:{UI_PORT} 起不来：{e}"
        if SERVER_MODE:
            msg += "（这个端口是映射出去给浏览器用的，换个端口容器外就访问不到，"
            msg += "所以这里直接退出 —— 检查一下是不是有另一个实例占着）"
        _log_line(msg)
        _safe_print(msg)
        return 2
    port = srv.server_address[1]
    url = f"http://127.0.0.1:{port}/index.html"
    try:
        UI_PORT_FILE.write_text(str(port), encoding="utf-8")
    except Exception:
        pass
    _log_line(f"界面服务已启动：{url}")

    if SERVER_MODE:
        # 容器里没有浏览器可开，也不该由服务端去开
        _safe_print(f"[就绪] 请用浏览器访问 http://<这台设备的地址>:{port}/")
        _log_line("服务器模式：不自动打开浏览器（请在别的设备上访问）")
    elif _open_browser(url):
        _log_line("已用默认浏览器打开界面")
    else:
        _log_line("❌ 打开浏览器失败，需要手动访问")
        _safe_print(f"[提示] 请手动在浏览器打开：{url}")

    stop = threading.Event()
    if not SERVER_MODE:
        # 本机跑才要「界面关了自动退出」；容器里必须一直活着。
        threading.Thread(target=_idle_watchdog, args=(api, stop), daemon=True).start()
    try:
        import signal
        signal.signal(signal.SIGTERM, lambda *_: api.exit_flag.set())   # docker stop 能干净退出
        signal.signal(signal.SIGINT, lambda *_: api.exit_flag.set())
    except Exception:
        pass
    try:
        while not stop.is_set() and not api.exit_flag.is_set():
            time.sleep(0.4)
    except KeyboardInterrupt:
        _log_line("收到 Ctrl+C")
    try:
        srv.shutdown()
    except Exception:
        pass
    try:
        UI_PORT_FILE.unlink()
    except Exception:
        pass
    _log_line("已退出")
    _safe_print("[退出] 界面服务已关闭")
    return 0


def _fatal(msg: str):
    """启动阶段就挂掉时：既写 launch.log，也弹一个系统对话框。
    因为 bat 用 pythonw 静默启动，没有控制台，光写日志用户看不到。"""
    try:
        _log_line(msg)
    except Exception:
        pass
    try:
        import ctypes
        if SERVER_MODE or os.name != "nt":
            raise RuntimeError("服务器模式：没有桌面可弹框，错误只写日志")
        ctypes.windll.user32.MessageBoxW(None, msg[-1400:],
                                         "发票报销工具 · 启动失败", 0x10)
    except Exception:
        pass


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException:
        import traceback
        _fatal("❌ 启动失败：\n" + traceback.format_exc())
        raise
