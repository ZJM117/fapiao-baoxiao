# -*- coding: utf-8 -*-
"""
隔离自检：/files/ 的暴露面（改造清单 P1-7）
==========================================
程序跑在 NAS 容器里时，「打开文件」是靠浏览器直接拉 /files/<别名>/... 实现的
（容器里没有资源管理器）。这条通道以前把整个数据目录都挂了出去 —— 那里面有
config.json（口令就在里面）、手机号姓名对照表、运行日志、台账备份和数据库本体。
局域网里随便谁都能拉到，等于把口令明着贴在墙上。

这里逐条验证：该给的给得了，不该给的一个都拿不到。

全部在 %TEMP% 副本里跑（起真实服务、走真实 HTTP），无窗口。
跑法：  python 测试\\diag_files_guard.py
"""
import io
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))

TMP = Path(tempfile.mkdtemp(prefix="inv_guard_"))
PW = "test-pw-123"

# ⚠️ 必须在 import app_web 之前改路径 / 口令（它用 from app_paths import ... 取值）
import app_paths                                                    # noqa: E402
app_paths.DATA_DIR = TMP
app_paths.LEDGER_FILE = TMP / "发票台账.xlsx"
app_paths.LEDGER_BAK = TMP / "台账备份"
app_paths.OUTPUT_DIR = TMP / "输出"
app_paths.WORK_DIR = TMP / "缓存"
app_paths.CONFIG_FILE = TMP / "config.json"
app_paths.LAUNCH_LOG = TMP / "launch.log"
app_paths.PASSWORD = PW

import app_web                                                      # noqa: E402

FAIL = []


def chk(cond, label, extra=""):
    print(("  [OK] " if cond else "  [XX] ") + label + ("" if cond else f"   → {extra}"))
    if not cond:
        FAIL.append(label)


def section(t):
    print("\n" + "=" * 66 + f"\n{t}\n" + "=" * 66)


# ---------------------------------------------------------------- 造数据
TICKETS = TMP / "票据"
for d in (TMP, app_paths.LEDGER_BAK, app_paths.OUTPUT_DIR, app_paths.WORK_DIR, TICKETS):
    d.mkdir(parents=True, exist_ok=True)

(app_paths.CONFIG_FILE).write_text(
    '{"source": "%s", "web_password": "should-not-leak"}' % str(TICKETS).replace("\\", "\\\\"),
    encoding="utf-8")
(TMP / "手机号姓名.json").write_text('{"13800000000": "张三"}', encoding="utf-8")
(TMP / "出行人核对.json").write_text('{"x": "李四"}', encoding="utf-8")
(TMP / "launch.log").write_text("启动日志，里面有内网路径", encoding="utf-8")
# ⚠️ 别往 发票台账.db 塞假内容 —— 那正是 db.db_path() 要用的文件，
#    塞了就变成「file is not a database」。真数据库由 Api() 启动时自己建。
(app_paths.LEDGER_BAK / "发票台账_20260916_085038.xlsx").write_bytes("假备份".encode("utf-8"))
(app_paths.WORK_DIR / "parse_cache.json").write_text("{}", encoding="utf-8")
(TICKETS / "车票.pdf").write_bytes(b"%PDF-1.4 " + "真票据".encode("utf-8"))
(app_paths.OUTPUT_DIR / "报销单.pdf").write_bytes(b"%PDF-1.4 " + "真报销单".encode("utf-8"))
# 台账 Excel 快照（只有表头）：它是要给人下载的，不在拦截名单里
import openpyxl                                                     # noqa: E402
_wb = openpyxl.Workbook()
_ws = _wb.active
_ws.title = "发票台账"
_ws.append(["序号", "状态", "凭证类型", "票面标记", "出行人", "发票号码", "开票日期", "票种",
            "行程/明细", "购方名称", "销方名称", "不含税金额", "税额", "价税合计", "项目/事由",
            "项目编号", "费用类别", "报销人", "报销批次", "入库时间", "文件路径", "提示"])
_wb.save(str(app_paths.LEDGER_FILE))

api = app_web.Api()
srv = app_web.start_server(api, BASE / "gui", 0)
PORT = srv.server_address[1]
BASE_URL = f"http://127.0.0.1:{PORT}"
print(f"临时目录：{TMP}\n服务：{BASE_URL}")


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k):
        return None


_opener = urllib.request.build_opener(_NoRedirect)


def http(path, cookie="", follow=False):
    """返回 (状态码, 正文前 6000 字)

    ⚠️ URL 里有中文（手机号姓名.json 这种）必须先 quote —— urllib 默认拿 ascii 编码，
       不引号会直接 UnicodeEncodeError，看起来像「服务端挂了」，其实是客户端没编码。
    """
    url = BASE_URL + urllib.parse.quote(path, safe="/%?=&")
    req = urllib.request.Request(url)
    if cookie:
        req.add_header("Cookie", cookie)
    op = urllib.request.urlopen if follow else _opener.open
    try:
        r = op(req, timeout=20)
        return getattr(r, "status", r.getcode()), r.read(6000).decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        return e.code, e.read(6000).decode("utf-8", "replace")
    except Exception as e:                                          # noqa: BLE001
        return 0, f"{type(e).__name__}: {e}"


def login():
    body = urllib.parse.urlencode({"password": PW}).encode("utf-8")
    req = urllib.request.Request(BASE_URL + "/login", data=body,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        r = _opener.open(req, timeout=20)
        return r.headers.get("Set-Cookie", "")
    except urllib.error.HTTPError as e:
        return e.headers.get("Set-Cookie", "")


try:
    # ==============================================================
    section("一、没登录的时候什么都拿不到")
    # ==============================================================
    code, _ = http("/files/data/config.json")
    chk(code == 302, "未登录访问 /files/ 被弹去登录页", code)
    code, _ = http("/files/source/车票.pdf")
    chk(code == 302, "未登录连票据也拿不到", code)

    cookie = login()
    chk("fb_auth=" in cookie, "口令正确，拿到会话 Cookie")
    cookie = cookie.split(";")[0]

    # ==============================================================
    section("二、敏感文件一个都不给（核心）")
    # ==============================================================
    for path, label in [
        ("/files/data/config.json", "配置（口令就在里面）"),
        ("/files/data/手机号姓名.json", "手机号→姓名对照表"),
        ("/files/data/出行人核对.json", "出行人核对记录"),
        ("/files/data/launch.log", "运行日志"),
        ("/files/data/发票台账.db", "数据库本体"),
        ("/files/data/发票台账.db-wal", "数据库预写日志"),
        ("/files/data/台账备份/发票台账_20260916_085038.xlsx", "台账备份目录"),
        ("/files/data/缓存/parse_cache.json", "缓存目录"),
    ]:
        code, body = http(path, cookie)
        chk(code == 403, f"{label} → 403", f"实际 {code} {body[:60]}")

    # ==============================================================
    section("三、该给的还得给")
    # ==============================================================
    code, body = http("/files/source/车票.pdf", cookie)
    chk(code == 200 and "真票据" in body, "票据原件能下载", f"{code}")
    code, body = http("/files/output/报销单.pdf", cookie)
    chk(code == 200 and "真报销单" in body, "生成的报销单能下载", f"{code}")

    # ==============================================================
    section("四、目录列表：不暴露服务器绝对路径、不列出被拦的东西")
    # ==============================================================
    code, body = http("/files/", cookie)
    chk(code == 200, "能列出可浏览的目录", code)
    chk(str(TMP) not in body and str(TMP).replace("\\", "/") not in body,
        "根列表里没有服务器绝对路径")
    chk("output" in body and "source" in body, "该有的别名还在")

    code, body = http("/files/data/", cookie)
    chk(code == 200, "数据目录本身能列", code)
    chk("config.json" not in body, "列表里看不到 config.json")
    chk("发票台账.db" not in body, "列表里看不到数据库")
    chk("台账备份" not in body, "列表里看不到台账备份目录")
    chk("缓存" not in body, "列表里看不到缓存目录")
    chk("发票台账.xlsx" in body, "但台账 Excel 要看得见（用户要下载它）")

    # ==============================================================
    section("五、路径穿越仍然拦得住")
    # ==============================================================
    code, _ = http("/files/data/../config.json", cookie)
    chk(code in (403, 404), ".. 穿越被拦", code)
    code, _ = http("/files/source/%2e%2e/config.json", cookie)
    chk(code in (403, 404), "编码过的 .. 也被拦", code)
    code, _ = http("/files/nosuchalias/x.txt", cookie)
    chk(code == 404, "不存在的别名 404", code)

finally:
    srv.shutdown()

print("\n" + "=" * 66)
print(f"结果：{'全部通过' if not FAIL else f'有失败（{len(FAIL)} 项）'}")
for f in FAIL:
    print("  ✗", f)
sys.exit(1 if FAIL else 0)
