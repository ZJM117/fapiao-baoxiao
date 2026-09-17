# -*- coding: utf-8 -*-
"""
隔离自检：排障日志链 + 「按钮亮着但点了必报错」
================================================
用户报了两件事：
  ① 「删除了选中的文件报错」—— 而且「容器日志啥也没有」，没法定位；
  ② 「docker 版本没有账号权限、不能增加账号」。

查下来两件都是同一个根：**后端一直是对的，前端和后端"没接上"** ——
  · delete_rows / clear_ledger 在 db.ROLE_ALLOW 里只有 admin 能调，
    但「删除选中」「清空台账」两颗红按钮对**所有角色**都显示 → 点了必然被挡；
  · 后端的拒绝理由写在返回值里，前端 bridge.js 把 error 字段丢了 → 界面只剩「未知错误」；
  · 后端抛异常时**哪儿都不记日志**（不写文件、不打 stdout、不进界面），docker logs 自然空白。

所以这里守住四件事：
  · 一条日志要同时进：launch.log 文件 / stdout（容器里＝docker logs）/ 界面「运行记录」
  · 接口失败那行必须带「谁、什么角色、为什么失败」—— 一眼能分清「没权限」还是「程序错」
  · 接口抛异常要带完整堆栈（以前连 error 都只是拼个字符串，没有栈）
  · 高频轮询（poll / list_jobs）不许刷屏，也不许把界面运行记录冲掉
  · whoami 要下发**完整权限表**（allowed），界面才能把点不了的按钮压灰

全部在 %TEMP% 副本里跑（起真实服务、走真实 HTTP），无窗口、不碰真实数据。
跑法：  python 测试\\diag_logging.py
"""
import io
import json
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

TMP = Path(tempfile.mkdtemp(prefix="inv_log_"))
PW = "admin-pw-1234"

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
app_paths.PASSWORD = PW

for _d in (TMP, app_paths.LEDGER_BAK, app_paths.OUTPUT_DIR, app_paths.WORK_DIR,
           app_paths.RECYCLE_DIR):
    _d.mkdir(parents=True, exist_ok=True)
app_paths.CONFIG_FILE.write_text(
    json.dumps({"source": str(TMP / "票据"), "recursive": True}), encoding="utf-8")
(TMP / "票据").mkdir(exist_ok=True)

import app_web                                                      # noqa: E402
import db                                                           # noqa: E402

# ---- 把 stdout 通道接出来（容器里它就是 docker logs）----
PRINTED = []
_orig_safe_print = app_web._safe_print


def _tee_print(msg):
    PRINTED.append(msg)
    _orig_safe_print(msg)


app_web._safe_print = _tee_print
# 默认 INFO 会吞掉「接口成功」那行；这里开 DEBUG 才能验证「高频轮询不刷屏」
app_web._LOG_LEVEL = "DEBUG"

FAIL = []


def chk(cond, label, extra=""):
    print(("  [OK] " if cond else "  [XX] ") + label + ("" if cond else f"   → {extra}"))
    if not cond:
        FAIL.append(label)


def section(t):
    print("\n" + "=" * 66 + f"\n{t}\n" + "=" * 66)


def logfile_text():
    try:
        return app_paths.LAUNCH_LOG.read_text(encoding="utf-8")
    except Exception:                                               # noqa: BLE001
        return ""


ADMIN_CK = ""          # 登录后填入（ui_logs 要带着它去 poll，否则 401）


def ui_logs():
    """界面「运行记录」拿到的那份（poll 的增量日志）"""
    r = api_call("poll", 0, cookie=ADMIN_CK)
    return [x.get("m", "") for x in (r.get("logs") or [])]


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k):
        return None


_noredir = urllib.request.build_opener(_NoRedirect)

api = app_web.Api()
srv = app_web.start_server(api, BASE_DIR / "gui", 0)
BASE = f"http://127.0.0.1:{srv.server_address[1]}"
print(f"临时目录：{TMP}\n服务：{BASE}")


def login(username, password):
    body = urllib.parse.urlencode({"username": username, "password": password}).encode()
    req = urllib.request.Request(BASE + "/login", data=body,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        return _noredir.open(req, timeout=20).headers.get("Set-Cookie", "")
    except urllib.error.HTTPError as e:
        return e.headers.get("Set-Cookie", "")


def api_call(name, *args, cookie=""):
    body = json.dumps({"args": list(args)}).encode("utf-8")
    req = urllib.request.Request(BASE + "/api/" + name, data=body,
                                 headers={"Content-Type": "application/json"})
    if cookie:
        req.add_header("Cookie", cookie)
    try:
        r = urllib.request.urlopen(req, timeout=60)
        d = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        return {"http": e.code, "raw": e.read(200).decode("utf-8", "replace")}
    except Exception as e:                                          # noqa: BLE001
        return {"error": f"{type(e).__name__}: {e}"}
    res = d.get("result")
    if isinstance(res, dict):
        if d.get("error"):
            res.setdefault("error", d["error"])
        return res
    return {"result": res, "error": d.get("error")}


def ck_of(cookie):
    return cookie.split(";")[0] if cookie else ""


try:
    print("\n=== 准备：管理员登录 + 建一个业务人员 ===")
    ck_admin = ck_of(login("", PW))
    ADMIN_CK = ck_admin
    ck_biz = ""
    r = api_call("create_user", "wangwu", "ww123456", "biz", "", "业务部", cookie=ck_admin)
    chk(r.get("ok") or not r.get("error"), "管理员能建账号（业务人员）", repr(r)[:120])
    ck_biz = ck_of(login("wangwu", "ww123456"))

    # ==============================================================
    section("一、日志的三个去处，一个都不能少")
    # ==============================================================
    app_web._log_line("诊断探针-PROBE-1")
    chk(any("诊断探针-PROBE-1" in x for x in PRINTED),
        "① stdout（容器里＝docker logs）能收到")
    chk("诊断探针-PROBE-1" in logfile_text(),
        "② launch.log 文件能收到（重启后还能翻）")
    chk(any("诊断探针-PROBE-1" in x for x in ui_logs()),
        "③ 界面「运行记录」能收到")

    # ==============================================================
    section("二、whoami 要下发完整权限表（界面靠它压灰按钮）")
    # ==============================================================
    me_admin = api_call("whoami", cookie=ck_admin)
    chk(me_admin.get("role") == "admin", "口令登录＝管理员", me_admin.get("role"))
    allowed_admin = me_admin.get("allowed")
    chk(isinstance(allowed_admin, list) and len(allowed_admin) > 20,
        "管理员拿到完整 allowed 表（不是只下发几个 can 键）",
        f"type={type(allowed_admin).__name__} n={len(allowed_admin or [])}")
    chk("delete_rows" in (allowed_admin or []), "管理员 allowed 里有 delete_rows")

    me_biz = api_call("whoami", cookie=ck_biz)
    chk(me_biz.get("role") == "biz", "业务人员登录成功", me_biz.get("role"))
    allowed_biz = me_biz.get("allowed") or []
    for m in ("delete_rows", "clear_ledger", "list_users", "audit_list"):
        chk(m not in allowed_biz, f"业务人员 allowed 里没有 {m}（界面据此压灰那颗按钮）")
    chk("make_report" in allowed_biz, "业务人员 allowed 里仍有 make_report（该给的不能少）")

    # ==============================================================
    section("三、业务人员点「删除选中」：拒绝理由必须写进日志")
    # ==============================================================
    PRINTED.clear()
    r = api_call("delete_rows", ["1"], cookie=ck_biz)
    chk("权限不足" in str(r.get("error") or ""),
        "后端明确拒绝，并说清是身份问题", repr(r)[:140])
    hit = [x for x in PRINTED if "delete_rows" in x and "权限不足" in x]
    chk(bool(hit), "stdout 里有「delete_rows 失败」那一行（docker logs 查得到）",
        repr(PRINTED[-3:])[:200])
    if hit:
        chk("wangwu" in hit[0] and "业务人员" in hit[0],
            "那行带「谁 + 什么角色」——一眼分清是没权限还是程序错", hit[0][:160])
    chk(any("delete_rows" in x and "权限不足" in x for x in logfile_text().splitlines()),
        "launch.log 里同样有这一行")
    chk(any("delete_rows" in x and "权限不足" in x for x in ui_logs()),
        "界面「运行记录」里也有 —— 不用 SSH 就能看到原因")

    # ==============================================================
    section("四、接口抛异常：日志里必须有完整堆栈")
    # ==============================================================
    PRINTED.clear()
    # set_phone_name 必须带 phone 参数；空参调用一定抛 TypeError（拿它当「一定会抛」的样本）
    r = api_call("set_phone_name", cookie=ck_admin)
    chk("TypeError" in str(r.get("error") or ""),
        "异常被回成 error（不是静默吞掉）", repr(r.get("error"))[:120])
    tb_lines = [x for x in PRINTED if "堆栈（set_phone_name）" in x]
    chk(bool(tb_lines), "stdout 里打了堆栈标题", repr(PRINTED[-2:])[:200])
    chk(any("Traceback" in x for x in PRINTED),
        "堆栈是完整的 Traceback（能直接看出哪一行炸的）")
    chk("set_phone_name" in logfile_text() and "Traceback" in logfile_text(),
        "堆栈也落进了 launch.log")

    # ==============================================================
    section("五、高频轮询不许刷屏（也不许冲掉界面运行记录）")
    # ==============================================================
    PRINTED.clear()
    before = len(app_web._LOGS)
    for _ in range(6):
        api_call("poll", 0, cookie=ADMIN_CK)
    for _ in range(3):
        api_call("list_jobs", 50, cookie=ck_admin)
    chk(not any("poll ok" in x for x in PRINTED), "连调 6 次 poll 不产生日志行")
    chk(not any("list_jobs ok" in x for x in PRINTED), "连调 3 次 list_jobs 不产生日志行")
    chk(len(app_web._LOGS) == before,
        "轮询没往界面运行记录里塞东西（否则真要看的会被冲走）",
        f"{before} → {len(app_web._LOGS)}")

    PRINTED.clear()
    api_call("get_ledger", {}, cookie=ck_admin)
    ok_lines = [x for x in PRINTED if "get_ledger ok" in x]
    chk(bool(ok_lines), "但用户真点的操作（查台账）有留痕（DEBUG 级）", repr(PRINTED[-2:])[:200])
    chk(not any("get_ledger ok" in x for x in ui_logs()),
        "DEBUG 流水只进文件/stdout，不塞给界面运行记录")

    # ==============================================================
    section("六、没登录就来调接口：也要留痕（而不是静默 401）")
    # ==============================================================
    PRINTED.clear()
    r = api_call("get_ledger", {})
    chk(r.get("http") == 401, "未登录请求接口 → 401", repr(r)[:120])
    chk(any("未登录" in x for x in PRINTED), "日志里记了一笔「未登录就访问接口」",
        repr(PRINTED[-2:])[:200])

    # ==============================================================
    section("七、stdout 编码必须是 UTF-8（否则中文日志会被 except 吞掉＝什么都没打）")
    # ==============================================================
    chk(hasattr(app_web, "_force_utf8_streams"), "有 _force_utf8_streams()")
    enc = getattr(sys.stdout, "encoding", "") or ""
    chk("utf" in enc.lower(), "本进程 stdout 已按 UTF-8 编码", enc)
    chk(any("诊断探针" in x or "台账" in x for x in PRINTED + ui_logs()),
        "中文日志能真的写出来（没被编码异常吞掉）")

finally:
    try:
        srv.shutdown()
    except Exception:                                               # noqa: BLE001
        pass

print("\n" + "=" * 66)
if FAIL:
    print(f"❌ 有 {len(FAIL)} 项没通过：")
    for f in FAIL:
        print("   -", f)
    sys.exit(1)
print("✅ 全部通过")
