# -*- coding: utf-8 -*-
"""
隔离自检：用户、角色与权限（改造清单 P0-2）
==========================================
原来只有一个「共享口令」：谁都能进，进去以后什么都能干，出了事查不到是谁。
这里验证升级后的行为：

  · 没有账号时用访问口令自动建管理员（不然部署完 NAS 人进不去）
  · 用户名 + 密码登录，会话里记住「是谁」
  · 角色卡口：业务人员不能清空台账、只读用户不能出单据、非管理员看不到用户列表
  · 改密码 / 停用 / 强制退出后，旧的会话立刻作废（服务端作废，不只是丢 Cookie）
  · 口令哈希绝不下发到浏览器

全部在 %TEMP% 副本里跑（起真实服务、走真实 HTTP），无窗口、不碰真实数据。
跑法：  python 测试\\diag_auth.py
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

TMP = Path(tempfile.mkdtemp(prefix="inv_auth_"))
PW = "admin-pw-1234"

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
import db                                                           # noqa: E402

FAIL = []


def chk(cond, label, extra=""):
    print(("  [OK] " if cond else "  [XX] ") + label + ("" if cond else f"   → {extra}"))
    if not cond:
        FAIL.append(label)


def section(t):
    print("\n" + "=" * 66 + f"\n{t}\n" + "=" * 66)


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k):
        return None


_noredir = urllib.request.build_opener(_NoRedirect)

for d in (TMP, app_paths.LEDGER_BAK, app_paths.OUTPUT_DIR, app_paths.WORK_DIR):
    d.mkdir(parents=True, exist_ok=True)
app_paths.CONFIG_FILE.write_text(json.dumps({"source": str(TMP / "票据")}),
                                 encoding="utf-8")

api = app_web.Api()                    # 构造时就会建库 + 用访问口令建管理员
srv = app_web.start_server(api, BASE_DIR / "gui", 0)
BASE = f"http://127.0.0.1:{srv.server_address[1]}"
print(f"临时目录：{TMP}\n服务：{BASE}")


def login(username, password):
    body = urllib.parse.urlencode({"username": username, "password": password}).encode()
    req = urllib.request.Request(BASE + "/login", data=body,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        r = _noredir.open(req, timeout=20)
        return r.headers.get("Set-Cookie", "")
    except urllib.error.HTTPError as e:
        return e.headers.get("Set-Cookie", "")


def api_call(name, *args, cookie=""):
    """调一个后端接口。

    后端的壳是 {"result":…, "error":…} —— 这里把 result 那一层拆开，
    接口自己返回的 dict（{"ok":…} / {"error":…}）直接用，少一层 _get 的绕。
    """
    body = json.dumps({"args": list(args)}).encode("utf-8")
    req = urllib.request.Request(BASE + "/api/" + name, data=body,
                                 headers={"Content-Type": "application/json"})
    if cookie:
        req.add_header("Cookie", cookie)
    try:
        r = urllib.request.urlopen(req, timeout=30)
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
    # ==============================================================
    section("一、没有账号时用访问口令自动建管理员")
    # ==============================================================
    users = db.user_list()
    chk(len(users) == 1 and users[0]["username"] == "admin",
        "自动建出了 admin 账号", [u["username"] for u in users])
    chk(users[0]["role"] == "admin", "而且角色是管理员", users[0]["role"])
    chk(not users[0]["pw_hash"].startswith(PW), "口令不是明文存的", users[0]["pw_hash"][:20])

    # ==============================================================
    section("二、登录：谁进来了")
    # ==============================================================
    ck = ck_of(login("admin", PW))
    chk("fb_auth=" in ck, "admin 用账号密码登录成功")
    r = api_call("list_users", cookie=ck)
    chk(r.get("me") == "admin", "会话里记住了「是谁」", r.get("me"))
    chk(r.get("my_role") == "admin", "也记住了角色", r.get("my_role"))
    code = api_call("clear_ledger", cookie="")           # 不带 Cookie
    chk(code.get("http") == 401, "没登录调接口 → 401", code)

    # ==============================================================
    section("三、给同事开账号")
    # ==============================================================
    r = api_call("create_user", "xiaozhang", "zhang123456", "biz", "小张", "财务部", cookie=ck)
    chk(r.get("ok"), "管理员能新增用户", r)
    r = api_call("create_user", "xiaozhang", "whatever123", "biz", cookie=ck)
    chk("已经有人用了" in str(r.get("error") or ""), "用户名不能重复", r.get("error"))
    r = api_call("create_user", "lisi", "123", "biz", cookie=ck)
    chk("至少 6 位" in str(r.get("error") or ""), "弱密码被拦", r.get("error"))
    api_call("create_user", "lisi", "lisi123456", "viewer", "李四", cookie=ck)

    rows = api_call("list_users", cookie=ck).get("rows") or []
    chk(all("pw_hash" not in u for u in rows), "用户列表里没有口令哈希（不下发）")
    chk(any(u.get("role_cn") == "业务人员" for u in rows), "带上了中文角色名")

    # ==============================================================
    section("四、角色卡口（核心）")
    # ==============================================================
    zk = ck_of(login("xiaozhang", "zhang123456"))
    lk = ck_of(login("lisi", "lisi123456"))
    chk(bool(zk) and bool(lk), "两个同事都登录成功")

    r = api_call("get_ledger", cookie=zk)
    chk("rows" in r, "业务人员能看台账", list(r)[:4])
    r = api_call("check_source", "", cookie=zk)
    chk("error" not in r or "权限不足" not in str(r.get("error")), "业务人员能查文件夹")

    r = api_call("clear_ledger", cookie=zk)
    chk("权限不足" in str(r.get("error") or ""), "业务人员不能清空台账", r.get("error"))
    r = api_call("delete_rows", [1], cookie=zk)
    chk("权限不足" in str(r.get("error") or ""), "业务人员不能删台账行", r.get("error"))
    r = api_call("list_users", cookie=zk)
    chk("权限不足" in str(r.get("error") or ""), "业务人员看不到用户列表", r.get("error"))
    r = api_call("set_config", {"x": 1}, cookie=zk)
    chk("权限不足" in str(r.get("error") or ""), "业务人员不能改配置", r.get("error"))

    r = api_call("make_report", [], cookie=lk)
    chk("权限不足" in str(r.get("error") or ""), "只读用户不能出报销单", r.get("error"))
    r = api_call("get_ledger", cookie=lk)
    chk("rows" in r, "只读用户能看台账", list(r)[:4])

    # ==============================================================
    section("五、改密码 / 停用 / 强制退出 → 旧会话立刻作废")
    # ==============================================================
    r = api_call("set_user_password", "xiaozhang", "zhang-new-456", cookie=ck)
    chk(r.get("ok") and r.get("kicked", 0) >= 1, "改密码时把旧会话踢掉了", r)
    r = api_call("get_ledger", cookie=zk)
    chk(r.get("http") == 401, "小张的旧 Cookie 立刻失效", r)
    chk(ck_of(login("xiaozhang", "zhang-new-456")), "用新密码能登录")

    r = api_call("set_user_disabled", "xiaozhang", True, cookie=ck)
    chk(r.get("ok"), "停用用户", r)
    chk(not ck_of(login("xiaozhang", "zhang-new-456")), "停用后登不进来")
    r = api_call("set_user_disabled", "admin", True, cookie=ck)
    chk("不能停用" in str(r.get("error") or ""), "内置管理员不能被停用", r.get("error"))

    n_before = len(app_web._SESSIONS)
    r = api_call("force_logout", "lisi", cookie=ck)
    chk(r.get("kicked", 0) >= 1 and len(app_web._SESSIONS) < n_before,
        "强制退出后服务端的会话表里就没了", r)

    # ==============================================================
    section("六、登录失败的几种情况")
    # ==============================================================
    chk(not ck_of(login("xiaozhang", "错的密码")), "密码不对登不进")
    chk(not ck_of(login("查无此人", "无所谓")), "不存在的账号登不进")
    chk(bool(ck_of(login("", PW))), "只填访问口令仍能进（老 NAS 部署不受影响）")
    r = api_call("get_ledger", cookie=ck_of(login("", PW)))
    chk("rows" in r, "用口令进来的是管理员身份，干活不受限", list(r)[:4])

    failed = db.list_audit(limit=50, action="登录失败")
    chk(len(failed) >= 2, "登录失败也进了审计（能看出有人在猜密码）", len(failed))
    ok_login = db.list_audit(limit=50, action="登录成功")
    chk(any(x["username"] == "admin" for x in ok_login), "成功的登录记下了是谁")
    chk(all(x["ip"] for x in ok_login), "也记下了从哪个 IP 来")

    # ==============================================================
    section("七、界面要用的那几个接口")
    # ==============================================================
    api_call("set_user_disabled", "xiaozhang", False, cookie=ck)
    zk = ck_of(login("xiaozhang", "zhang-new-456"))
    chk(bool(zk), "重新启用之后又能登录了")

    w = api_call("whoami", cookie=zk)
    chk(w.get("user") == "xiaozhang" and w.get("role") == "biz",
        "whoami 说得出「我是谁、什么角色」", w)
    chk(w.get("can", {}).get("clear_ledger") is False,
        "whoami 的 can 表：业务人员不能清空台账（界面据此藏按钮）")
    chk(w.get("can", {}).get("start_import") is True, "也说得出他能入库")
    chk(w.get("can", {}).get("list_users") is False, "也说得出他看不了用户管理")

    info = api_call("ledger_db_info", cookie=zk)
    chk(info.get("ok") and "发票台账.db" in str(info.get("db")),
        "数据库状态查得到（设置页那块用它）", info)
    chk(info.get("invoices") == 0, "刚建库，台账 0 行", info.get("invoices"))
    chk(info.get("users") == 3, "账号数对得上（admin + 小张 + 李四）", info.get("users"))

    au = api_call("audit_list", 50, cookie=zk)
    chk("权限不足" in str(au.get("error") or ""),
        "审计日志不给非管理员（谁改了什么，只该管理员看得到）", au.get("error"))
    au2 = api_call("audit_list", 50, cookie=ck)
    chk("rows" in au2 and au2.get("total", 0) > 0, "管理员能看审计日志",
        {k: au2.get(k) for k in ("total",)})
    chk(all("pw_hash" not in str(x) for x in (au2.get("rows") or [])),
        "审计内容里也没有口令哈希")

    gj = api_call("list_jobs", cookie=zk)
    chk("rows" in gj, "后台任务列表能读", list(gj)[:3])

    gf = api_call("gen_files", cookie=zk)
    chk("rows" in gf, "生成文件清单能读", list(gf)[:3])

finally:
    srv.shutdown()

print("\n" + "=" * 66)
print(f"结果：{'全部通过' if not FAIL else f'有失败（{len(FAIL)} 项）'}")
for f in FAIL:
    print("  ✗", f)
sys.exit(1 if FAIL else 0)
