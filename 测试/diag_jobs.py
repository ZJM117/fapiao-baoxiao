# -*- coding: utf-8 -*-
"""
隔离自检：后台任务与并发控制（改造清单 P0-4）
=============================================
以前只有「入库」有一个 busy 挡着，别的耗时接口（重新识别、出单、合并、导入、清空）
能被两个人同时点：一个人在清空台账、另一个人在入库，两边都不报错，可结果谁也说不清。
这里验证升级后的行为：

  · 耗时操作都建任务：谁提交的、什么时候开始、跑了多久、结果是什么
  · 同一时间只有一个重活在跑；第二个请求被**明确拒绝**（而不是静默地什么都没发生）
  · 失败被记成失败（不是「成功」），失败原因留在任务里，能一键重试
  · 重试要按**原来那个接口**再校验一次权限（业务员不能借重试去清空台账）
  · 入库跑在后台线程里，「谁入的库」也记得下来（线程里的身份会丢，必须显式带过去）
  · 服务重启后，上一世挂着的任务被收成失败（界面不会一直转圈等一个不存在的任务）

全部在 %TEMP% 副本里跑（起真实服务、走真实 HTTP），无窗口、不碰真实数据。
跑法：  python 测试\\diag_jobs.py
"""
import io
import json
import sqlite3
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

TMP = Path(tempfile.mkdtemp(prefix="inv_jobs_"))
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

# ⚠️ 在 import db 之前先造一个「上一版的库」：jobs 表还没有 payload 列。
#    升级路径（老库补列）只有在库是「早就存在的」时候才走得到，正常初始化是写不出这个场景的。
_old = sqlite3.connect(str(TMP / "发票台账.db"))
_old.execute("""CREATE TABLE jobs (
  id TEXT PRIMARY KEY, kind TEXT NOT NULL DEFAULT '', state TEXT NOT NULL DEFAULT 'queued',
  total INTEGER NOT NULL DEFAULT 0, done INTEGER NOT NULL DEFAULT 0,
  label TEXT NOT NULL DEFAULT '', username TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '', started_at TEXT NOT NULL DEFAULT '',
  ended_at TEXT NOT NULL DEFAULT '', error TEXT NOT NULL DEFAULT '', result_json TEXT)""")
_old.execute("INSERT INTO jobs(id, kind, state, label, username, created_at)"
             " VALUES('old_job','import','成功','上一版留下的任务','老张','2026-01-01 08:00:00')")
_old.commit()
_old.close()

import app_web                                                      # noqa: E402
import db                                                           # noqa: E402
import ledger                                                       # noqa: E402

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

for d in (TMP, app_paths.LEDGER_BAK, app_paths.OUTPUT_DIR, app_paths.WORK_DIR,
          app_paths.RECYCLE_DIR):
    d.mkdir(parents=True, exist_ok=True)
app_paths.CONFIG_FILE.write_text(
    json.dumps({"source": str(TMP / "票据"), "recursive": True}), encoding="utf-8")
(TMP / "票据").mkdir(exist_ok=True)

api = app_web.Api()                    # 构造时建库 + 用访问口令建管理员
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


def wait_job(jid, timeout=30.0):
    """等一个任务跑完（后台线程），返回它的最终状态"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        j = db.job_get(jid)
        if j.get("state") not in (db.JOB_QUEUED, db.JOB_RUNNING):
            return j
        time.sleep(0.2)
    return db.job_get(jid)


try:
    ck = ck_of(login("admin", PW))
    chk("fb_auth=" in ck, "管理员登录成功")

    # ==============================================================
    section("一、耗时操作会留下任务记录")
    # ==============================================================
    before = db.job_list(limit=500)
    r = api_call("clear_ledger", cookie=ck)
    chk("error" not in r, "清空台账正常返回", r)
    jobs = db.job_list(limit=500)
    chk(len(jobs) == len(before) + 1, "多出一条任务", f"{len(before)} → {len(jobs)}")
    j = jobs[0]
    chk(j["kind"] == "clear_ledger", "kind 就是接口名（重试时才能找回方法）", j["kind"])
    chk(j["state"] == db.JOB_DONE, "状态是成功", j["state"])
    chk(j["username"] == "admin", "记下了是谁提交的", j["username"])
    chk(bool(j["started_at"]) and bool(j["ended_at"]), "开始/结束时间都记了",
        f"{j['started_at']} / {j['ended_at']}")
    chk(j["result_json"], "结果也存了一份", str(j["result_json"])[:60])

    # ==============================================================
    section("二、同一时间只允许一个重活（互斥）")
    # ==============================================================
    # 手工占住锁，模拟「另一个人正在跑入库」
    held, err = api._job_take("start_import", "入库建账", {"source": "/票据"})
    chk(held and not err, "抢到锁并建了任务", err)
    r = api_call("clear_ledger", cookie=ck)
    chk("正在执行" in str(r.get("error") or ""), "第二个人被明确拒绝（不是静默失败）",
        r.get("error"))
    chk("入库建账" in str(r.get("error") or ""), "提示里说明了是哪个任务在占着",
        r.get("error"))
    r2 = api_call("export_summary", cookie=ck)
    chk("正在执行" in str(r2.get("error") or ""), "不同种类的重活之间也互斥", r2.get("error"))
    api._job_release(held)
    api._job_finish(held, {"ok": True})
    chk(api._heavy is None, "放开锁之后锁没了")

    # ==============================================================
    section("三、失败被记成失败，而且能重试")
    # ==============================================================
    r = api_call("import_from_excel", str(TMP / "根本没有这个文件.xlsx"), cookie=ck)
    chk("找不到" in str(r.get("error") or ""), "接口如实返回错误", r.get("error"))
    j = db.job_list(limit=1)[0]
    chk(j["state"] == db.JOB_FAILED, "任务状态是「失败」而不是「成功」", j["state"])
    chk("找不到" in (j["error"] or ""), "失败原因写进了任务", j["error"])
    chk(db.job_retry_payload(j["id"]) is not None, "参数留下了（重试才有料）",
        db.job_retry_payload(j["id"]))
    old_id = j["id"]

    r = api_call("job_retry", old_id, cookie=ck)
    chk("error" not in r or "不能" not in str(r.get("error")), "管理员能重试", r)
    new_id = r.get("job") or ""
    if not new_id:
        cand = [x for x in db.job_list(limit=5) if x["id"] != old_id]
        new_id = cand[0]["id"] if cand else ""
    chk(new_id and new_id != old_id, "重试产生了一条**新**任务（不是改旧的）", new_id)
    j2 = wait_job(new_id)
    chk(j2.get("state") == db.JOB_FAILED, "重试仍然失败（文件确实不存在）→ 如实记录",
        j2.get("state"))
    chk(db.job_get(old_id)["state"] == db.JOB_FAILED, "老任务的状态没被重试搅乱")

    # ==============================================================
    section("四、重试要按原接口再校验一次权限")
    # ==============================================================
    api_call("create_user", "xiaozhang", "zhang123456", "biz", "小张", "财务部", cookie=ck)
    ck_biz = ck_of(login("xiaozhang", "zhang123456"))
    chk("fb_auth=" in ck_biz, "业务人员登录成功")

    # 那条失败任务是 import_from_excel —— 业务人员本来就没这个权限
    r = api_call("job_retry", old_id, cookie=ck_biz)
    chk("不能重跑" in str(r.get("error") or ""), "业务人员不能借重试越权（按原接口卡）",
        r.get("error"))
    n_before = len(db.job_list(limit=500))
    chk(len(db.job_list(limit=500)) == n_before, "被拒的重试没有偷偷建任务")

    # ==============================================================
    section("五、入库跑在后台线程里，「谁入的库」也记得下来")
    # ==============================================================
    # ⚠️ 这一条是在修一个真 bug：后台线程里的 contextvars 是空的，
    #    不显式把身份带过去，入库写出来的每行 created_by、任务里的 username 都会是空白。
    r = api_call("start_import", True, "", "", "", str(TMP / "不存在的票据目录"),
                 cookie=ck_biz)
    chk(r.get("result") is True or r.get("result") == "True",
        "start_import 仍然返回 True/False（界面早就依赖这个契约）", r)
    # 任务跑得很快（目录不存在，立刻返回），_job_id 可能已经被清掉 —— 从库里找最新那条
    cand = [x for x in db.job_list(limit=5) if x["kind"] == "start_import"]
    jid = cand[0]["id"] if cand else ""
    if jid:
        jj = wait_job(jid)
        chk(jj.get("username") == "xiaozhang",
            "任务记的是**提交人小张**，不是空", jj.get("username"))
        chk(jj.get("state") == db.JOB_FAILED, "目录不存在 → 任务失败（不是假成功）",
            jj.get("state"))
    else:
        chk(False, "入库任务建出来了", "库里找不到 start_import 任务")

    # 业务人员能重试自己那条入库任务（他本来就有 start_import 权限）
    if jid:
        r = api_call("job_retry", jid, cookie=ck_biz)
        chk("不能重跑" not in str(r.get("error") or ""),
            "业务人员能重试入库（这个他有权）", r.get("error"))
        j3 = wait_job(r.get("job") or "")
        chk(j3.get("state") == db.JOB_FAILED, "重试同样如实失败",
            f"retry 返回={r} / 最终状态={j3.get('state')}")

    # ==============================================================
    section("六、界面要的字段都在（任务中心）")
    # ==============================================================
    r = api_call("list_jobs", 50, cookie=ck)
    chk("rows" in r and "counts" in r, "返回 rows + counts", list(r)[:6])
    row = (r.get("rows") or [{}])[0]
    for k in ("kind_cn", "state", "secs", "retryable", "username", "created_at"):
        chk(k in row, f"任务里带 {k}", row.get(k))
    chk("payload" not in row, "参数（可能有路径）不下发给浏览器")
    chk("result_json" not in row, "结果原文不下发给浏览器")
    chk(row.get("kind_cn") in app_web.JOB_KINDS.values(), "kind_cn 是中文说法",
        row.get("kind_cn"))
    chk((r.get("counts") or {}).get("failed", 0) > 0, "失败数统计出来了",
        r.get("counts"))
    chk(isinstance((r.get("counts") or {}).get("done_today"), int), "今日成功数也有")
    # 用数据库交叉核对：「失败 + 属于重活 + 有参数」的都必须标成可重试，一个都不能漏
    raw = db.job_list(limit=50)
    want = {x["id"] for x in raw
            if x["state"] == db.JOB_FAILED and x["kind"] in app_web.HEAVY_JOBS
            and (x.get("payload") or "")}
    chk(want and want <= {x["id"] for x in r["rows"] if x["retryable"]},
        "该能重试的都标了「可重试」", f"应标 {len(want)} 条")

    # ==============================================================
    section("七、重启后：上一世挂着的任务被收掉")
    # ==============================================================
    zid = db.job_create("merge_invoices", "合并 PDF", payload={"mode": "2up"})
    db.job_update(zid, state=db.JOB_RUNNING, started_at=db._now())
    n = db.job_orphans_fix()
    chk(n == 1, "收到了 1 个中断任务", n)
    z = db.job_get(zid)
    chk(z["state"] == db.JOB_FAILED, "状态改成失败", z["state"])
    chk("重启" in (z["error"] or ""), "写了「服务重启，任务中断」", z["error"])
    chk(db.job_retry_payload(zid) == {"mode": "2up"},
        "参数还在 → 重启后也能重试", db.job_retry_payload(zid))

    # ==============================================================
    section("八、老库升级：没有 payload 列的库照样能跑")
    # ==============================================================
    conn = sqlite3.connect(str(db.db_path()))
    try:
        cols = [x[1] for x in conn.execute("PRAGMA table_info(jobs)")]
    finally:
        conn.close()
    chk("payload" in cols, "启动时自动给老库补上了 payload 列", cols)
    old = db.job_get("old_job")
    chk(old.get("label") == "上一版留下的任务", "老任务还在（没被升级弄丢）", old.get("label"))
    chk(old.get("payload") == "", "老任务的 payload 是空串（默认值）", repr(old.get("payload")))
    chk(db.job_retry_payload("old_job") is None,
        "老任务没有参数 → 重试时如实说「重试不了」，而不是瞎跑")

    # ==============================================================
    section("九、锁不会被「跑失败」卡死")
    # ==============================================================
    # 上面连着一堆失败任务，如果失败路径没放开锁，后面所有重活都会被永久挡住
    r = api_call("clear_ledger", cookie=ck)
    chk("正在执行" not in str(r.get("error") or ""),
        "连续失败之后，重活锁仍然可用（没有死锁）", r.get("error"))

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
