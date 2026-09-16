# -*- coding: utf-8 -*-
"""
服务器模式（Docker / NAS）隔离测试
==================================
全程无窗口、不碰真实台账：用 %TEMP% 下的副本目录 + 子进程起真实服务。

覆盖：
  ① 环境变量生效（FB_SERVER / FB_HOST / FB_PORT / FB_DATA_DIR / FB_SOURCE / FB_PASSWORD）
  ② 访问口令：未登录挡下来、口令不对不放行、登录后放行、退出后再次被挡
  ③ /files/ 下发：列目录、下文件、路径穿越被拦
  ④ 服务器模式下生成报销单 PDF（无头浏览器那条路没坏）
  ⑤ 本机模式（不设环境变量）行为不变：只听 127.0.0.1、不启用鉴权
"""
import io
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
PY = sys.executable

fails = []


def chk(cond, label, extra=""):
    print(("  ok   " if cond else "  FAIL ") + label + ("" if cond else "  <<< " + str(extra)))
    if not cond:
        fails.append(label)


def free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def hdr(headers, key, default=""):
    """HTTP 头大小写不敏感地取值"""
    for k, v in headers.items():
        if k.lower() == key.lower():
            return v
    return default


def http(base, path, method="GET", data=None, cookie=None, timeout=15):
    """请求一次，**不自动跟随 302**（302 本身就是我们要断言的东西）"""

    class NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, *a, **k):
            return None

    body = None
    req = urllib.request.Request(base + path, method=method)
    if data is not None:
        body = urllib.parse.urlencode(data).encode("utf-8")
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
    if cookie:
        req.add_header("Cookie", cookie)
    opener = urllib.request.build_opener(NoRedirect)
    try:
        r = opener.open(req, body, timeout=timeout)
        return r.status, dict(r.headers), r.read()
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers), e.read()
    except Exception as e:
        return 0, {}, str(e).encode()


def wait_up(base, proc, log_path, tries=140):
    for _ in range(tries):
        if proc.poll() is not None:
            txt = Path(log_path).read_text(encoding="utf-8", errors="replace")
            raise RuntimeError("子进程提前退出了：\n" + txt[-1800:])
        code, _, _ = http(base, "/api/ping")
        if code == 200:
            return True
        time.sleep(0.25)
    return False


TMP = Path(tempfile.mkdtemp(prefix="srvmode_"))
DATA = TMP / "data"
SRC = TMP / "票据"
for d in (DATA, DATA / "缓存", DATA / "输出", DATA / "台账备份", SRC):
    d.mkdir(parents=True, exist_ok=True)
(DATA / "输出" / "样例报销单.pdf").write_bytes(b"%PDF-1.4\n% fake for download test\n")
(SRC / "一张发票.pdf").write_bytes(b"%PDF-1.4\n% fake invoice\n")
(TMP / "secret.txt").write_text("TOP SECRET - should never be downloadable", encoding="utf-8")

PW = "k7m3x9qd"
PORT = free_port()
BASE = f"http://127.0.0.1:{PORT}"
SRV_LOG = TMP / "srv.log"

print("=" * 74)
print(f"临时根目录 {TMP}")
print(f"服务地址   {BASE}")
print("=" * 74)

env = dict(os.environ)
env.update({"FB_SERVER": "1", "FB_HOST": "127.0.0.1", "FB_PORT": str(PORT),
            "FB_DATA_DIR": str(DATA), "FB_SOURCE": str(SRC), "FB_PASSWORD": PW})

logf = open(SRV_LOG, "wb")
proc = subprocess.Popen([PY, str(ROOT / "app_web.py")], cwd=str(ROOT), env=env,
                        stdout=logf, stderr=subprocess.STDOUT)

try:
    print("\n— ① 服务器模式起服务 —")
    up = wait_up(BASE, proc, SRV_LOG)
    chk(up, "服务起来了（/api/ping 免鉴权就能通，main() 靠它判重）")
    if not up:
        raise SystemExit(1)
    # 日志落在数据目录里（容器里就挂在卷上，重启不丢）
    LOG = DATA / "launch.log"
    log = LOG.read_text(encoding="utf-8", errors="replace")
    chk("服务器（Docker/NAS）" in log, "日志里写明运行模式是服务器", log[-400:])
    chk(f"监听 127.0.0.1:{PORT}" in log, "日志里的监听地址端口是环境变量给的那个")
    chk(str(DATA) in log, "数据目录用的是 FB_DATA_DIR 指定的临时目录")
    chk("访问口令：已启用" in log, "日志里提示已启用口令")
    chk(LOG.parent == DATA, "日志写在数据目录里（容器挂卷 → 重启不丢）")
    chk("不自动打开浏览器" in log, "服务器模式不自动开浏览器（容器里也没有浏览器可开）")
    # docker logs 看到的就是进程的标准输出，得让用户一眼知道该访问什么地址
    sout = SRV_LOG.read_text(encoding="utf-8", errors="replace")
    chk("[就绪] 请用浏览器访问" in sout, "控制台（docker logs）里给出了访问地址", sout[-300:])

    print("\n— ② 访问口令 —")
    code, h, body = http(BASE, "/index.html")
    chk(code == 302 and hdr(h, "Location") == "/login", "没登录访问页面 → 302 跳登录页", (code, hdr(h, "Location")))
    code, h, body = http(BASE, "/api/poll", "POST", {}, )
    chk(code == 401, "没登录调接口 → 401（不是页面跳转）", code)
    code, h, body = http(BASE, "/login")
    chk(code == 200 and b'name="password"' in body, "登录页能打开且有输入框", code)
    code, h, body = http(BASE, "/login", "POST", {"password": "wrong-pass"})
    chk(code == 200 and "不对".encode() in body, "口令不对 → 回登录页并提示", code)
    chk("登录失败" in LOG.read_text(encoding="utf-8", errors="replace"), "失败记进日志了")
    code, h, body = http(BASE, "/login", "POST", {"password": PW})
    setck = hdr(h, "Set-Cookie")
    chk(code == 302 and hdr(h, "Location") == "/index.html", "口令对 → 跳首页", (code, hdr(h, "Location")))
    chk("fb_auth=" in setck and "HttpOnly" in setck, "发了 HttpOnly 的登录 Cookie", setck)
    ck = setck.split(";")[0]
    code, h, body = http(BASE, "/index.html", cookie=ck)
    chk(code == 200 and b"<!doctype html" in body.lower(), "带着 Cookie 就能打开界面了", code)
    code, h, body = http(BASE, "/api/poll", "POST", {}, cookie=ck)
    chk(code == 200 and b'"result"' in body, "带着 Cookie 能调接口", code)
    code, h, body = http(BASE, "/logout", cookie=ck)
    chk(code == 302 and hdr(h, "Location") == "/login", "退出 → 回登录页", (code, hdr(h, "Location")))
    chk("Max-Age=0" in hdr(h, "Set-Cookie"), "退出时把 Cookie 清掉了", hdr(h, "Set-Cookie"))
    code, h, body = http(BASE, "/index.html", cookie=ck)
    chk(code == 302, "退出后这个会话立刻失效（不是等浏览器自己丢 Cookie）", code)

    print("\n— ③ /files/ 文件下发 —")
    code, h, body = http(BASE, "/files/", cookie=ck)
    chk(code == 302 and hdr(h, "Location") == "/login",
        "/files/ 也要登录（没登录被挡去登录页）", (code, hdr(h, "Location")))
    # 重新登录拿一个有效的 Cookie，继续往下测
    _, h2, _ = http(BASE, "/login", "POST", {"password": PW})
    ck = hdr(h2, "Set-Cookie").split(";")[0]
    code, h, body = http(BASE, "/files/", cookie=ck)
    txt = body.decode("utf-8", "replace")
    chk(code == 200 and "output/" in txt and "data/" in txt and "source/" in txt,
        "列出可浏览的目录别名（output / data / source）", txt[:200])
    code, h, body = http(BASE, "/files/output/", cookie=ck)
    txt = body.decode("utf-8", "replace")
    chk(code == 200 and "样例报销单.pdf" in txt, "输出目录列出来了（含中文名）", txt[:200])
    code, h, body = http(BASE, "/files/output/" + urllib.parse.quote("样例报销单.pdf"), cookie=ck)
    chk(code == 200 and body.startswith(b"%PDF"), "能把输出目录里的 PDF 下下来", code)
    chk(hdr(h, "Content-Disposition") == "",
        "PDF 走内联（浏览器里直接看，不强制下载）", hdr(h, "Content-Disposition"))
    code, h, body = http(BASE, "/files/source/" + urllib.parse.quote("一张发票.pdf"), cookie=ck)
    chk(code == 200 and body.startswith(b"%PDF"), "票据文件夹里的文件也能看", code)
    code, h, body = http(BASE, "/files/data/%2e%2e%2fsecret.txt", cookie=ck)
    chk(code == 403, "路径穿越（%2e%2e%2f）被拦成 403", code)
    chk(b"TOP SECRET" not in body, "穿越没把秘密文件漏出去")
    code, h, body = http(BASE, "/files/nope/", cookie=ck)
    chk(code == 404, "不认识的别名 → 404", code)

    print("\n— ④ 服务器模式下生成报销单（无头浏览器那条路）—")
    gen = TMP / "gen.py"
    gen.write_text(f'''
# -*- coding: utf-8 -*-
import io, os, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = r"{ROOT}"
sys.path.insert(0, ROOT)
import report_pdf as rp
from app_paths import DATA_DIR, EDGE, SERVER_MODE

print("SERVER_MODE =", SERVER_MODE)
print("EDGE =", EDGE, os.path.exists(EDGE))

def R(seq, ctype, no, date, trip, total, who):
    return {{"序号": str(seq), "状态": "未报销", "凭证类型": ctype, "票面标记": "",
            "出行人": who, "发票号码": no, "开票日期": date, "票种": "",
            "行程/明细": trip, "购方名称": "国网湖北省电力有限公司", "销方名称": "某某公司",
            "不含税金额": "", "税额": "", "价税合计": total, "项目/事由": "",
            "项目编号": "", "费用类别": "交通费", "报销人": "张三", "报销批次": "",
            "入库时间": "2026-09-16 10:00", "文件路径": "", "提示": ""}}

rows = [R(1, "火车票", "", "2026-08-17", "G1234  济南东-烟台  06:12  二等座", 174.00, "李强"),
        R(2, "打车行程单", "", "2026-08-18", "济南西-大明湖", 76.10, "李强"),
        R(3, "发票", "26349119423005870002", "2026-08-20", "", 208.00, "李贵清")]
meta = {{"kind": "费用报销单", "person": "张三", "reason": "服务器模式自检",
        "date": "2026年09月16日", "start": "2026-08-17", "end": "2026-08-20"}}
outs = rp.make_report(rows, meta, str(DATA_DIR / "输出" / "服务器模式样张"), fmt="both")
for p in outs:
    print("OUT", os.path.basename(p), os.path.getsize(p))
''', encoding="utf-8")

    r = subprocess.run([PY, str(gen)], cwd=str(ROOT), env=env, capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    out = (r.stdout or "") + (r.stderr or "")
    chk("SERVER_MODE = True" in out, "子进程确实在服务器模式下", out[-400:])
    pdfs = [l for l in out.splitlines() if l.startswith("OUT ") and ".pdf" in l]
    chk(len(pdfs) >= 1 and int(pdfs[0].split()[2]) > 5000,
        "服务器模式下报销单 PDF 生成成功", pdfs or out[-400:])
    print("      " + "; ".join(pdfs))

finally:
    try:
        proc.terminate()
        proc.wait(timeout=10)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass
    logf.close()

print("\n— ⑤ 本机模式（不设环境变量）行为不变 —")
os.environ.pop("FB_SERVER", None)
os.environ.pop("FB_HOST", None)
os.environ.pop("FB_PASSWORD", None)
os.environ.pop("FB_DATA_DIR", None)
os.environ.pop("FB_SOURCE", None)
sys.path.insert(0, str(ROOT))
import app_paths                                                    # noqa: E402
app_paths.CONFIG_FILE = TMP / "cfg.json"                            # 隔离，别读真实配置
import app_web                                                      # noqa: E402

chk(app_paths.SERVER_MODE is False, "不设 → 不是服务器模式")
chk(app_paths.HOST == "127.0.0.1", "不设 → 只监听 127.0.0.1（外网访问不到）", app_paths.HOST)
chk(app_paths.PORT == 8766, "不设 → 端口还是 8766", app_paths.PORT)
chk(Path(app_paths.DATA_DIR) == Path(app_paths.APP_DIR), "不设 → 数据就放在程序目录（老行为）")
chk(app_paths.PASSWORD == "", "不设 → 没有口令")
chk(app_web._auth_on() is False, "本机模式不启用登录校验")
chk(app_web.UI_PORT == 8766, "UI_PORT 仍取 8766", app_web.UI_PORT)

srv = app_web.start_server(app_web.Api(), app_paths.GUI_DIR, 0)
port2 = srv.server_address[1]
host2 = srv.server_address[0]
chk(host2 == "127.0.0.1", "start_server 绑的是 127.0.0.1", host2)
b2 = f"http://127.0.0.1:{port2}"
code, h, body = http(b2, "/index.html")
chk(code == 200, "本机模式打开界面不用登录（跟以前一模一样）", code)
code, h, body = http(b2, "/api/poll", "POST", {})
chk(code == 200, "本机模式调接口不用登录", code)
srv.shutdown()

shutil.rmtree(TMP, ignore_errors=True)

# ---------------------------------------------------------------- ⑥ 部署文件自检
print("\n— ⑥ Docker 部署文件 —")
DOCKERFILE = ROOT / "Dockerfile"
COMPOSE = ROOT / "docker-compose.yml"
IGNORE = ROOT / ".dockerignore"
GUIDE = ROOT / "Docker部署说明.md"

chk(DOCKERFILE.exists() and COMPOSE.exists() and IGNORE.exists() and GUIDE.exists(),
    "四个部署文件都在（Dockerfile / docker-compose.yml / .dockerignore / 说明）",
    [p.name for p in (DOCKERFILE, COMPOSE, IGNORE, GUIDE) if not p.exists()])

if DOCKERFILE.exists():
    df = DOCKERFILE.read_text(encoding="utf-8")
    chk("chromium" in df, "装了 chromium（PDF 打印靠它）")
    chk("fonts-noto-cjk" in df, "装了中文字体（不然报销单中文全是方块）")
    chk("FB_SERVER=1" in df or "FB_SERVER=1 \\" in df, "镜像里默认开服务器模式")
    chk("FB_HOST=0.0.0.0" in df, "默认绑 0.0.0.0（不然容器外访问不到）")
    chk("FB_DATA_DIR=/data" in df, "数据默认落 /data 卷")
    chk('VOLUME ["/data"]' in df, "声明了 /data 卷")
    chk("EXPOSE 8766" in df, "对外暴露 8766")
    chk("openpyxl" in df and "pypdf" in df and "pdfminer" in df,
        "三个 Python 依赖都装了")

if COMPOSE.exists():
    cp = COMPOSE.read_text(encoding="utf-8")
    chk("FB_PASSWORD" in cp, "compose 里留了口令这一项")
    chk("8766:8766" in cp, "端口映射 8766")
    chk("/data" in cp and "/票据" in cp, "两个卷都挂上了（数据 + 票据）")
    chk("restart:" in cp, "配了自动重启")
    try:
        import yaml                                                # noqa: E402
        y = yaml.safe_load(cp)
        svc = list((y.get("services") or {}).values())[0]
        chk(bool(svc.get("volumes")) and bool(svc.get("ports")),
            "compose 能被解析，且 volumes / ports 写全了", svc.keys())
    except ImportError:
        print("  ℹ️  没装 pyyaml，跳过 YAML 解析检查（文本检查已经过了）")

if IGNORE.exists():
    ig = IGNORE.read_text(encoding="utf-8")
    chk("发票台账.xlsx" in ig, ".dockerignore 排掉了真实台账（别打进镜像）")
    chk("config.json" in ig, ".dockerignore 排掉了本机配置")
    chk("测试/" in ig, ".dockerignore 排掉了测试目录")

print()
if fails:
    print(f"❌ {len(fails)} 项不通过：")
    for f in fails:
        print("   -", f)
    sys.exit(1)
print("✅ 全部通过")
