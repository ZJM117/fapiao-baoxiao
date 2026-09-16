# -*- coding: utf-8 -*-
"""
「上传本机文件夹」隔离测试（服务器模式 / NAS）
==============================================
全程无窗口、不碰真实台账：%TEMP% 下建一套副本目录 + 子进程起真实服务，
用真实的 HTTP 请求走一遍上传，最后只看临时目录里的文件长成什么样。

覆盖：
  ① upload_info：服务器模式报出落点、单文件上限、可写性
  ② 落盘：中文名 / 多层子目录原样保留 —— 发票和它的行程单必须还在同一个文件夹里
  ③ 重复上传同一个文件 → skip（不重写，mtime 不动）
  ④ 同名但内容不同 → 存成 -2，两份都在（覆盖比留两份危险）
  ⑤ 越界（../、..\\）被拒；空文件被拒；超过单文件上限被拒
  ⑥ 未登录不让上传（跟其它接口一个待遇）
  ⑦ 落到「上传」目录之后再上传，不会变成「上传/上传」（不越挖越深）
  ⑧ 传上去的文件能从 /files/source/ 下载回来（服务器模式下「打开文件」走这条）
  ⑨ 界面侧静态检查：按钮平时藏着、靠 upload_info().server 决定要不要露出来
"""
import base64
import io
import json
import os
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
    for k, v in headers.items():
        if k.lower() == key.lower():
            return v
    return default


# ---------------------------------------------------------------- 临时环境
TMP = Path(tempfile.mkdtemp(prefix="upload_"))
DATA = TMP / "data"
SRC = TMP / "票据"
for d in (DATA, DATA / "缓存", DATA / "输出", DATA / "台账备份", SRC):
    d.mkdir(parents=True, exist_ok=True)

PW = "up9x2k7q"
PORT = free_port()
BASE = f"http://127.0.0.1:{PORT}"
UP = SRC / "上传"                     # 上传落点（票据目录/上传）
SRV_LOG = TMP / "srv.log"
MAX_BYTES = 512 * 1024 * 1024


def http_get(path, cookie=""):
    req = urllib.request.Request(BASE + path)
    if cookie:
        req.add_header("Cookie", cookie)
    try:
        r = urllib.request.urlopen(req, timeout=20)
        return r.status, r.read(), dict(r.headers)
    except urllib.error.HTTPError as e:
        return e.code, e.read(), dict(e.headers)
    except Exception as e:
        return 0, str(e).encode(), {}


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """登录成功是 302 —— 不能被自动跟走，否则拿不到那个 Set-Cookie"""
    def redirect_request(self, *a, **k):
        return None


_OPENER = urllib.request.build_opener(_NoRedirect)


def api(name, *args, cookie="", method="POST", form=None):
    """调一个后端接口，返回 (response, raw_text)；HTTP 错误码拿 .code 看"""
    req = urllib.request.Request(BASE + ("/login" if form is not None else "/api/" + name),
                                 method=method)
    if form is not None:
        body = urllib.parse.urlencode(form).encode("utf-8")
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
    else:
        body = json.dumps({"args": list(args)}).encode("utf-8")
        req.add_header("Content-Type", "application/json")
    if cookie:
        req.add_header("Cookie", cookie)
    try:
        r = _OPENER.open(req, body, timeout=30)
        return r, r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        return e, e.read().decode("utf-8", "replace")


def call(name, *args, cookie=""):
    r, raw = api(name, *args, cookie=cookie)
    code = getattr(r, "status", getattr(r, "code", 0))
    if code != 200:
        return None, f"HTTP {code}"
    j = json.loads(raw)
    return j.get("result"), j.get("error")


def upload(rel, data, cookie="", code_expect=200):
    """走真实的上传接口：body = 原始字节，相对路径放 X-Upload-Path 头（base64）"""
    body = data if isinstance(data, bytes) else data.encode("utf-8")
    req = urllib.request.Request(BASE + "/api/upload_raw", method="POST", data=body)
    req.add_header("X-Upload-Path", base64.b64encode(rel.encode("utf-8")).decode("ascii"))
    req.add_header("X-File-Size", str(len(body)))
    req.add_header("Content-Type", "application/octet-stream")
    if cookie:
        req.add_header("Cookie", cookie)
    try:
        r = urllib.request.urlopen(req, timeout=60)
        code, raw = r.status, r.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        code, raw = e.code, e.read().decode("utf-8", "replace")
    if code != code_expect:
        return code, {"_raw": raw}
    try:
        return code, json.loads(raw)
    except ValueError:
        return code, {"_raw": raw}


def raw_upload(rel_b64, decl_len, body=b"", cookie=""):
    """手写一条 HTTP 请求 —— 用来伪造超大的 Content-Length（urllib 不让这么干）"""
    s = socket.create_connection(("127.0.0.1", PORT), timeout=15)
    head = ("POST /api/upload_raw HTTP/1.0\r\nHost: 127.0.0.1\r\n"
            + (f"Cookie: {cookie}\r\n" if cookie else "")
            + f"X-Upload-Path: {rel_b64}\r\nContent-Length: {decl_len}\r\n\r\n")
    s.sendall(head.encode("ascii"))
    if body:
        s.sendall(body)
    s.settimeout(6)
    buf = b""
    try:
        while True:
            chunk = s.recv(4096)
            if not chunk:
                break
            buf += chunk
    except (socket.timeout, ConnectionResetError):
        pass
    s.close()
    return buf.decode("utf-8", "replace")


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


def wait_up(tries=160):
    for _ in range(tries):
        if proc.poll() is not None:
            raise RuntimeError("子进程提前退出：\n"
                               + SRV_LOG.read_text(encoding="utf-8", errors="replace")[-1800:])
        code, _, _ = http_get("/api/ping")
        if code == 200:
            return True
        time.sleep(0.25)
    return False


cookie = ""
try:
    if not wait_up():
        raise SystemExit("服务没起来")

    print("\n— ① 登录与 upload_info —")
    r, _ = api("", form={"password": PW}, method="POST")
    cookie = hdr(dict(getattr(r, "headers", {})), "Set-Cookie").split(";")[0]
    chk(cookie.startswith("fb_auth="), "登录拿到会话 Cookie", cookie)

    info, err = call("upload_info", cookie=cookie)
    chk(err is None and info, "upload_info 能调通", err)
    chk(info.get("server") is True, "服务器模式下 server=True（界面靠它决定露不露上传按钮）", info)
    chk(info.get("root") == str(UP), "落点 = 票据目录/上传", info.get("root"))
    chk(info.get("base") == str(SRC), "基底 = 票据目录（FB_SOURCE）", info.get("base"))
    chk(info.get("ok") is True, "落点可写", info.get("msg"))
    chk(info.get("max_size") == MAX_BYTES, "单文件上限报出来了", info.get("max_size"))

    print("\n— ② 落盘：中文名 + 多层子目录原样保留 —")
    A = "差旅发票：郝琳-武汉.pdf\n%PDF-1.4 AAAA".encode("utf-8")
    B = "滴滴行程单-9月3日.ofd\n%OFD BBBB".encode("utf-8")
    C = "住宿专用发票.ofd\n%OFD CCCC".encode("utf-8")
    c1, j1 = upload("郝琳9月武汉/9月/差旅发票：郝琳-武汉.pdf", A, cookie)
    c2, j2 = upload("郝琳9月武汉/9月/滴滴行程单-9月3日.ofd", B, cookie)
    c3, j3 = upload("郝琳9月武汉/10月/住宿专用发票.ofd", C, cookie)
    chk(j1.get("result", {}).get("mode") == "new", "第 1 个文件是新增", j1)
    chk(j2.get("result", {}).get("mode") == "new", "第 2 个文件是新增", j2)
    chk(j3.get("result", {}).get("mode") == "new", "第 3 个文件是新增", j3)
    f1 = UP / "郝琳9月武汉" / "9月" / "差旅发票：郝琳-武汉.pdf"
    f2 = UP / "郝琳9月武汉" / "9月" / "滴滴行程单-9月3日.ofd"
    f3 = UP / "郝琳9月武汉" / "10月" / "住宿专用发票.ofd"
    chk(f1.is_file() and f1.read_bytes() == A, "文件落在 票据目录/上传/<你选的文件夹>/9月/ 下，内容一致", f1)
    chk(f2.is_file() and f2.read_bytes() == B, "行程单也在同一个文件夹（配对规则才认得出它俩是一对）",
        f2.parent == f1.parent)
    chk(f3.is_file() and f3.read_bytes() == C, "深层子目录（10月）也原样保留", f3)
    chk(j1["result"]["rel"] == "郝琳9月武汉/9月/差旅发票：郝琳-武汉.pdf",
        "返回的相对路径是给界面看的那个（带原目录结构）", j1["result"]["rel"])
    chk(j1["result"]["root"] == str(UP), "回包里带落点，界面据此把路径框切过去", j1["result"]["root"])
    chk(j1["result"]["size"] == len(A), "回包里的字节数等于发出去的量", j1["result"]["size"])

    print("\n— ③ 重复上传同一个文件 → skip —")
    t_before = f1.stat().st_mtime_ns
    time.sleep(0.05)
    c4, j4 = upload("郝琳9月武汉/9月/差旅发票：郝琳-武汉.pdf", A, cookie)
    chk(j4.get("result", {}).get("mode") == "skip", "同名同大小 → skip（不重复写）", j4)
    chk(f1.stat().st_mtime_ns == t_before, "磁盘上的文件确实没被重写（mtime 没动）", t_before)
    chk(len(list((UP / "郝琳9月武汉" / "9月").iterdir())) == 2, "目录里没有多出第二份",
        [p.name for p in (UP / "郝琳9月武汉" / "9月").iterdir()])

    print("\n— ④ 同名但内容不同 → 存成 -2，两份都留 —")
    D = "差旅发票：郝琳-武汉.pdf（改过的版本，只多了这一行）".encode("utf-8")
    c5, j5 = upload("郝琳9月武汉/9月/差旅发票：郝琳-武汉.pdf", D, cookie)
    f1b = UP / "郝琳9月武汉" / "9月" / "差旅发票：郝琳-武汉-2.pdf"
    chk(j5.get("result", {}).get("mode") == "new", "内容不一样 → 当新文件收下", j5)
    chk(f1b.is_file() and f1b.read_bytes() == D, "新那份存成 -2", f1b)
    chk(f1.read_bytes() == A, "原来那份没被覆盖（覆盖比留两份危险）", f1.read_bytes()[:20])

    print("\n— ⑤ 越界 / 空文件 / 超大文件 —")
    up_before = sorted(p.name for p in TMP.iterdir())
    c6, j6 = upload("../secret.txt", b"x", cookie)
    chk("不合法" in str(j6.get("error", "")), "../ 被拒（不合法）", j6)
    c7, j7 = upload("a/../../b.txt", b"x", cookie)
    chk("不合法" in str(j7.get("error", "")), "a/../../b.txt 被拒", j7)
    c8, j8 = upload("..\\..\\win.txt", b"x", cookie)
    chk("不合法" in str(j8.get("error", "")), "反斜杠 ..\\..\\ 也被拒", j8)
    c9, j9 = upload("/etc/passwd", b"x", cookie)
    chk(j9.get("result", {}) and "passwd" in str(j9["result"].get("rel", "")),
        "绝对路径不打到头上去，只当成相对路径收下", j9)
    chk((UP / "etc" / "passwd").is_file(), "它落在上传目录里面（跑不出 root）")
    chk(sorted(p.name for p in TMP.iterdir()) == up_before,
        "临时根目录没有被多写出文件（没有任何一个落到票据目录外面）",
        sorted(p.name for p in TMP.iterdir()))

    c10, j10 = upload("空文件.pdf", b"", cookie)
    chk("空文件" in str(j10.get("error", "")), "空文件被拒", j10)

    rel_b64 = base64.b64encode("超大.pdf".encode("utf-8")).decode("ascii")
    raw = raw_upload(rel_b64, MAX_BYTES + 1, b"x" * 16, cookie)
    chk("超过单文件上限" in raw, "超过单文件上限被拒（超大 body 也不硬收）", raw[-300:])
    chk(not (UP / "超大.pdf").exists(), "被拒的文件没有留下半个残骸")

    print("\n— ⑥ 未登录不让上传 —")
    c11, j11 = upload("没登录.pdf", b"x", code_expect=401)
    chk(c11 == 401, "不带 Cookie 上传 → 401", c11)
    chk(not (UP / "没登录.pdf").exists(), "未登录的文件一个字节都没落盘")

    print("\n— ⑦ 落到「上传」目录后再传，不会越挖越深 —")
    res, err = call("set_config", {"source": str(UP)}, cookie=cookie)
    chk(err is None, "把票据目录设成「上传」目录（界面传完就是这么切的）", err)
    info2, _ = call("upload_info", cookie=cookie)
    chk(info2.get("root") == str(UP), "落点仍然是 票据目录/上传（没有变成 上传/上传）", info2.get("root"))
    c12, j12 = upload("第二批/新票.pdf", b"%PDF-1.4 EEEE", cookie)
    chk((UP / "第二批" / "新票.pdf").is_file(), "第二批落在 上传/第二批/ 下", j12)
    chk(not (UP / "上传").exists(), "没有出现 上传/上传 这种套娃目录")

    print("\n— ⑧ 传上去的文件能从 /files/source/ 下载回来 —")
    code, body, _ = http_get("/files/source/" + urllib.parse.quote("郝琳9月武汉/9月/差旅发票：郝琳-武汉.pdf"),
                             cookie=cookie)
    chk(code == 200, "/files/source/... 取得到（服务器模式下「打开文件」走这条）", code)
    chk(body == A, "下载回来的内容和传上去的一模一样", body[:20])
    code2, _, _ = http_get("/files/source/" + urllib.parse.quote("uploading-not-exist.txt"), cookie=cookie)
    chk(code2 == 404, "不存在的文件是 404", code2)

finally:
    proc.terminate()
    try:
        proc.wait(timeout=8)
    except subprocess.TimeoutExpired:
        proc.kill()
    logf.close()

print("\n— ⑨ 界面侧静态检查（本机模式不摆这个按钮） —")
appjs = (ROOT / "gui" / "app.js").read_text(encoding="utf-8")
html = (ROOT / "gui" / "index.html").read_text(encoding="utf-8")
bridge = (ROOT / "gui" / "bridge.js").read_text(encoding="utf-8")
chk("if (!info || !info.server) { row.hidden = true; return; }" in appjs,
    "app.js：不是服务器模式就把上传按钮藏起来（能直接指路径，不必多一份副本）")
chk("webkitdirectory" in html, "index.html：用 webkitdirectory 选文件夹（局域网 http 下也能用）")
chk('id="up-row" hidden' in html, "index.html：上传行默认藏着，等确认是服务器模式再露")
chk("'X-Upload-Path'" in appjs, "app.js：相对路径走 X-Upload-Path 头（中文路径 base64 编码）")
chk("upload_info" in bridge, "bridge.js：接口清单里登记了 upload_info")
chk("'/api/upload_raw'" in appjs, "app.js：文件直传 /api/upload_raw")

print("\n" + "=" * 74)
if fails:
    print(f"❌ 失败 {len(fails)} 项：")
    for f in fails:
        print("   - " + f)
else:
    print("✅ 全部通过")
print(f"临时目录（可留可删）：{TMP}")
print("=" * 74)
sys.exit(1 if fails else 0)
