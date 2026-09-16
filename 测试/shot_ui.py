# -*- coding: utf-8 -*-
"""
界面截图（无头 Edge，全程无窗口）
================================
目的：让用户不用启动程序就能看到界面长什么样；也当作「界面真的能渲染」的证据。

做法：
  1. 复制项目到 %TEMP%，拷真实发票样本，起服务，跑一遍入库 + 标状态，
     让界面有真实数据可显示（全程不碰用户真实台账）；
  2. 用 msedge --headless=new --screenshot 对几个页面各截一张：
        #page-in       发票入库（含入库结果）
        #page-ledger   发票台账
        ?pick=pending#page-report   报销单（自动带入全部未报销发票）
        #page-stats    统计汇总
  3. 把图拷到 <项目>/输出/界面预览/

用法：python 测试/shot_ui.py
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve()
ROOT = HERE.parent.parent
REAL_SRC = Path(r"E:\桌\某某公司\开具的发票")
SHOT_DIR = ROOT / "输出" / "界面预览"

# 这里刻意不 import app_paths —— 必须在 sys.path 指向临时副本之后才 import 项目模块，
# 否则会加载到真实目录的 app_paths，进而把数据写进用户的真实台账。
EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

SHOTS = [
    ("index.html#page-in", "1-发票入库.png"),
    ("index.html#page-ledger", "2-发票台账.png"),
    ("index.html?pick=pending#page-report", "3-报销单.png"),
    ("index.html#page-stats", "4-统计汇总.png"),
]


def post(base, name, *args):
    body = json.dumps({"args": list(args)}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(f"{base}/api/{name}", data=body,
                                headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as r:
        return json.loads(r.read().decode("utf-8")).get("result")


def copy_samples(dst: Path):
    dst.mkdir(parents=True, exist_ok=True)
    got, seen = 0, set()
    for p in sorted(REAL_SRC.rglob("*")):
        if not p.is_file():
            continue
        if p.suffix.lower() not in (".pdf", ".zip"):
            continue
        if p.parent in seen:
            continue
        seen.add(p.parent)
        try:
            shutil.copy2(p, dst / p.name)
            got += 1
        except OSError:
            pass
        if got >= 8:
            break
    return got


def main():
    EDGE = next((p for p in EDGE_CANDIDATES if Path(p).exists()), "")
    if not EDGE:
        print("❌ 没找到 Edge")
        return 1
    tmp = Path(tempfile.mkdtemp(prefix="invrp_shot_"))
    appdir = tmp / "app"
    srcdir = tmp / "待报销发票"
    print(f"临时环境：{tmp}")

    try:
        shutil.copytree(ROOT, appdir, ignore=shutil.ignore_patterns(
            "__pycache__", "缓存", "台账备份", "输出", "browser_debug",
            "*.xlsx", "config.json", "*.log", "ui_port.txt", "*.bat"))
        sys.path.insert(0, str(appdir))
        os.chdir(appdir)

        n = copy_samples(srcdir)
        print(f"样本文件 {n} 个")

        import app_web
        api = app_web.Api()
        srv = app_web.start_server(api, appdir / "gui", 0)
        port = srv.server_address[1]
        base = f"http://127.0.0.1:{port}"
        print(f"服务：{base}")

        post(base, "set_config", {"source": str(srcdir), "recursive": True,
                                 "people": "张三", "dept": "审计部"})
        post(base, "start_import", True, "", "张三", "9月现场审计")
        t0 = time.time()
        while time.time() - t0 < 180:
            r = post(base, "poll", 0)
            if r.get("result"):
                print(f"入库：新增 {r['result'].get('added')} 张，"
                      f"需人工 {r['result'].get('manual')} 个")
                break
            time.sleep(0.4)
        led = post(base, "get_ledger", {}) or {}
        rows = led.get("rows") or []
        if rows:
            seqs = [x["序号"] for x in rows[:max(1, len(rows) // 3)]]
            post(base, "mark_rows", seqs, "已报销", "", "张三")
            post(base, "update_rows", [x["序号"] for x in rows[len(seqs):]],
                 {"费用类别": "咨询审计费"})
        print(f"台账 {len(rows)} 行，合计 {led.get('sum')}")

        SHOT_DIR.mkdir(parents=True, exist_ok=True)
        print("\n开始截图：")
        ok = True
        for rel, fname in SHOTS:
            out = (SHOT_DIR / fname).resolve()
            if out.exists():
                out.unlink()
            cmd = [EDGE, "--headless=new", "--disable-gpu", "--no-first-run",
                   "--no-default-browser-check", "--disable-extensions",
                   f"--user-data-dir={tmp / 'edgeprof'}",
                   "--hide-scrollbars", "--force-device-scale-factor=1",
                   "--window-size=1560,1010",
                   "--virtual-time-budget=4000",
                   f"--screenshot={out}", f"{base}/{rel}"]
            subprocess.run(cmd, capture_output=True, timeout=180,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            size = out.stat().st_size if out.exists() else 0
            good = size > 15000
            ok = ok and good
            print(f"  {'✅' if good else '❌'} {fname}　{size / 1024:.0f} KB")

        # ---------- 顺带留两份「实际产出」样例，方便用户先看版式 ----------
        print("\n生成产出样例：")
        all_seqs = [x["序号"] for x in rows]
        meta = {"kind": "费用报销单", "person": "张三", "dept": "审计部",
                "reason": "项目现场审计差旅费", "date": "2026年09月15日"}
        rep = post(base, "make_report", meta, all_seqs) or {}
        if rep.get("out"):
            tgt = SHOT_DIR / "示例-费用报销单.pdf"
            shutil.copy2(rep["out"], tgt)
            print(f"  ✅ {tgt.name}　{tgt.stat().st_size / 1024:.0f} KB"
                  f"　合计 {rep.get('sum')} = {rep.get('upper')}")
        pdf_seqs = [x["序号"] for x in rows if str(x.get("文件路径", "")).lower().endswith(".pdf")]
        if pdf_seqs:
            mg = post(base, "merge_invoices", pdf_seqs, "2up") or {}
            if mg.get("out"):
                tgt = SHOT_DIR / "示例-发票合并（A4上下两张）.pdf"
                shutil.copy2(mg["out"], tgt)
                print(f"  ✅ {tgt.name}　{tgt.stat().st_size / 1024:.0f} KB　{mg.get('pages')} 页")

        srv.shutdown()
        print(f"\n输出目录：{SHOT_DIR}")
        return 0 if ok else 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
