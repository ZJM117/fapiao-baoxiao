# -*- coding: utf-8 -*-
"""
端到端冒烟测试（全程无窗口、不碰真实台账）
==========================================
做法：
  1. 把整个项目复制到 %TEMP%\\invrp_e2e_xxx\\app 里跑 —— 台账、输出、缓存全在临时目录，
     用户真实的「发票台账.xlsx / 输出 / config.json」一个字节都不会动。
  2. 从真实发票归档里拷 5 个 PDF + 3 个 zip 到临时「发票文件夹」，当作待入库的样本
     （只读源文件，不修改）。
  3. 起 app_web 的本地 HTTP 服务（不调 main()，所以不会打开浏览器），
     照界面的调用顺序把 15 个接口跑一遍：
        get_config → set_config → check_source → get_paths → poll → get_ledger
        → start_import（后台，等 poll 推 result）→ get_ledger
        → mark_rows → update_rows → make_report → merge_invoices(2up/plain)
        → export_summary → 再 get_ledger 核对 → GET /index.html
  4. 逐项打印 ✅/❌，最后给结论；退出时删掉临时目录。

用法：python 测试/diag_e2e.py
"""
import json
import os
import shutil
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve()
ROOT = HERE.parent.parent

REAL_SRC = Path(r"E:\桌\某某公司\开具的发票")

FAILS = []
STEPS = []


def check(name, ok, detail=""):
    STEPS.append((name, bool(ok), detail))
    print(f"  {'✅' if ok else '❌'} {name}" + (f"　{detail}" if detail else ""))
    if not ok:
        FAILS.append(name)
    return ok


def post(base, name, *args):
    body = json.dumps({"args": list(args)}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(f"{base}/api/{name}", data=body,
                                headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as r:
        j = json.loads(r.read().decode("utf-8"))
    if j.get("error"):
        raise RuntimeError(f"{name} 后端报错：{j['error']}")
    return j.get("result")


def pick_samples(dst: Path):
    """从真实归档里挑几个有代表性的文件（PDF 拿出来单独用，zip 用来验证 XML 解析）"""
    pdfs, zips = [], []
    for p in sorted(REAL_SRC.rglob("*")):
        if not p.is_file():
            continue
        if p.suffix.lower() == ".pdf" and len(pdfs) < 5 and p.parent not in [x.parent for x in pdfs]:
            pdfs.append(p)
        elif p.suffix.lower() == ".zip" and len(zips) < 3:
            zips.append(p)
        if len(pdfs) >= 5 and len(zips) >= 3:
            break
    dst.mkdir(parents=True, exist_ok=True)
    n = 0
    for p in pdfs + zips:
        try:
            shutil.copy2(p, dst / p.name)
            n += 1
        except OSError as e:
            print(f"    （跳过 {p.name}：{e}）")
    return n


def main():
    print("=" * 60)
    print("发票报销工具 · 端到端冒烟测试（隔离环境）")
    print("=" * 60)

    tmp = Path(tempfile.mkdtemp(prefix="invrp_e2e_"))
    appdir = tmp / "app"
    srcdir = tmp / "待报销发票"
    print(f"\n临时目录：{tmp}")

    try:
        # ---------- 准备隔离环境 ----------
        shutil.copytree(ROOT, appdir, ignore=shutil.ignore_patterns(
            "__pycache__", "缓存", "台账备份", "输出", "browser_debug",
            "*.xlsx", "config.json", "*.log", "ui_port.txt", "*.bat"))
        sys.path.insert(0, str(appdir))
        os.chdir(appdir)
        print(f"已复制项目到：{appdir}")

        n = pick_samples(srcdir)
        print(f"已从真实归档拷入 {n} 个样本文件 → {srcdir}")
        if n < 4:
            print("❌ 样本文件不足，测试无法继续")
            return 1

        # ---------- 起服务 ----------
        import app_web
        api = app_web.Api()
        srv = app_web.start_server(api, appdir / "gui", 0)
        port = srv.server_address[1]
        base = f"http://127.0.0.1:{port}"
        print(f"本地服务：{base}\n")

        # ---------- 1. 配置 / 路径 ----------
        print("[1] 配置与路径")
        cfg = post(base, "get_config")
        check("get_config 返回配置", isinstance(cfg, dict) and "categories" in cfg,
              f"{len(cfg.get('categories') or [])} 个费用类别")
        cfg = post(base, "set_config", {"source": str(srcdir), "recursive": True,
                                       "people": "张三", "dept": "审计部"})
        check("set_config 写入成功", cfg.get("source") == str(srcdir))
        paths = post(base, "get_paths")
        check("get_paths 返回四个路径", all(k in paths for k in ("source", "ledger", "output", "app")))
        check("台账文件指向临时目录", str(tmp) in paths["ledger"], paths["ledger"])

        # ---------- 2. 统计待入库文件 ----------
        print("\n[2] 目录检查")
        cs = post(base, "check_source", str(srcdir))
        check("check_source 找到文件", cs.get("ok") and cs.get("files", 0) >= 4,
              f"{cs.get('files')} 个　{cs.get('kinds')}")
        bad = post(base, "check_source", str(tmp / "不存在的目录"))
        check("不存在的目录能正确报警", bad.get("ok") is False, bad.get("msg", ""))

        # ---------- 3. 入库建账（后台任务 + poll 推结果）----------
        print("\n[3] 入库建账")
        led0 = post(base, "get_ledger", {})
        check("空台账读取正常", led0.get("rows") == [] and led0.get("total") == 0)

        started = post(base, "start_import", True, "", "张三", "9月现场审计")
        check("start_import 启动成功", started is True)

        result = None
        t0 = time.time()
        log_lines = []
        while time.time() - t0 < 180:
            poll = post(base, "poll", 0)
            log_lines = poll.get("logs") or log_lines
            if poll.get("result"):
                result = poll["result"]
                break
            if not poll.get("busy") and time.time() - t0 > 3:
                break
            time.sleep(0.4)
        check("后台任务通过 poll 推回结果", result is not None, f"耗时 {time.time() - t0:.1f}s")
        if result is None:
            print("    最近日志：")
            for x in log_lines[-8:]:
                print("      " + x["m"][:110])
            return 1
        check("入库没有报错", not result.get("error"), str(result.get("error") or ""))
        check("解析出的发票数 > 0", (result.get("added", 0) + result.get("in_ledger", 0)) > 0,
              f"新增 {result.get('added')} 张，已有 {result.get('in_ledger')} 张，"
              f"需人工 {result.get('manual')} 个，重复 {result.get('dup')} 张")
        check("返回了需人工明细字段", "manual_list" in result and "dup_list" in result)
        for m in result.get("manual_list") or []:
            print(f"      · 需人工：{m['name'][:56]}")

        # 结果要能「再看一次」——界面刷新后不该丢（靠 result_id 去重，而不是推完就清）
        poll2 = post(base, "poll", 0)
        check("刷新后仍能拿到上次入库结果（result_id 未变）",
              poll2.get("result") is not None and poll2.get("result_id") == poll.get("result_id"),
              f"result_id={poll2.get('result_id')}")

        # ---------- 4. 台账查询 / 筛选 / 汇总 ----------
        print("\n[4] 台账")
        led = post(base, "get_ledger", {})
        rows = led.get("rows") or []
        check("台账有数据", len(rows) > 0, f"{len(rows)} 行，合计 {led.get('sum')}")
        check("台账字段齐全", all(k in rows[0] for k in
              ("序号", "状态", "发票号码", "开票日期", "价税合计", "费用类别", "文件路径")))
        check("带回了费用类别/月份候选",
              bool(led.get("categories")) and bool(led.get("months")),
              f"类别 {led.get('categories')[:4]}… 月份 {led.get('months')[:4]}")
        agg = led.get("aggregates") or {}
        check("汇总含 5 个维度", all(k in agg for k in ("按月", "按费用类别", "按销方", "按项目", "按状态")),
              "、".join(f"{k}{len(v)}类" for k, v in agg.items()))
        sample_no = rows[0]["发票号码"]
        kw = post(base, "get_ledger", {"kw": sample_no})
        check("关键词筛选生效", kw.get("total", 0) >= 1, f"命中 {kw.get('total')} 行")

        # ---------- 5. 改状态 / 改类别 ----------
        print("\n[5] 台账操作")
        seqs = [r["序号"] for r in rows[:min(3, len(rows))]]
        r1 = post(base, "mark_rows", seqs, "已报销", "", "张三")
        check("mark_rows 标已报销", r1.get("updated") == len(seqs), f"更新 {r1.get('updated')} 行")
        done = post(base, "get_ledger", {"status": "已报销"})
        check("按状态筛选能查到已报销", done.get("total") == len(seqs), f"{done.get('total')} 行")
        r2 = post(base, "update_rows", [seqs[0]], {"费用类别": "办公费"})
        check("update_rows 改费用类别", r2.get("updated") == 1)

        # ---------- 6. 出报销单 ----------
        print("\n[6] 报销单 / 合并")
        meta = {"kind": "费用报销单", "person": "张三", "dept": "审计部",
                "reason": "项目现场审计差旅费", "date": "2026年09月15日"}
        rep = post(base, "make_report", meta, seqs)
        out = Path(str(rep.get("out") or ""))
        check("生成报销单 PDF", out.exists() and out.stat().st_size > 1000,
              f"{out.name}（{out.stat().st_size if out.exists() else 0} 字节）"
              f"　合计 {rep.get('sum')}　大写 {rep.get('upper')}")

        # 合并：只挑文件路径确实是 PDF 的那几行
        pdf_rows = [r for r in rows if str(r.get("文件路径", "")).lower().endswith(".pdf")]
        if pdf_rows:
            m_seqs = [r["序号"] for r in pdf_rows[:4]]
            m1 = post(base, "merge_invoices", m_seqs, "2up")
            o1 = Path(str(m1.get("out") or ""))
            check("合并（A4 上下两张）", o1.exists() and m1.get("pages", 0) >= 1,
                  f"{o1.name}　{m1.get('pages')} 页 / {m1.get('files')} 张")
            m2 = post(base, "merge_invoices", m_seqs, "plain")
            o2 = Path(str(m2.get("out") or ""))
            check("合并（每页一张）", o2.exists() and m2.get("pages", 0) >= 1,
                  f"{o2.name}　{m2.get('pages')} 页")
        else:
            check("合并测试（需要 PDF 样本）", False, "样本里没有 PDF，跳过")

        # ---------- 7. 导出统计 ----------
        print("\n[7] 导出统计")
        ex = post(base, "export_summary", {})
        exp = Path(str(ex.get("out") or ""))
        check("导出统计 Excel", exp.exists() and exp.stat().st_size > 1000,
              f"{exp.name}（{ex.get('rows')} 行）")
        import openpyxl
        wb = openpyxl.load_workbook(str(exp))
        check("Excel 含多张 sheet", len(wb.sheetnames) >= 5, "、".join(wb.sheetnames))
        ex2 = post(base, "export_summary", {"status": "已报销"})
        check("导出的筛选口径和界面一致", ex2.get("rows") == len(seqs),
              f"筛「已报销」导出 {ex2.get('rows')} 行（界面同样条件查到 {done.get('total')} 行）")

        # ---------- 8. 静态资源 ----------
        print("\n[8] 界面资源")
        with urllib.request.urlopen(f"{base}/index.html", timeout=10) as r:
            html = r.read().decode("utf-8")
        check("GET /index.html 正常", r.status == 200 and "发票报销工具" in html,
              f"{len(html)} 字节")
        for f in ("app.js", "style.css", "extra.css", "bridge.js"):
            with urllib.request.urlopen(f"{base}/{f}", timeout=10) as rr:
                check(f"GET /{f}", rr.status == 200, f"{rr.headers.get('Content-Length') or '?'} 字节")

        # ---------- 收尾 ----------
        print("\n[9] 收尾")
        api.shutdown()
        srv.shutdown()
        check("shutdown 能关掉服务", api.exit_flag.is_set())

    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        print(f"\n已清理临时目录：{tmp}")

    print("\n" + "=" * 60)
    total = len(STEPS)
    bad = len(FAILS)
    print(f"结果：{total - bad} / {total} 项通过")
    if bad:
        print("未通过：")
        for f in FAILS:
            print("  ❌ " + f)
        return 1
    print("✅ 全链路端到端通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
