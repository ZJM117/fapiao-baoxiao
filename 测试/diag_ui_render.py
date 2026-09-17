# -*- coding: utf-8 -*-
"""
界面渲染自检（无头 Edge，全程无窗口）
====================================
静态自检（diag_ui_check.py）只能保证「id/class/变量名对得上」，
但 JS 里写错一个运行期细节（比如 document.querySelector 返回 null 后取属性）
只有真跑一遍浏览器才看得出来。

做法：
  1. 在进程内起 app_web 的本地服务（不调 main()，不打开浏览器）；
  2. 用 msedge --headless=new --dump-dom 把界面加载一遍 —— JS 会真的执行，
     然后把它跑完之后的 DOM 打出来；
  3. 在 DOM 里找「只有 JS 跑成功才会出现的痕迹」：
        · 12 张主题卡片（renderThemes 跑过了）
        · 日志里有「界面已就绪」（boot + pushLogs 跑过了）
        · html[data-theme] 被设过（applyTheme 跑过了）
        · body 有 native-frame（boot 跑过了）
  4. 同时抓 stderr 里的 JS 报错关键字。

⚠️ 只读：不写任何项目文件（Api() 只读 config.json，服务只读 gui/）。
用法：python 测试/diag_ui_render.py
"""
import re
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve()
ROOT = HERE.parent.parent
sys.path.insert(0, str(ROOT))


def main():
    import app_web
    from app_paths import EDGE

    if not Path(EDGE).exists():
        print(f"❌ 没找到 Edge：{EDGE}")
        return 1

    api = app_web.Api()
    srv = app_web.start_server(api, ROOT / "gui", 0)
    port = srv.server_address[1]
    url = f"http://127.0.0.1:{port}/index.html"
    print(f"本地服务：{url}\n")

    out = ""
    err = ""
    tmp = tempfile.mkdtemp(prefix="invrp_uirender_")
    try:
        cmd = [EDGE, "--headless=new", "--disable-gpu", "--no-first-run",
               "--no-default-browser-check", "--disable-extensions",
               f"--user-data-dir={tmp}", "--enable-logging=stderr", "--v=0",
               "--virtual-time-budget=3000", "--dump-dom", url]
        p = subprocess.run(cmd, capture_output=True, timeout=120)
        out = p.stdout.decode("utf-8", "replace")
        err = p.stderr.decode("utf-8", "replace")
    finally:
        srv.shutdown()

    checks = []
    def ck(name, ok, detail=""):
        checks.append((name, bool(ok)))
        print(f"  {'✅' if ok else '❌'} {name}" + (f"　{detail}" if detail else ""))

    print(f"DOM {len(out)} 字节，stderr {len(err)} 字节\n")
    print("[1] JS 执行痕迹")
    ck("页面标题正确", "发票报销工具" in out)
    n_theme = out.count('class="theme-card"')
    ck("4 张主题卡片已渲染", n_theme == 4, f"实际 {n_theme} 张")
    n_swatch = out.count('class="theme-swatch"')
    # ⚠️ 别把带反斜杠的引号写进 f-string 的表达式里：Python 3.11 及更早会直接 SyntaxError
    ck("主题卡片带预览色块", n_swatch == 12, f"{n_swatch} 个色块")
    ck("applyTheme 已生效（html 带 data-theme）", bool(re.search(r'<html[^>]*data-theme="', out)))
    ck("boot 已生效（body 带 native-frame）", 'class="native-frame"' in out)
    ck("日志已写入（boot 里的就绪提示）", "界面已就绪" in out,
       (re.search(r'id="log-count"[^>]*>([^<]*)<', out) or ["", "?"])[1])

    print("\n[2] 各页面结构")
    # 「凭证台账」+「报销单」已在 2026-09-17 合并成「报销作业」（page-ledger）一页
    for pid, label in [("page-in", "发票入库"), ("page-ledger", "报销作业"),
                       ("page-stats", "统计汇总"), ("page-jobs", "任务中心"),
                       ("page-log", "运行记录"), ("page-set", "设置")]:
        ck(f"{label} 页存在", f'id="{pid}"' in out)
    ck("合并后的作业页里有出单区", 'id="work-panel"' in out)
    ck("入库页是默认激活页", re.search(r'class="page active" id="page-in"', out) is not None)
    ck("导航项数量正确（6 个）", out.count('class="nav-item') >= 6)

    print("\n[3] JS 报错")
    bad = re.findall(r"(Uncaught[^\n]*|ReferenceError[^\n]*|TypeError[^\n]*|SyntaxError[^\n]*)", err)
    # Edge 自己的一些噪音（GPU/DevTools 等）不算
    bad = [b for b in bad if "WebGL" not in b and "GPU" not in b]
    ck("没有 JS 运行时报错", not bad, ("；".join(bad[:3]))[:200])

    ok = all(v for _, v in checks)
    print("\n" + "=" * 56)
    print(f"结果：{sum(1 for _, v in checks if v)} / {len(checks)} 项通过")
    print("✅ 界面在真实浏览器里渲染正常" if ok else "❌ 有问题，见上面")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
