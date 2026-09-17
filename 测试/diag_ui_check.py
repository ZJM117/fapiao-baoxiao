# -*- coding: utf-8 -*-
"""
界面静态自检（只读，不启动任何窗口）
====================================
1. app.js 里所有 $('xxx') 引用的元素 id，必须在 index.html 里存在
   （写错一个字母的后果是「点按钮没反应」，最难查，所以单独扫一遍）
2. index.html 里用到的 class，必须在 style.css 或 extra.css 里有定义
3. extra.css 里引用的 CSS 变量，必须在 style.css 里有定义
4. index.html 引用的 js/css 文件都得存在

用法：python 测试/diag_ui_check.py
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
GUI = ROOT / "gui"

# 由 JS 动态建出来的元素（modal 里内联生成的），本来就不在 index.html 里。
# ⚠️ 放进这里只代表「允许不在 index.html」，仍会去 app.js 里核对
#    `id="xxx"` 这个字面量真的存在 —— 否则 $('xxx') 与生成的 id 对不上就是死按钮。
DYNAMIC_IDS = {"ledger-stats", "tv-input", "tvr-box", "tvr-bulk"}


def read(p):
    return p.read_text(encoding="utf-8")


def main():
    html = read(GUI / "index.html")
    appjs = read(GUI / "app.js")
    style = read(GUI / "style.css")
    extra = read(GUI / "extra.css")

    problems = []

    # ---------- 1. id ----------
    html_ids = set(re.findall(r'\bid="([^"]+)"', html))
    js_ids = set(re.findall(r"\$\('([^']+)'\)", appjs))
    missing = sorted(js_ids - html_ids - DYNAMIC_IDS)
    unused = sorted(html_ids - js_ids)
    print(f"[1] index.html 里定义了 {len(html_ids)} 个 id；app.js 引用了 {len(js_ids)} 个")
    if missing:
        problems.append(f"app.js 引用了 index.html 里不存在的 id：{missing}")
        print("    ❌ 缺失:", missing)
    else:
        print("    ✅ app.js 引用的 id 全部存在")
    if unused:
        print(f"    ℹ️  index.html 里没被 app.js 用到的 id（可能正常）：{unused}")

    # 动态 id 也要对得上：$('tv-input') 必须能在 app.js 里找到 id="tv-input"，
    # 或 holder.id = 'ledger-stats' 这种赋值写法。对不上就是死按钮。
    def _has_dyn_id(name: str) -> bool:
        return bool(re.search(rf"""id\s*=\s*['"]{re.escape(name)}['"]""", appjs))

    bad_dyn = sorted(i for i in DYNAMIC_IDS if not _has_dyn_id(i))
    if bad_dyn:
        problems.append(f"动态 id 在 app.js 里找不到对应的 id=\"…\" 字面量：{bad_dyn}")
        print("    ❌ 动态 id 对不上:", bad_dyn)
    else:
        print(f"    ✅ {len(DYNAMIC_IDS)} 个动态 id 都能在 app.js 里找到对应的 id=\"…\"")

    # ---------- 2. class ----------
    html_classes = set()
    for attr in re.findall(r'\bclass="([^"]*)"', html):
        html_classes.update(attr.split())
    css_classes = set()
    for src in (style, extra):
        css_classes.update(re.findall(r"\.([A-Za-z_][\w-]*)", src))
    no_style = sorted(c for c in html_classes if c not in css_classes)
    print(f"\n[2] index.html 用了 {len(html_classes)} 个 class；样式表里定义了 {len(css_classes)} 个")
    if no_style:
        print(f"    ⚠️ 没有样式定义的 class：{no_style}")
    else:
        print("    ✅ 所有 class 都有样式定义")

    # ---------- 3. CSS 变量 ----------
    used_vars = set(re.findall(r"var\(\s*(--[\w-]+)", extra))
    defined_vars = set(re.findall(r"(--[\w-]+)\s*:", style))
    # style.css 里也允许有 fallback 形式
    bad = sorted(v for v in used_vars if v not in defined_vars)
    print(f"\n[3] extra.css 引用 {len(used_vars)} 个变量；style.css 定义了 {len(defined_vars)} 个")
    if bad:
        problems.append(f"extra.css 引用了 style.css 里没有的变量：{bad}")
        print("    ❌ 未定义:", bad)
    else:
        print("    ✅ extra.css 用到的变量全部有定义")

    # ---------- 4. 资源存在 ----------
    print("\n[4] index.html 引用的资源：")
    ok = True
    for m in re.findall(r'<(?:link[^>]+href|script[^>]+src)="([^"]+)"', html):
        if m.startswith(("http:", "https:", "//")):
            continue
        exists = (GUI / m).exists()
        print(f"    {'✅' if exists else '❌'} {m}")
        ok = ok and exists
    if not ok:
        problems.append("index.html 引用了不存在的资源文件")

    # ---------- 5. 服务器模式（Docker/NAS）的前端契约 ----------
    # 容器里开不了资源管理器，后端会回 {url}，由 bridge.js 替它开新标签页。
    # 这几条一旦被改掉，NAS 上「打开/定位」就会静默失效（点了没反应），很难查。
    print("\n[5] 服务器模式相关的前端契约：")
    bridge = read(GUI / "bridge.js")
    contracts = [
        (bridge, r"window\.open\(\s*r\.url",
         "bridge.js：后端给了 url 就用浏览器打开"),
        (bridge, r"open_file',\s*'open_folder',\s*'reveal'",
         "bridge.js：包装的是 open_file / open_folder / reveal 这三个方法"),
        (appjs, r"path\.server",
         "app.js：「选择文件夹」在服务器模式下改提示手填路径"),
        (appjs, r"call\('pick_folder'",
         "app.js：仍然通过 pick_folder 走后端"),
    ]
    for src, pat, label in contracts:
        hit = bool(re.search(pat, src))
        print(f"    {'✅' if hit else '❌'} {label}")
        if not hit:
            problems.append(label + " —— 这条前端契约没了")

    # ---------- 6. 主题变量：颜色类必须四套齐全 ----------
    # :root 装的是深色那一套（也是所有变量的底），浅色主题靠 [data-theme="…"] 覆盖。
    # **只要漏写一条颜色变量，那条在浅色主题下就会漏出深色的值** ——
    # 2026-09-17「凭证明细表头压着一条黑带」就是 --table-header-bg 漏写造成的
    # （退回 :root 的 #0a1521）。这条断言专门拦它再次发生。
    root_block = re.search(r':root\s*\{(.*?)\n\}', style, re.S).group(1)
    root_vars = dict(re.findall(r'(--[\w-]+)\s*:\s*([^;]+);', root_block))
    # 与主题无关的设计令牌（圆角 / 字体 / 间距 / 动效）、以及 var() 引用式变量，不要求每套重写
    theme_indep = re.compile(r'^--(r|r-sm|r-lg|r-xl|font|font-mono|sp-|transition|glass-blur)')
    color_keys = sorted(k for k, v in root_vars.items()
                        if not theme_indep.match(k) and not v.strip().startswith('var('))
    print(f"\n[6] 主题变量完整性：:root 共 {len(root_vars)} 条，其中「随主题变色」的 {len(color_keys)} 条")
    print("    （:root = 深色；浅色主题只要漏一条，那条就会漏出深色值）")
    # 只认「真的跟着 { 的」那种 [data-theme="…"]，注释里提到的不算
    for name in sorted(set(re.findall(r'\[data-theme="([^"]+)"\]\s*\{', style))):
        blk = re.search(r'\[data-theme="' + re.escape(name) + r'"\]\s*\{(.*?)\n\}', style, re.S)
        have = set(re.findall(r'(--[\w-]+)\s*:', blk.group(1))) if blk else set()
        miss = [k for k in color_keys if k not in have]
        if miss:
            problems.append(
                f'[data-theme="{name}"] 漏了颜色变量 {miss} —— '
                f"这几条在浅色主题下会漏出 :root 的深色值")
            print(f"    ❌ {name}: 漏 {len(miss)} 条 -> {miss}")
        else:
            print(f"    ✅ {name}: {len(color_keys)} 条颜色变量齐全")

    print("\n" + "=" * 56)
    if problems:
        print("发现问题：")
        for p in problems:
            print("  ❌ " + p)
        return 1
    print("✅ 静态自检全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
