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

    print("\n[7] 登录页文案必须跟代码一致（界面说的 ≠ 程序做的 是最难查的那种坑）")
    # 代码里：只用访问口令登录 → _after_login("口令用户", "admin") —— 是管理员。
    # 登录页原来却写着「没有账号就只填访问口令（只读）」，跟程序反着说，
    # 用户照着理解就会以为「必须用账号密码登录」，然后拿一个非管理员账号
    # 进去发现「设置页里账号与权限整块没了」。
    web = read(ROOT / "app_web.py") if (ROOT / "app_web.py").exists() else ""
    login = re.search(r'LOGIN_PAGE\s*=\s*"""(.*?)"""', web, re.S)
    if not login:
        problems.append("在 app_web.py 里找不到 LOGIN_PAGE，没法核对登录页文案")
        print("    ❌ 找不到 LOGIN_PAGE")
    else:
        page = login.group(1)
        if "只读" in page:
            problems.append("登录页还写着访问口令是「只读」，但代码给的是管理员（界面与程序不一致）")
            print("    ❌ 登录页仍写着访问口令「只读」")
        else:
            print("    ✅ 登录页没有「访问口令＝只读」这种与代码相反的说法")
        if "管理员" in page:
            print("    ✅ 登录页写明了「只填访问口令＝管理员」")
        else:
            problems.append("登录页没说明「只填访问口令就是管理员」，用户不知道该往哪登录")
            print("    ❌ 登录页没说访问口令＝管理员")
        if "口令用户" in web and "role=\"admin\"" in web:
            print("    ✅ 代码里访问口令那条路确实给的是 admin（文案与代码同向）")
        else:
            problems.append("代码里「口令登录＝管理员」这条不见了，登录页文案要跟着改")
            print("    ❌ 代码里找不到「口令登录＝admin」")

    print("\n[8] 排障日志链 + 按钮权限（用户报「删除报错、容器日志啥也没有」）")
    # 那次的两个根因都不是「后端坏了」：
    #   ① 三个日志闸门各管一段（_log_line 只写文件、Api._log 只进内存+stdout、
    #      do_POST 里的异常三个都不进）→ docker logs 永远空白；
    #   ② delete_rows/clear_ledger 只有 admin 能调，可那两颗红按钮对所有角色都显示。
    # 这两条一旦被改回去，用户又会掉进「点了报错但查不到」的坑里，所以写死在这里守着。
    dockerfile = read(ROOT / "Dockerfile") if (ROOT / "Dockerfile").exists() else ""
    app = read(ROOT / "gui" / "app.js") if (ROOT / "gui" / "app.js").exists() else ""

    def need(cond, ok_msg, bad_msg):
        if cond:
            print("    ✅ " + ok_msg)
        else:
            problems.append(bad_msg)
            print("    ❌ " + bad_msg)

    need("def _force_utf8_streams" in web,
         "启动时把 stdout/stderr 强制成 UTF-8（否则中文日志被编码异常吞掉＝什么都没打）",
         "少了 _force_utf8_streams：容器 locale 是 C 时中文日志会静默丢失")
    need('f.write(line + "\\n")' in web and "_safe_print(line)" in web,
         "_log_line 同时写文件 + stdout（docker logs 才看得到）",
         "_log_line 又只写文件了 —— docker logs 会变回一片空白")
    need("_push_log(msg" in web and '"seq": _LOG_SEQ' in web,
         "日志同时进界面「运行记录」（不用 SSH 就能看）",
         "日志没接进界面运行记录（poll 还在读旧的 self._logs）")
    need("def _log_api_call" in web and "self._log_api_call(name, error, t0, tb)" in web,
         "每个接口调用都留一行（接口名/身份/角色/耗时/成败）",
         "do_POST 没有逐调用记日志：出错了查不到谁在什么时候调了什么")
    need("traceback.format_exc()" in web,
         "接口抛异常时打完整堆栈",
         "异常没有堆栈：只有一句 error，定位不到哪一行")
    need("QUIET_API" in web,
         "高频轮询（poll / list_jobs）静音，不刷屏",
         "没有静音名单：poll 每 0.4s 一条会把真正要看的东西冲走")
    need("def _allowed_methods" in web and '"allowed": _allowed_methods(r)' in web,
         "whoami 下发完整权限表（界面据此压灰按钮）",
         "whoami 没下发 allowed：界面没法判断哪些按钮该灰")
    need("function applyPermissions" in app and "PERM_BTNS" in app
         and "'btn-del-sel'" in app and "'delete_rows'" in app,
         "界面按权限压灰「删除选中」这类按钮",
         "界面没按权限压灰按钮：业务员会看到点了必然报错的「删除选中」")
    need("FB_LOG_LEVEL" in dockerfile and "PYTHONIOENCODING=utf-8" in dockerfile,
         "Dockerfile 里设了 FB_LOG_LEVEL / PYTHONIOENCODING",
         "Dockerfile 没设 FB_LOG_LEVEL / PYTHONIOENCODING：容器日志要么没有、要么乱码")

    print("\n[9] 金额口径只有一处 + 票面姓名保护（用户报「4361 加了两遍打车、高铁票名字识别不到」）")
    # 同样不是「后端算错了」，而是**同一个数在两边各算一遍、算法还不一样**：
    #   ① 台账页的合计自己把每行硬加，把附件（打车行程单）也加进去了；
    #      生成单据那边是排除附件的 → 同一份数据两个数（4361.00 / 4284.90）；
    #   ② 两个文件夹的票（太仓批 / 烟台批）被加成一个数 → 现在按批次分组下发；
    #   ③ 高铁票的姓名票面上**读得完全正确**，是被「设置出行人」批量填名字盖掉了。
    # 这三条被改回去任何一条，用户又会看到「金额对不上 / 名字全是我自己输的」。
    ledger_src = read(ROOT / "ledger.py") if (ROOT / "ledger.py").exists() else ""
    need('"sum_raw"' in web and "attach_seqs" in web and "def _by_batch" in web,
         "get_ledger 同时给「计费合计」+「全部行硬加」+「按批次分组」",
         "get_ledger 又只给一个合计：界面无从解释「为什么和你算的差一张行程单」")
    need("def _face_names" in web and "protect=True" in web and "tv.SRC_FACE" in web,
         "save_traveler_review 默认保护票面来源的姓名（要覆盖得二次确认）",
         "票面姓名保护没了：批量填名字又会把票面上白纸黑字写的姓名一起盖掉")
    need("function isAttach" in app and "function renderBatchBar" in app
         and "function askFaceOverwrite" in app,
         "界面三件都在：附件不计钱 / 按批次分开报 / 覆盖前二次确认",
         "界面少了「附件不计钱」「批次条」「覆盖确认」中的某一项")
    need("let modalCancel = null" in app and "modalCancel = o.onCancel" in app,
         "弹层支持 onCancel 回调（关掉弹层时把界面状态收拾干净）",
         "showModal 不支持 onCancel：用户点「保持票面姓名」后界面会和台账对不上")
    need("_export_excel_soft" in ledger_src and "def excel_stale" in ledger_src,
         "Excel 快照导出失败不再让整个操作报错（库才是唯一可信源）",
         "Excel 被占用时又会报「文件正被占用」：其实库已经改好了，用户以为白改了一遍")

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
