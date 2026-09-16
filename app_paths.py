# -*- coding: utf-8 -*-
"""
统一路径配置
============
所有路径集中在这里，整个项目文件夹可以随便搬位置，程序自动跟随。

同一套代码两种跑法：
  ① 本机双击 bat —— 数据就放在程序目录里（跟以前完全一样），只监听 127.0.0.1
  ② Docker/NAS  —— 程序在镜像里（只读），数据落到挂载卷上，监听 0.0.0.0

环境变量（都有默认值，不设就是第 ① 种）：
  FB_SERVER=1     服务器模式：空闲不自动退出、不自动开浏览器、默认绑 0.0.0.0
  FB_HOST         监听地址（默认 127.0.0.1；FB_SERVER=1 时默认 0.0.0.0）
  FB_PORT         监听端口（默认 8766）
  FB_DATA_DIR     数据目录：台账/配置/缓存/输出/回收站都放这儿
  FB_SOURCE       票据文件夹（容器里通常挂到 /票据）
  FB_PASSWORD     访问口令，留空=不校验（局域网直接进）
  FB_CHROME       Chromium/Edge 可执行文件（打印 PDF 用，自动找得到就不用设）
"""
import os
import shutil
import sys
from pathlib import Path


def _env(name: str, default: str = "") -> str:
    return (os.environ.get(name) or "").strip() or default


def _flag(name: str) -> bool:
    return (os.environ.get(name) or "").strip().lower() in ("1", "true", "yes", "on")


def app_dir() -> Path:
    """程序所在目录：打包成 exe 时为 exe 所在目录；源码运行时为脚本目录"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# ---------- 运行模式 ----------
SERVER_MODE = _flag("FB_SERVER")


# ---------- 程序自身 ----------
APP_DIR = app_dir()

# ---------- 数据目录 ----------
# 本机跑：数据就在程序目录（老行为，双击 bat 一切照旧）。
# 容器跑：镜像里的 /app 只读，数据必须落到挂载卷 → FB_DATA_DIR=/data。
DATA_DIR = Path(_env("FB_DATA_DIR") or APP_DIR).expanduser()

CONFIG_FILE = DATA_DIR / "config.json"        # 用户设置（默认发票文件夹、报销人、费用类别…）
LEDGER_FILE = DATA_DIR / "发票台账.xlsx"       # 台账本体（一个 Excel 文件，双击就能看/改）
LEDGER_BAK = DATA_DIR / "台账备份"
RUN_LOG = DATA_DIR / "run.log"                # 操作日志
LAUNCH_LOG = DATA_DIR / "launch.log"          # 启动日志（排障用）
UI_PORT_FILE = DATA_DIR / "ui_port.txt"       # 界面服务端口记录
WORK_DIR = DATA_DIR / "缓存"                   # 中间产物（报销单 HTML 等），可随时删
DEBUG_DIR = DATA_DIR / "browser_debug"
# 容器里没有 Windows 回收站，删掉的东西挪进这里（同样是「删错能捞」的语义）
RECYCLE_DIR = DATA_DIR / "回收站"

# ---------- 票据与输出 ----------
# 默认票据文件夹：用户自己整理的、要报销的票据都放一个文件夹里，程序递归遍历。
_SRC_ENV = _env("FB_SOURCE")
_SOURCE_CANDIDATES = ([Path(_SRC_ENV)] if _SRC_ENV else []) + [
    Path(r"E:\桌\湖北报销发票整理"),
    Path(r"E:\桌\某某公司\开具的发票"),
]
# 环境变量指了一个还不存在的目录（容器首次启动挂载前）也算数，别再退到 Windows 路径
DEFAULT_SOURCE = next((p for p in _SOURCE_CANDIDATES if p.exists()),
                      _SOURCE_CANDIDATES[0] if _SOURCE_CANDIDATES else DATA_DIR)
# 生成物（报销单 PDF、合并 PDF）放这里，不放票据文件夹里，免得下次遍历又把它们当票据
OUTPUT_DIR = DATA_DIR / "输出"

# ---------- 监听 ----------
HOST = _env("FB_HOST") or ("0.0.0.0" if SERVER_MODE else "127.0.0.1")
try:
    PORT = int(_env("FB_PORT", "8766"))
except ValueError:
    PORT = 8766
PASSWORD = _env("FB_PASSWORD")          # 非空才启用登录校验


# ---------- 界面 ----------
def _gui_dir() -> Path:
    """界面目录：优先用 exe/脚本旁边的 gui/（方便直接改界面），否则用打包内置的"""
    external = APP_DIR / "gui"
    if external.exists():
        return external
    base = getattr(sys, "_MEIPASS", None)
    if base:
        packed = Path(base) / "gui"
        if packed.exists():
            return packed
    return external


GUI_DIR = _gui_dir()


# ---------- 浏览器（生成 PDF 用无头浏览器打印，免装 PDF 排版库）----------
# 本机是 Edge；容器里装的是 Chromium —— 命令行参数一样，都是 Chromium 内核。
def _find_browser() -> str:
    env = _env("FB_CHROME")
    if env and Path(env).exists():
        return env
    if os.name == "nt":
        cands = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                 r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                 r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                 r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    else:
        cands = ["/usr/bin/chromium", "/usr/bin/chromium-browser",
                 "/usr/bin/google-chrome", "/usr/bin/google-chrome-stable",
                 "/usr/lib/chromium/chromium", "/opt/google/chrome/chrome"]
    for p in cands:
        if Path(p).exists():
            return p
    for name in ("chromium", "chromium-browser", "google-chrome", "msedge"):
        found = shutil.which(name)
        if found:
            return found
    return cands[0]                     # 都没有也返回一个名字，报错信息里能看出找的是什么


EDGE = _find_browser()                  # 名字保留（历史代码都引用它）
BROWSER = EDGE

# ---------- 默认设置 ----------
DEFAULTS = {
    "source": str(DEFAULT_SOURCE),      # 票据文件夹（可改，界面里选）
    "recursive": True,                  # 是否递归遍历子文件夹
    "people": "",                       # 默认报销人
    "dept": "",                         # 默认部门
    "categories": ["交通费", "住宿费", "餐饮费", "办公费", "会议费", "培训费", "招待费", "其他"],
    "skip_archived_batches": True,      # 遍历时跳过「已报销」历史，避免重复入库
}

# 只建数据目录：容器里 APP_DIR（/app）是只读的，别去碰它
for _d in (DATA_DIR, OUTPUT_DIR, WORK_DIR, LEDGER_BAK, DEBUG_DIR, RECYCLE_DIR):
    try:
        _d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
