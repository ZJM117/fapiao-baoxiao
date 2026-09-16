# -*- coding: utf-8 -*-
"""
生成「启动发票报销工具.bat」
============================
为什么要用脚本生成而不是直接写文件：

  · Write 工具只能写 UTF-8；而 cmd.exe 默认按 GBK（代码页 936）解析 bat 内容。
    里面只要有中文文件名/中文路径，UTF-8 存盘 → cmd 读出来就是乱码 →
    报「Windows 找不到文件 'XXX'」。本项目所在路径 "E:\\桌\\个人文件同步\\…" 就是中文，
    所以 bat **必须存成 GBK + CRLF**。
  · 这个脚本还兼作「以后要改启动逻辑」的唯一入口：改这里，重新跑一遍就重新生成 bat。

用法：python 测试/make_launcher.py
"""
from pathlib import Path

BAT_NAME = "启动发票报销工具.bat"

# 注意：里面尽量不直接写中文文件名，靠通配符定位 exe，彻底免疫编码问题
BAT = r"""@echo off
setlocal
rem ============================================
rem  发票报销工具 启动器
rem  默认直接跑源码：改完代码存盘、关掉窗口再双击这个 bat 就生效，不用重新打包
rem  %~dp0 = 本文件所在目录，整个文件夹搬到哪都能用
rem  本文件必须是 GBK 编码 + CRLF 换行（见 测试/make_launcher.py 的说明）
rem ============================================
set "APPDIR=%~dp0"
cd /d "%APPDIR%"

rem ---- 1) 优先：用 Python 跑源码 ----
set "PY=C:\Users\JM\.workbuddy\binaries\python\envs\gui\Scripts\pythonw.exe"
if exist "%PY%" if exist "%APPDIR%app_web.py" (
    start "" "%PY%" "%APPDIR%app_web.py"
    exit /b 0
)

rem ---- 2) 兜底：没有 Python 环境时用 exe（通配符定位，避开中文编码问题）----
set "EXE="
for %%F in ("%APPDIR%*.exe") do if exist "%%~fF" if not defined EXE set "EXE=%%~fF"
if defined EXE (
    start "" "%EXE%"
    exit /b 0
)

echo [错误] 既没找到 Python 运行环境，也没找到可执行文件。
echo        请确认 app_web.py 和 gui 文件夹与本 bat 在同一目录。
echo        启动失败详情看同目录的 launch.log
pause
"""


def main():
    root = Path(__file__).resolve().parent.parent
    out = root / BAT_NAME
    # 关键：GBK 编码 + CRLF 换行
    out.write_bytes(BAT.replace("\r\n", "\n").replace("\n", "\r\n").encode("gbk"))
    raw = out.read_bytes()
    crlf = raw.count(b"\r\n")
    lf_only = raw.count(b"\n") - crlf
    print(f"已生成：{out}")
    print(f"大小：{len(raw)} 字节；CRLF {crlf} 处，孤立 LF {lf_only} 处（应为 0）")
    print("前 20 字节：", raw[:20])


if __name__ == "__main__":
    main()
