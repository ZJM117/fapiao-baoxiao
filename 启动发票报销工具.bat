@echo off
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
