# -*- coding: utf-8 -*-
"""桌面文件夹选择器由独立 STA 进程承载，不依赖 Tk 主线程消息泵。"""
import base64
import sys
import threading
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import app_web

FAIL = []
DIAG_LOGS = []
app_web._log_line = lambda message, *_args, **_kwargs: DIAG_LOGS.append(str(message))


def chk(cond, label, extra=""):
    print(("  [OK] " if cond else "  [XX] ") + label
          + ("" if cond else f"   → {extra}"))
    if not cond:
        FAIL.append(label)


api = app_web.Api.__new__(app_web.Api)
api._log = lambda message: DIAG_LOGS.append(str(message))

# Server mode must keep its existing manual-path behavior and never create Tk.
app_web.SERVER_MODE = True
server_result = api.pick_folder("")
chk(server_result.get("server") is True, "服务器模式返回手动填写提示")

app_web.SERVER_MODE = False
calls = []
selected_path = r"D:\选择结果"


def fake_picker(initial):
    calls.append((initial, threading.current_thread()))
    return selected_path
result = {}


def request_picker():
    try:
        result["path"] = api.pick_folder(r"D:\初始目录")
    except Exception as e:  # noqa: BLE001
        result["error"] = repr(e)


worker = threading.Thread(target=request_picker, daemon=True)
with patch.object(app_web, "_windows_folder_dialog", side_effect=fake_picker):
    worker.start()
    worker.join(timeout=3)

chk(not worker.is_alive(), "HTTP 请求线程在选择器完成后正常返回")
chk(result.get("path") == selected_path, "选择结果返回给调用方", result)
chk(calls == [(r"D:\初始目录", worker)],
    "选择器在独立调用线程运行且收到了初始目录", calls)
chk(selected_path not in "\n".join(DIAG_LOGS),
    "诊断日志不记录用户选择的实际路径")
chk(any("Windows STA 文件夹选择器" in msg for msg in DIAG_LOGS),
    "诊断日志记录进入 Windows 选择器阶段", DIAG_LOGS)

encoded_result = base64.b64encode(selected_path.encode("utf-8")).decode("ascii")
with patch.object(app_web.shutil, "which", return_value="powershell.exe"), \
        patch.object(app_web.subprocess, "run",
                     return_value=SimpleNamespace(
                         returncode=0, stdout=f"FB_FOLDER:{encoded_result}\n", stderr="")) as run:
    returned = app_web._windows_folder_dialog(r"D:\初始目录")
    call_args, kwargs = run.call_args
    command = call_args[0]
    script = base64.b64decode(command[-1]).decode("utf-16le")
    chk(returned == selected_path, "PowerShell 返回值正确解码", returned)
    chk("-STA" in command and "FolderBrowserDialog" in script,
        "Windows 对话框运行在 STA 子进程中")
    chk(kwargs.get("timeout") == 90, "选择器有 90 秒超时保护", kwargs)

# check_source must accept the same quoted paths that _import_worker normalizes.
seen = {}


class FakePath:
    def __init__(self, value):
        seen["path"] = value

    def expanduser(self):
        return self

    def exists(self):
        return seen.get("exists", True)

    def is_dir(self):
        return True


fake_invparse = SimpleNamespace(
    collect_files=lambda _path, _recursive: [SimpleNamespace(suffix=".pdf")])
with patch.object(app_web, "Path", FakePath), \
        patch.object(app_web, "load_config", return_value={"recursive": True}), \
        patch.dict(sys.modules, {"invparse": fake_invparse}):
    checked = api.check_source('"C:\\票据目录"')
    chk(checked.get("ok") is True and checked.get("files") == 1,
        "带双引号的手动路径能扫描到支持格式文件", checked)
    chk(seen.get("path") == "C:\\票据目录",
        "扫描接口传给文件系统的是去掉外围引号的路径", seen)
    seen["exists"] = False
    missing = api.check_source("C:\\不存在")
    chk(missing.get("ok") is False and "路径不存在或不是文件夹" in missing.get("msg", ""),
        "无效路径返回可读原因", missing)

saved = {}
with patch.object(app_web, "load_config", return_value={}), \
        patch.object(app_web, "save_config", side_effect=lambda cfg: saved.update(cfg)):
    configured = api.set_config({"source": '"C:\\票据目录"'})
    chk(configured.get("source") == "C:\\票据目录" and saved.get("source") == "C:\\票据目录",
        "配置保存时也去掉路径外围引号", configured)

print("\n" + ("✅ 全部通过" if not FAIL else f"❌ {len(FAIL)} 项不通过"))
sys.exit(1 if FAIL else 0)
