# -*- coding: utf-8 -*-
"""辅导员工作台 · 双击启动入口（由 PyInstaller 打包为 辅导员工作台.exe）

双击后不再打开浏览器，而是弹出一个独立窗口（和普通桌面软件一样）：
1. 自动定位 exe 所在目录（数据目录），不需要用户配置任何路径；
2. 若 dashboard.html 缺失，先自动生成一次；
3. 在后台启动仅监听 127.0.0.1 的本地服务，并把工作台装进一个原生窗口；
4. 关闭窗口 = 停止服务，不联网、不上传任何数据。

窗口外壳优先级：WebView2（pywebview）→ Edge --app 独立窗口 → 默认浏览器（兜底）。
"""
import os
import re
import sys
import time
import base64
import traceback

from app_paths import APP_HOME

APP_TITLE = "辅导员工作台"
LOG_PATH = os.path.join(APP_HOME, "启动日志.txt")
DASH = os.path.join(APP_HOME, "dashboard.html")


# ---------------------------------------------------------------- 基础工具
def _log(msg):
    """窗口模式下没有控制台，写文件方便排查。"""
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("[%s] %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass
    try:
        if sys.stdout is not None:
            print(msg)
    except Exception:
        pass


def _fix_streams():
    """--noconsole 打包后 sys.stdout 为 None，避免 print 报错。"""
    for name in ("stdout", "stderr"):
        if getattr(sys, name, None) is None:
            try:
                setattr(sys, name, open(LOG_PATH, "a", encoding="utf-8", buffering=1))
            except Exception:
                pass
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _dpi_aware():
    """避免高分屏下窗口发虚。"""
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _icon_path():
    from app_paths import res
    p = res("app.ico", os.path.join("installer", "assets", "app.ico"))
    return p if os.path.exists(p) else None


def _find_hwnd(title, timeout=8.0):
    """按标题找窗口句柄（窗口刚创建时可能还没出来，轮询一会）。"""
    if os.name != "nt":
        return 0
    import ctypes
    u = ctypes.windll.user32
    end = time.time() + timeout
    while time.time() < end:
        h = u.FindWindowW(None, title)
        if h:
            return h
        time.sleep(0.2)
    return 0


def _apply_icon():
    """把窗口图标换成工作台自己的图标（标题栏左上角 + 任务栏）。"""
    ico = _icon_path()
    if os.name != "nt" or not ico:
        return
    try:
        import ctypes
        u = ctypes.windll.user32
        h = _find_hwnd(APP_TITLE)
        if not h:
            return
        IMAGE_ICON, LR_LOADFROMFILE = 1, 0x10
        WM_SETICON, ICON_SMALL, ICON_BIG = 0x0080, 0, 1
        for size, which in ((16, ICON_SMALL), (32, ICON_BIG)):
            hicon = u.LoadImageW(None, ico, IMAGE_ICON, size, size, LR_LOADFROMFILE)
            if hicon:
                u.SendMessageW(h, WM_SETICON, which, hicon)
    except Exception:
        _log("设置窗口图标失败：\n" + traceback.format_exc())


# ---------------------------------------------------------------- 准备数据
def _prepare():
    try:
        os.chdir(APP_HOME)
    except Exception:
        pass

    data_dir = os.path.join(APP_HOME, "data")
    if not os.path.isdir(data_dir):
        # 全新机器首次运行：用 data_template 的空表结构自动初始化
        try:
            import init_data
            init_data.ensure_data()
        except Exception:
            _log("初始化数据目录失败：\n" + traceback.format_exc())
    if not os.path.isdir(data_dir):
        return False

    if not os.path.exists(DASH):
        _log("首次运行，正在生成看板…")
        try:
            import build
            build.main()
        except Exception:
            _log("看板生成失败：\n" + traceback.format_exc())
            return False
    return True


# ---------------------------------------------------------------- 三种窗口外壳
def _downloads_dir():
    """定位系统的「下载」文件夹（中文系统可能叫 Downloads）。"""
    if os.name == "nt":
        try:
            import winreg
            key = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
                p = winreg.QueryValueEx(k, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
            if p and os.path.isdir(p):
                return p
        except Exception:
            pass
    home = os.path.expanduser("~")
    for cand in (os.path.join(home, "Downloads"), os.path.join(home, "下载")):
        if os.path.isdir(cand):
            return cand
    return home


def _safe_name(name):
    """去掉路径成分和非法字符，只留文件名。"""
    name = os.path.basename(str(name or "").replace("\\", "/"))
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name or "导出文件"


class _JsApi:
    """给页面提供的少量本地能力（WebView2 默认禁止浏览器下载，这里直接落盘）。"""

    def save_bytes(self, filename, b64):
        name = _safe_name(filename)
        folder = _downloads_dir()
        try:
            data = base64.b64decode(b64 or "")
        except Exception:
            return {"ok": False, "msg": "数据解码失败"}
        stem, ext = os.path.splitext(name)
        path = os.path.join(folder, name)
        n = 2
        while os.path.exists(path):          # 同名不覆盖，自动加序号
            path = os.path.join(folder, "%s(%d)%s" % (stem, n, ext))
            n += 1
        try:
            with open(path, "wb") as f:
                f.write(data)
        except Exception as e:
            _log("保存文件失败：%s" % e)
            return {"ok": False, "msg": "写入失败：%s" % e}
        _log("已保存：%s" % path)
        return {"ok": True, "path": path, "dir": folder,
                "name": os.path.basename(path), "size": len(data)}

    def open_folder(self, path):
        """打开文件所在文件夹（页面保存成功后提示用）。"""
        try:
            folder = path if os.path.isdir(path) else os.path.dirname(path)
            if folder and os.path.isdir(folder):
                if os.name == "nt":
                    os.startfile(folder)
                return {"ok": True}
        except Exception as e:
            _log("打开文件夹失败：%s" % e)
        return {"ok": False}


def _run_webview(url):
    """首选：WebView2 原生窗口（无地址栏、无标签页，关窗即退出）。"""
    import webview
    try:
        webview.settings["ALLOW_DOWNLOADS"] = True   # 允许页面直接触发的下载
    except Exception:
        pass
    webview.create_window(
        APP_TITLE, url,
        width=1366, height=880, min_size=(1080, 700),
        background_color="#f2f5fa", text_select=True,
        js_api=_JsApi(),
    )
    webview.start(_apply_icon)


def _edge_exe():
    for p in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"):
        if os.path.exists(p):
            return p
    return None


def _run_edge_app(url):
    """次选：用 Edge 的应用窗口模式，视觉上同样没有浏览器界面。"""
    exe = _edge_exe()
    if not exe:
        return False
    import subprocess
    profile = os.path.join(APP_HOME, "webview-profile")
    p = subprocess.Popen([
        exe, "--app=" + url,
        "--user-data-dir=" + profile,
        "--no-first-run", "--no-default-browser-check",
        "--window-size=1366,880",
    ])
    _find_hwnd("辅导员多工作域工作台", timeout=20)
    p.wait()
    return True


def _run_browser(url):
    """兜底：老办法，用默认浏览器打开，并保持服务运行。"""
    import threading
    try:
        import webbrowser
        webbrowser.open(url)
    except Exception:
        pass
    threading.Event().wait()


# ---------------------------------------------------------------- 入口
def _start_url(port):
    """启动页。

    默认先打开 boot.html：它会顺手检查一次更新，发现新版本就下载、静默安装、
    装完自动重启到新版；没有更新就立刻跳进首页，用户基本感觉不到中间这一步。
    自动更新关掉（或没有 boot.html，比如老版本残留目录）时直接进看板。
    """
    try:
        import updater
        if updater.auto_enabled() and os.path.exists(os.path.join(APP_HOME, "boot.html")):
            return "http://127.0.0.1:%d/boot.html" % port
    except Exception:
        _log("自动更新检查不可用，直接进入看板：\n" + traceback.format_exc())
    return "http://127.0.0.1:%d/dashboard.html" % port


def _activate_existing():
    """已有实例在跑时，把它的窗口提到前台，而不是再开一个。"""
    if os.name != "nt":
        return False
    import ctypes
    u = ctypes.windll.user32
    h = _find_hwnd(APP_TITLE, timeout=2.0)
    if not h:
        return False
    try:
        if u.IsIconic(h):
            u.ShowWindow(h, 9)     # SW_RESTORE
        u.SetForegroundWindow(h)
        u.BringWindowToTop(h)
        return True
    except Exception:
        return False


def main():
    _fix_streams()
    _dpi_aware()

    if not _prepare():
        _msg = ("没有找到数据目录，或看板生成失败。\n\n"
                "请确认本程序位于完整的工作台文件夹内，\n"
                "该文件夹里应当有：data（或 data_template）、template.html 等文件。\n\n"
                "详细原因见：" + LOG_PATH)
        _log(_msg)
        if os.name == "nt":
            try:
                import ctypes
                ctypes.windll.user32.MessageBoxW(None, _msg, APP_TITLE, 0x10)
            except Exception:
                pass
        return 1

    import server
    url = _start_url(server.PORT)
    httpd = None
    try:
        httpd = server.start_server()
        if httpd is None:
            _log("检测到已有实例在运行：%s" % url)
            if _activate_existing():
                _log("已把现有窗口带到前台，本次启动不再新开窗口。")
                return 0
            _log("没找到现有窗口，另开一个窗口连过去。")
        else:
            _log("服务已启动：%s" % url)

        try:
            _run_webview(url)
            _log("窗口已关闭，正在退出。")
        except Exception:
            _log("WebView2 窗口不可用，改用 Edge 应用窗口：\n" + traceback.format_exc())
            try:
                if not _run_edge_app(url):
                    raise RuntimeError("未找到 Microsoft Edge")
            except Exception:
                _log("Edge 应用窗口也不可用，改用默认浏览器：\n" + traceback.format_exc())
                _run_browser(url)
    finally:
        if httpd is not None:
            try:
                httpd.shutdown()
            except Exception:
                pass
    return 0


if __name__ == "__main__":
    try:
        code = main()
    except Exception:
        _log("程序异常：\n" + traceback.format_exc())
        code = 1
    sys.exit(code)
