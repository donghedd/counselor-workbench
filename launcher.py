# -*- coding: utf-8 -*-
"""辅导员工作台 · 双击启动入口（由 PyInstaller 打包为 辅导员工作台.exe）

双击后：
1. 自动定位 exe 所在目录（数据目录），不需要用户配置任何路径；
2. 若 dashboard.html 缺失，先自动生成一次；
3. 启动仅监听 127.0.0.1 的本地服务，并用默认浏览器打开工作台；
4. 关闭黑窗口 = 停止服务，不联网、不上传任何数据。
"""
import os
import sys
import traceback

from app_paths import APP_HOME

TITLE = "辅导员工作台 - 编辑模式（关闭本窗口即停止服务）"


def _setup_console():
    if os.name != "nt":
        return
    try:
        os.system("title " + TITLE)
    except Exception:
        pass
    # 避免控制台中文乱码
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _pause(msg="\n按回车键退出..."):
    try:
        input(msg)
    except Exception:
        pass


def main():
    _setup_console()
    try:
        os.chdir(APP_HOME)
    except Exception:
        pass

    data_dir = os.path.join(APP_HOME, "data")
    if not os.path.isdir(data_dir):
        # 全新机器首次运行：用 data_template 的空表结构自动初始化，不让用户先手动建目录
        try:
            import init_data
            init_data.ensure_data()
        except Exception:
            traceback.print_exc()
    if not os.path.isdir(data_dir):
        print("[错误] 没有找到数据目录：")
        print("       %s" % data_dir)
        print("")
        print("请把本程序放在完整的工作台文件夹内运行，")
        print("该文件夹里应当有：data（或 data_template）、template.html 等文件。")
        return 1

    if not os.path.exists(os.path.join(APP_HOME, "dashboard.html")):
        print("首次运行，正在生成看板，请稍候...")
        try:
            import build
            build.main()
        except Exception:
            print("[错误] 看板生成失败：")
            traceback.print_exc()
            return 1

    import server
    server.main()
    return 0


if __name__ == "__main__":
    try:
        code = main()
    except Exception:
        print("\n[错误] 程序运行出现异常：")
        traceback.print_exc()
        code = 1
    if code:
        _pause()
