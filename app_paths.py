# -*- coding: utf-8 -*-
"""统一路径解析：兼容"源码运行"与"PyInstaller 打包成 exe 运行"。

- APP_HOME：数据与产物的根目录。exe 运行时 = exe 所在目录；源码运行时 = 本文件目录。
            这个目录必须可写（要写 CSV / dashboard.html / db）。
- RES_DIR ：只读资源目录。exe 运行时为打包解压的临时目录（sys._MEIPASS）；源码运行时同 APP_HOME。

环境变量 WORKBENCH_HOME 可显式覆盖 APP_HOME（便于测试或多副本共存）。
"""
import os
import sys


def _app_home():
    env = os.environ.get("WORKBENCH_HOME")
    if env and os.path.isdir(env):
        return os.path.abspath(env)
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_HOME = _app_home()
RES_DIR = getattr(sys, "_MEIPASS", None) or APP_HOME

# 打包后脚本被收进 exe，import 子模块需要 APP_HOME 在 sys.path 里
if APP_HOME not in sys.path:
    sys.path.insert(0, APP_HOME)


def p(*parts):
    """数据/产物路径（可写目录）"""
    return os.path.join(APP_HOME, *parts)


def res(name, *fallbacks):
    """资源路径：优先 APP_HOME（允许用户就近覆盖），其次打包内部。

    例如 template.html 用户想自定义版式时，把它放到 exe 旁边即可生效。
    """
    for n in (name,) + fallbacks:
        f = os.path.join(APP_HOME, n)
        if os.path.exists(f):
            return f
    for n in (name,) + fallbacks:
        f = os.path.join(RES_DIR, n)
        if os.path.exists(f):
            return f
    return os.path.join(APP_HOME, name)
