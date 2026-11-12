# -*- coding: utf-8 -*-
"""打包发布版（绿色免安装 zip）。

用法：
    python make_release.py            # 生成 dist/辅导员工作台-v2.0.0-win64.zip
    python make_release.py 2.1.0      # 指定版本号

打出来的包里 **不含任何真实学生数据**：
- data/ 用的是 data_template 的空表结构（只有表头）
- 照片、备份、数据库、导出文件一律不带
真实数据请自己留在本机（或另行加密备份），不要上传 GitHub。
"""
import os
import shutil
import sys
import zipfile

from app_paths import APP_HOME as HERE

DIST = os.path.join(HERE, "dist")
DEFAULT_VERSION = "2.0.0"

# 需要随包发布的源码与文档
FILES = [
    "launcher.py", "server.py", "build.py", "academic.py", "xlsx_io.py",
    "db_sync.py", "extract.py", "app_paths.py", "init_data.py", "verify.py",
    "template.html",
    "使用说明.md", "字段说明.md", "数据库说明.md",
    "启动编辑模式.bat", "一键更新看板.bat", "同步数据库到看板.bat",
]
EXE_NAMES = ["辅导员工作台.exe"]
EMPTY_DIRS = ["photos", "photos_src", "备份"]

FIRST_RUN = """辅导员工作台 · 首次使用说明
========================================

一、直接打开看板（只读）
  双击「辅导员工作台.exe」→ 浏览器自动打开。
  关闭黑色命令行窗口 = 停止服务。全程只监听 127.0.0.1，不上传任何数据。

二、没有 exe / 提示缺少 Python 时
  双击「启动编辑模式.bat」（需要电脑已安装 Python 3.9+，安装时勾选 Add to PATH）。

三、想让看板显示自己的数据
  1) 编辑模式下点右上角「导入表格」，选学生主表等 xlsx / csv；
  2) 点「导入成绩」上传教务「全校成绩查询」xlsx；
  3) 点「导入培养方案」上传「20xx级xx主修培养方案--方案内表现课程信息.xlsx」。
  数据会写进 data/ 目录下的 CSV / JSON，照片放进 photos_src/。

四、目录说明
  data/          数据表（首次为空表结构，导入后填充）
  photos_src/    学生电子照片（文件名含学号即可自动匹配）
  photos/        缩略图缓存（自动生成，可删）
  备份/          建议定期把 data/ 复制到这里

五、注意
  本程序不联网、不上传数据。data/ 与 photos_src/ 含学生个人信息，
  请勿上传到 GitHub、网盘公开链接或任何公网位置。
"""


def copytree_files(root, dst):
    ok, miss = 0, []
    for fn in FILES:
        src = os.path.join(HERE, fn)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(dst, fn))
            ok += 1
        else:
            miss.append(fn)
    return ok, miss


def build(version):
    out_dir = os.path.join(DIST, "辅导员工作台-v%s" % version)
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir)
    os.makedirs(out_dir)

    ok, miss = copytree_files(HERE, out_dir)
    print("复制源码/文档 %d 个" % ok)
    if miss:
        print("! 缺少文件（已跳过）:", "、".join(miss))

    # 空数据目录（仅表头，不含任何真实数据）
    tpl = os.path.join(HERE, "data_template")
    if os.path.isdir(tpl):
        shutil.copytree(tpl, os.path.join(out_dir, "data"))
        print("data/ <- data_template（%d 个空表）" % len(os.listdir(tpl)))
    else:
        print("! 没有 data_template 目录，发布包将缺少 data/ 空表结构")

    for d in EMPTY_DIRS:
        os.makedirs(os.path.join(out_dir, d), exist_ok=True)

    for exe in EXE_NAMES:
        src = os.path.join(HERE, exe)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(out_dir, exe))
            print("已打包启动器:", exe, "%.1f MB" % (os.path.getsize(src) / 1048576))
        else:
            print("! 未找到 %s，发布包里将只能用 bat 启动" % exe)

    with open(os.path.join(out_dir, "首次使用说明.txt"), "w", encoding="utf-8-sig") as f:
        f.write(FIRST_RUN)
    readme = os.path.join(HERE, "README.md")
    if os.path.isfile(readme):
        shutil.copy2(readme, os.path.join(out_dir, "README.md"))

    zip_path = os.path.join(DIST, "辅导员工作台-v%s-win64.zip" % version)
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for root, _, files in os.walk(out_dir):
            for fn in files:
                p = os.path.join(root, fn)
                z.write(p, os.path.relpath(p, DIST))
    print("\n完成：%s（%.1f MB）" % (zip_path, os.path.getsize(zip_path) / 1048576))
    print("可直接上传到 GitHub Releases，作为下载附件。")
    return zip_path


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_VERSION)
