# -*- coding: utf-8 -*-
"""首次运行初始化数据目录。

把 data_template/ 里的「空表结构」（只有表头、没有任何真实数据）复制成 data/，
让程序在全新电脑上第一次双击就能跑起来，之后再通过「导入表格 / 导入成绩 /
导入培养方案」把自己的数据填进去。

已有 data/ 且里面不是空的时，绝不覆盖、绝不清空。
"""
import os
import shutil

from app_paths import APP_HOME, res

DATA_DIR = os.path.join(APP_HOME, "data")
TEMPLATE_DIR = res("data_template")
OPTIONAL_DIRS = ("photos", "photos_src", "备份")


def has_data():
    """data 目录存在且至少有一个数据文件"""
    if not os.path.isdir(DATA_DIR):
        return False
    for fn in os.listdir(DATA_DIR):
        if fn.lower().endswith((".csv", ".json")):
            return True
    return False


def ensure_data(verbose=True):
    """确保 data/ 可用。返回 True 表示本次新建了数据目录。"""
    if has_data():
        return False
    if not os.path.isdir(TEMPLATE_DIR):
        os.makedirs(DATA_DIR, exist_ok=True)
        if verbose:
            print("! 未找到 data_template，已创建空 data 目录")
        return True
    os.makedirs(DATA_DIR, exist_ok=True)
    copied = 0
    for fn in sorted(os.listdir(TEMPLATE_DIR)):
        src = os.path.join(TEMPLATE_DIR, fn)
        dst = os.path.join(DATA_DIR, fn)
        if os.path.isfile(src) and not os.path.exists(dst):
            shutil.copy2(src, dst)
            copied += 1
    for d in OPTIONAL_DIRS:
        os.makedirs(os.path.join(APP_HOME, d), exist_ok=True)
    if verbose:
        print("已初始化数据目录：%s（%d 个空表模板）" % (DATA_DIR, copied))
        print("提示：在编辑模式里用「导入表格 / 导入成绩 / 导入培养方案」填入自己的数据。")
    return True


if __name__ == "__main__":
    ensure_data()
