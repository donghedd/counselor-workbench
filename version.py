# -*- coding: utf-8 -*-
"""版本号唯一来源。

改版本只改这一个文件：
  - make_release.py 用它生成 dist 目录名与 zip 名
  - installer/installer.iss 由 编译安装包.bat 通过 /DAppVersion 传入
  - build.py 写进 dashboard 的 DATA.version，页面侧栏显示
  - updater.py 用它和 GitHub 上最新 release 的 tag 比对
"""

VERSION = "2.1.16"
