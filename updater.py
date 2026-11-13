# -*- coding: utf-8 -*-
"""检查 GitHub Release 更新，并支持一键下载安装。

更新源（仓库）解析优先级：
  1. 环境变量 WORKBENCH_REPO
  2. data/settings.json 的 repo 字段（页面上填写后写在这里）
  3. 程序目录 update.json 的 repo 字段（随程序一起发布）

仓库为空时返回 no_repo，页面会弹出输入框让用户填一次，之后记住。

私有仓库可额外配置 token（可选，只读权限即可）；公开仓库留空就行。
"""
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import urllib.error
import urllib.request

import version
from app_paths import APP_HOME, res

DATA = os.path.join(APP_HOME, "data")
SETTINGS = os.path.join(DATA, "settings.json")
SHIPPED = os.path.join(APP_HOME, "update.json")

API = "https://api.github.com/repos/%s/releases/latest"
UA = "CounselorWorkbench-Updater/1.0"
TIMEOUT = 8
CACHE_SEC = 1800          # 手动之外自动检查的缓存时间
_cache = {"at": 0.0, "data": None}

# 下载任务状态（/api/update/progress 轮询）
JOB = {"state": "idle", "pct": 0, "msg": "", "size": 0, "done": 0, "err": ""}
_job_lock = threading.Lock()


# ------------------------------------------------------------------ 配置
def _load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def _save_settings(obj):
    os.makedirs(DATA, exist_ok=True)
    with open(SETTINGS, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def get_repo():
    env = (os.environ.get("WORKBENCH_REPO") or "").strip()
    if env:
        return env
    for src in (_load_json(SETTINGS), _load_json(SHIPPED)):
        r = (src.get("repo") or "").strip()
        if r:
            return r
    return ""


def get_token():
    for src in (_load_json(SETTINGS), _load_json(SHIPPED)):
        t = (src.get("token") or "").strip()
        if t:
            return t
    return (os.environ.get("WORKBENCH_TOKEN") or "").strip()


def set_repo(repo, token=None):
    repo = (repo or "").strip()
    if repo and not re.match(r"^[\w.-]+/[\w.-]+$", repo):
        raise ValueError("仓库格式应为 用户名/仓库名，例如 zhangsan/counselor-workbench")
    cfg = _load_json(SETTINGS)
    cfg["repo"] = repo
    if token is not None:
        cfg["token"] = token.strip()
    _save_settings(cfg)
    _cache["at"] = 0.0
    return repo


# ------------------------------------------------------------------ 版本比较
def _ver_tuple(s):
    nums = re.findall(r"\d+", str(s or ""))
    return tuple(int(x) for x in nums[:4]) or (0,)


def is_newer(latest, current=None):
    a, b = _ver_tuple(latest), _ver_tuple(current or version.VERSION)
    return a > b


# ------------------------------------------------------------------ 检查更新
def _api_get(url, token=""):
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": "application/vnd.github+json",
    })
    if token:
        req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        return json.loads(r.read().decode("utf-8"))


def _pick_asset(assets):
    """优先挑安装版 exe，其次 zip。"""
    exe, zipf = None, None
    for a in assets or []:
        name = a.get("name") or ""
        low = name.lower()
        if low.endswith(".exe") and ("安装" in name or "setup" in low or "install" in low):
            if exe is None:
                exe = a
        elif low.endswith(".zip") and zipf is None:
            zipf = a
    return exe or zipf


def check(force=False):
    repo = get_repo()
    if not repo:
        return {"ok": False, "code": "no_repo", "current": version.VERSION,
                "error": "尚未设置更新源（GitHub 仓库）"}
    now = time.time()
    if (not force) and _cache["data"] and now - _cache["at"] < CACHE_SEC:
        return _cache["data"]
    try:
        rel = _api_get(API % repo, get_token())
    except urllib.error.HTTPError as e:
        if e.code == 404:
            err = "仓库不存在或为私有（私有仓库需在设置里填 token）：" + repo
        elif e.code == 403:
            err = "GitHub 接口限流，请稍后再试"
        else:
            err = "GitHub 返回 %s" % e.code
        return {"ok": False, "code": "http", "current": version.VERSION,
                "repo": repo, "error": err}
    except Exception as e:
        return {"ok": False, "code": "net", "current": version.VERSION,
                "repo": repo, "error": "网络不可用：%s" % str(e)[:120]}
    tag = rel.get("tag_name") or ""
    asset = _pick_asset(rel.get("assets"))
    out = {
        "ok": True,
        "current": version.VERSION,
        "latest": tag.lstrip("vV"),
        "tag": tag,
        "hasUpdate": is_newer(tag),
        "notes": (rel.get("body") or "").strip()[:1200],
        "page": rel.get("html_url") or "",
        "publishedAt": rel.get("published_at") or "",
        "repo": repo,
        "asset": None if not asset else {
            "name": asset.get("name"), "size": asset.get("size"),
            "url": asset.get("browser_download_url"),
        },
    }
    _cache.update(at=now, data=out)
    return out


# ------------------------------------------------------------------ 下载与安装
def _download(url, dest, token=""):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    if token:
        req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=30) as r, open(dest, "wb") as f:
        total = int(r.headers.get("Content-Length") or 0)
        JOB["size"] = total
        got = 0
        while True:
            chunk = r.read(65536)
            if not chunk:
                break
            f.write(chunk)
            got += len(chunk)
            JOB["done"] = got
            JOB["pct"] = int(got * 100 / total) if total else 0
    return got


def install_async(url, name="setup.exe"):
    """后台下载 → 启动安装程序 → 退出本程序。返回是否成功启动任务。"""
    with _job_lock:
        if JOB["state"] == "downloading":
            return False
        JOB.update(state="downloading", pct=0, msg="正在下载…", size=0, done=0, err="")

    def work():
        try:
            fn = re.sub(r"[^\w.\u4e00-\u9fa5-]", "_", name or "setup.exe")
            if not fn.lower().endswith(".exe"):
                raise ValueError("该版本没有提供安装版文件，请到 Release 页面手动下载")
            dest = os.path.join(tempfile.gettempdir(), "辅导员工作台-更新", fn)
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            JOB["msg"] = "正在下载…"
            size = _download(url, dest, get_token())
            if size < 1024 * 1024:
                raise ValueError("下载文件异常（%d 字节）" % size)
            JOB.update(state="launching", pct=100, msg="下载完成，正在启动安装程序…")
            # 让出锁：本进程稍后退出，安装程序接管文件替换。
            # 用 DETACHED_PROCESS 启动，安装包不挂在主程序的进程组里
            # （安装脚本会 taskkill 掉「辅导员工作台.exe」，挂在一起容易被连带结束）。
            # 装到哪个目录交给安装脚本自己决定（固定为 %LOCALAPPDATA%\Programs\辅导员工作台，
            # 且 UsePreviousAppDir=no）；这里不传 /DIR，避免把“当前运行的目录”当成安装目标——
            # 程序万一跑在临时目录里，就会一直更新到临时目录去，版本永远不变。
            # /LOG 留一份安装日志，下次再出问题可以直接看。
            logf = os.path.join(os.path.dirname(dest), "install.log")
            args = [dest, "/SILENT", "/CLOSEAPPLICATIONS", "/NORESTART", "/LOG=" + logf]
            flags = 0
            if os.name == "nt":
                flags = 0x00000008 | 0x00000200   # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
            subprocess.Popen(args, close_fds=True, creationflags=flags)
            JOB["state"] = "done"
            threading.Timer(2.5, lambda: os._exit(0)).start()
        except Exception as e:
            JOB.update(state="error", err=str(e)[:200], msg="更新失败")

    threading.Thread(target=work, daemon=True).start()
    return True


def progress():
    return dict(JOB)
