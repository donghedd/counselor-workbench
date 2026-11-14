# -*- coding: utf-8 -*-
"""学业模块：培养方案解析 + 全校成绩导入分析。

数据文件（均在 data/ 下）：
- 12_培养方案.json   各专业培养方案（类别树 + 要求学分 + 方案内课程清单 + 课程→课组归属 cgroups）
- 13_成绩分析.json   最近一次全校成绩导入的逐生分析（挂科/应修未修/学科交叉/学分/警示）
- 14_学业警示.csv    人工登记的学业警示（自动分析不落这张表，前端合并展示）

培养方案 xlsx 是教务系统导出的特殊版式（合并单元格交错），无法按普通表格读。
本模块直接解析 xlsx 的 sharedStrings 与单元格流：
- 类别树来自「类别#g（要求学分：X 已修学分：Y）g#/子类」编码串与面包屑碎片；
- 方案内课程集合来自「课程名[课程号]」形 token；
- 课程→课组归属（cgroups）来自编码串分块的先后顺序：编码串之后紧随的课程
  就属于该课组；同一行「编码串+课程名」是学校认定的可替代关系，一并收录。

成绩分析口径（大四上=当前学期进行中）：
- 挂科：已完成学期中 成绩<60 / 文本"不及格" / 是否及格=否且学分>0（六级等 >=300 大分不算）；
- 成绩替代 / 成绩认定（方案外）：这类行的「替代课程号」才是培养方案内的课号（占 152/156 条），
  视同该课已修读——计入 taken（不再误报未选课）、学分归口到方案内课号；
  同一门课（含替代课号）任一次学期通过即视为已获学分，之后重修再挂不算未获学分；
- 已修学分：已完成/在修学期中通过课程学分之和（同课程号取最大学分一次）；
- 应修未修：按培养方案课组口径 —— 先把成绩行按 cgroups 归口到具体课组，
  课组学分没修满的才需要再修；组内没修的课按计划学期顺序列出，
  学分凑够缺口即止（组里已经修够的，剩下的课不属于"必须修"，不再点名）；
  计划学期仍按同专业同届修读学期众数推断（只用于判断这学期是否该修）；
  文化素质类（通识限选>文化素质类-计科）与学科交叉课组属"自由选课凑学分"课组，
  只记学分缺口、不点名具体课程（不必修某一门，达标即可）；
- 学科交叉：培养方案外的课程（课号不在本专业方案、课名也对不上方案课程），
  只要不是文化素质类（含校公选课类别）、体育专项/俱乐部、四史、外语、必修与专业类课程，
  学分一律计入"学科交叉课组"——不限于方案里列出的那几门交叉课；
- 挂科：文化素质类 / 学科交叉课程未通过标记 flex（不强制重修，可在课组内另选补足学分），
  不计入警示挂科门数；其余必修/限选课程挂科才计入警示；
- 学分警示：按修业年限阈值表（1:30/25/15 2:70/55/45 3:110/85/75 4:留级115 退学105）；
- 警示类型三类（从重到轻）：退学警示/留级警示看累计学分有没有达到该年级学分线；
  学业警示按培养方案看应修未修、以及挂了方案内且课组学分仍缺口的课（课组达标则不提醒）。
"""
import csv, datetime, io, json, os, re, zipfile, collections
import xml.etree.ElementTree as ET

from app_paths import APP_HOME as HERE

DATA = os.path.join(HERE, "data")
PLAN_FILE = os.path.join(DATA, "12_培养方案.json")
ANALYSIS_FILE = os.path.join(DATA, "13_成绩分析.json")
WARN_FILE = os.path.join(DATA, "14_学业警示.csv")
WARN_HEADER = ["记录ID", "学号", "姓名", "警示类型", "级别", "原因", "学期", "来源", "更新时间"]
FOLLOW_FILE = os.path.join(DATA, "15_成绩跟进.csv")
FOLLOW_HEADER = ["学号", "课程号", "课程名", "跟进方式", "新成绩", "是否通过", "备注", "更新时间"]

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RE_CREDIT = re.compile(r"（要求学分：([\d.]+) 已修学分：([\d.]+)）")
RE_COURSE = re.compile(r"^[^#\[\]]{2,60}（?[^（）]*）?\[(\d{6,})\]$")
RE_TERM = re.compile(r"^(\d{4})-(\d{4})学年\s*([春秋])季学期$")
RE_NAME = re.compile(r"^[\u4e00-\u9fa5A-Za-z0-9（）()\-·ⅠⅡⅢⅣ／/ ]{2,30}$")
MULTI_MARK = re.compile(r"^\((?:多选|单选)\d\)")

YEAR_THRESHOLDS = {1: (30, 25, 15), 2: (70, 55, 45), 3: (110, 85, 75), 4: (None, 115, 105), 5: (None, None, 135)}
GRADED_TOPS = ("通识必修", "学科必修", "专业必修", "通识限选", "专业限选", "学科交叉课组")
# 自由选课凑学分课组（挂科不强制重修、应修未修不点名具体课程、只查课组总学分）
FLEX_GROUPS = ("文化素质类-计科", "外语基础课组I类", "外语拓展课组I类", "学科交叉课组", "艺术类")
# 外语类课程（含各语种），用于「外语拓展课组」的学分合池
RE_EN = re.compile(r"英语|外语|日语|法语|西班牙语|德语|俄语")
FALLBACK_ROOTS = ("通识必修", "学科必修", "专业必修", "通识限选", "专业限选", "学科交叉课组",
                  "第二课堂-专业品牌活动", "第二课堂-通识类")

SCAFFOLD = {"课程名[课程号]", "是否通过", "成绩", "学分", "替代课程名[课程号]", "获取方式",
            "选课学年学期", "成绩学年学期", "计划学年学期", "课程性质", "重修重考",
            "header", "content", "", " "}
NATURE_SET = {"必修", "限选", "任选", "选修"}
STOPWORDS = {"优秀", "良好", "中等", "及格", "不及格", "已选课", "未修读", "已修", "未修",
             "通过", "不通过", "替代", "课组", "培养方案", "初修", "重修", "补考", "重考",
             "未通过", "已通过"} | NATURE_SET


def _strip_multi(name):
    return MULTI_MARK.sub("", (name or "").strip()).strip()


def _clean_markers(v):
    v = re.sub(r"#([a-z])", "\x01", v)
    v = re.sub(r"([a-z])#", "\x01", v)
    return v


def parse_path(v):
    """编码串 -> [(类别名, 要求学分), ...]。要求每段都有名称，防面包屑碎片误入。"""
    v = _clean_markers(v)
    segs = []
    for m in re.finditer(r"/?([^\x01/（]+?)\x01?（要求学分：([\d.]+) 已修学分：([\d.]+)）", v):
        name = _strip_multi(m.group(1))
        if name:
            segs.append((name, float(m.group(2))))
    return segs


def _term_key(t):
    m = RE_TERM.match((t or "").strip())
    if not m:
        return None
    return (int(m.group(1)), 0 if m.group(3) == "秋" else 1)


def _plan_term_of(t):
    """方案自带的「计划学年学期」（可能是逗号分隔的多学期串）-> 学期键。"""
    for part in (t or "").split(","):
        tk = _term_key(part)
        if tk is not None:
            return tk
    return None


def term_label(tk):
    return "%d-%d学年 %s季学期" % (tk[0], tk[0] + 1, "秋" if tk[1] == 0 else "春")


def _read_strings_and_cells(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    ss_xml = z.read("xl/sharedStrings.xml").decode("utf-8")
    strings = re.findall(r"<t[^>]*>([^<]*)</t>", ss_xml)
    root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    cells = []
    for row in root.iter(M + "row"):
        for c in row:
            v = c.find(M + "v")
            if v is None:
                continue
            cells.append(strings[int(v.text)] if c.get("t") == "s" else v.text)
    return strings, cells


# ---------------------------------------------------------------- 培养方案解析
GROUP_OPEN = re.compile(r"#[a-z]（要求学分")
RE_TERM_ANY = re.compile(r"^\d{4}-\d{4}学年\s*[春秋]季学期")
RE_NUM = re.compile(r"^\d+(?:\.\d+)?$")


def _col_idx(letter):
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n


def _plan_scan(data):
    """解析方案表：课组层级 + 课程归口 + 学分/计划学期/课程性质。

    教务导出的这份 xlsx 有两个特点，必须分开处理：

    一、**组头编码串（`(多选4)通识限选#g（要求学分：17.0）g#/…`）是「孤儿字符串」**
        —— 它们留在 sharedStrings 里，但已没有任何单元格引用（导出工具隐藏了这些行）。
        所以课组层级和「课程归哪个组」只能按 sharedStrings 的出现顺序还原：
        一串组头后面跟着的课程就属于该组，组头用 / 分隔的每一段入树（父 > 子 > 孙）。

    二、**课程与字段是「多块粘贴拼图」**：每个课组块占一组列，块与块在行方向错位，
        同一个 Excel 行里常常并排着两三个不同块，列位置没有全局含义。但同一列内
        的顺序是稳定的：

            A78  中国文化概论(英文)[2314602020]   ← 课程
            A79  2025-2026学年 秋季学期,…         ← 该课的选课/计划学期
            A80  中国经典英译赏析[2323001020]      ← 下一门课
            …
            A88  4.0                               ← 上一门课的学分

        所以字段按「列流」取：列内顺序配对，本列缺的再用同一个 Excel 行兜底。

    返回 (tree, cg, courses)：
      tree     {层级元组: 要求学分}
      cg       {课号: [课组 key]}
      courses  {课号: {n 课程名, c 学分, t 计划学年学期, k 课程性质, cat 归口课组}}
    """
    z = zipfile.ZipFile(io.BytesIO(data))
    names = z.namelist()
    sheet = next((n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)), None)
    root = ET.fromstring(z.read(sheet))
    strings = []
    if "xl/sharedStrings.xml" in names:
        sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in sroot.iter(M + "si"):
            strings.append("".join((t.text or "") for t in si.iter(M + "t")))

    tree, cg, courses = {}, {}, {}
    cur = None      # 当前课组 key（组头里的最深层）
    crumbs = []     # 面包屑栈 [(名, 要求学分)]
    has_chain = any(GROUP_OPEN.search(x) for x in strings)

    def put(path):
        for i in range(1, len(path) + 1):
            k = tuple(n for n, _ in path[:i])
            r = path[i - 1][1]
            if r is None:
                tree.setdefault(k, 0.0)
            else:
                tree[k] = max(tree.get(k, 0.0), r)

    # ---------- 第一段：sharedStrings 顺序 → 课组树 + 课程归口 ----------
    for raw in strings:
        t = raw.strip()
        if not t:
            continue
        if GROUP_OPEN.search(raw):
            seg = parse_path(raw)
            if seg:
                crumbs = [[n, r] for n, r in seg]
                put(seg)
                cur = " > ".join(n for n, _ in seg)
            continue
        # 老版导出格式（无编码串）才用面包屑碎片建层级
        if not has_chain and t.startswith("/"):
            nm = _strip_multi(t.lstrip("/").strip())
            if nm:
                idx = None
                for i in range(len(crumbs) - 1, -1, -1):
                    if crumbs[i][0] == nm:
                        idx = i
                        break
                if idx is None:
                    crumbs.append([nm, None])
                else:
                    crumbs = crumbs[:idx + 1]
                    crumbs[idx][1] = None
            continue
        if not has_chain and t.startswith("（要求学分"):
            m = RE_CREDIT.search(t)
            if m:
                if not crumbs:
                    crumbs = [["", None]]
                crumbs[-1][1] = float(m.group(1))
                put([(n, r) for n, r in crumbs if n])
            continue
        if t in SCAFFOLD:
            continue
        m = RE_COURSE.match(t)
        if m:
            code = m.group(1)
            nm = re.sub(r"\[\d{6,}\]$", "", t).strip()
            d = courses.setdefault(code, {"n": nm, "c": 0.0, "t": "", "k": "", "cat": ""})
            if not d["n"]:
                d["n"] = nm
            if cur:
                lst = cg.setdefault(code, [])
                if cur not in lst:
                    lst.append(cur)
                if not d["cat"]:
                    d["cat"] = cur

    # ---------- 第二段：单元格网格 → 学分 / 计划学期 / 课程性质 ----------
    def cval(c):
        if c.get("t") == "s":
            v = c.find(M + "v")
            try:
                return strings[int(v.text)] if v is not None else ""
            except Exception:
                return ""
        if c.get("t") == "inlineStr":
            return "".join((t.text or "") for t in c.iter(M + "t"))
        v = c.find(M + "v")
        return (v.text or "") if v is not None else ""

    cells = []
    for c in root.iter(M + "c"):
        ref = c.get("r") or ""
        m = re.match(r"([A-Z]+)(\d+)$", ref)
        if not m:
            continue
        t = (cval(c) or "").strip()
        if not t:
            continue
        cells.append((int(m.group(2)), _col_idx(m.group(1)), t))
    cells.sort(key=lambda x: (x[0], x[1]))

    # 「替代课程名[课程号]」那一列装的是被替代的课，不是方案课程本身
    alt_cols = {ci for (_r, ci, t) in cells if t == "替代课程名[课程号]"}
    byrow = {}
    for cell in cells:
        byrow.setdefault(cell[0], []).append(cell)
    wins = {}

    def row_fields(ci, r):
        """同一个 Excel 行（以及上下各一行）里其它单元格的字段。

        多块粘贴会把一张卡片的「值」挤到相邻行/相邻列，所以性质在 ±1 行内找；
        学分和学期仍只用同一行——跨行的数字极易是隔壁卡片的成绩。
        """
        acc = {"rnums": [], "rterms": [], "rnats": []}
        for rr in (r, r - 1, r + 1):
            for (_r2, ci2, t2) in byrow.get(rr, []):
                if _r2 == r and ci2 == ci:
                    continue
                if GROUP_OPEN.search(t2) or RE_COURSE.match(t2):
                    continue
                if t2 in SCAFFOLD or t2.startswith("（要求学分") or t2.startswith("/"):
                    continue
                if t2 in NATURE_SET:          # 先认性质：STOPWORDS 里也含这些词
                    acc["rnats"].append(t2)
                    continue
                if t2 in STOPWORDS:
                    continue
                if _r2 != r:
                    continue
                elif RE_TERM_ANY.match(t2):
                    acc["rterms"].append(t2)
                elif RE_NUM.match(t2) and 0 < float(t2) <= 12:
                    acc["rnums"].append(float(t2))
        return acc

    def settle(ci):
        w = wins.pop(ci, None)
        if not w:
            return
        d = courses.get(w["code"])
        if d is None:
            return
        if not d["c"]:
            nums = w["nums"] or w["rnums"]
            if nums:
                d["c"] = round(min(nums), 2)
        if not d["t"]:
            # 同一行常有「选课学年学期 / 成绩学年学期 / 计划学年学期」三列，
            # 计划列在最后；方案要的是计划学期，所以优先取行内最后一个。
            terms = (w["rterms"][-1:] or w["terms"][:1])
            if terms:
                d["t"] = terms[0].split(",")[0].strip()
        if not d["k"]:
            nats = w["nats"] or w["rnats"]
            if nats:
                d["k"] = nats[0]

    for (r, ci, t) in cells:
        if GROUP_OPEN.search(t):
            settle(ci)
            continue
        if t.startswith("（要求学分") or t.startswith("/"):
            settle(ci)
            continue
        if t in SCAFFOLD:
            continue
        m = RE_COURSE.match(t)
        if m:
            settle(ci)
            if ci in alt_cols:
                continue
            wins[ci] = dict(code=m.group(1), nums=[], terms=[], nats=[], **row_fields(ci, r))
            continue
        w = wins.get(ci)
        if w is None:
            continue
        if RE_TERM_ANY.match(t):
            w["terms"].append(t)
        elif t in NATURE_SET:
            w["nats"].append(t)
        elif RE_NUM.match(t) and 0 < float(t) <= 12:
            w["nums"].append(float(t))
    for ci in list(wins):
        settle(ci)
    return tree, cg, courses


def _plan_cgroups(data):
    """兼容旧接口：只要「课程 -> 课组」。"""
    return _plan_scan(data)[1]


def parse_plan(fname, data):
    tree, cg, courses = _plan_scan(data)
    cats = [{"key": " > ".join(k), "path": list(k), "req": tree[k]}
            for k in sorted(tree, key=lambda p: (len(p), p))]
    codes = {c: d["n"] for c, d in courses.items()}
    year, major = "", ""
    m = re.search(r"(\d{4})级\s*([\u4e00-\u9fa5A-Za-z]+?)主修培养方案", fname)
    if m:
        year, major = m.group(1), m.group(2)
    else:
        m = re.search(r"([\u4e00-\u9fa5A-Za-z]+?)主修培养方案", fname)
        if m:
            major = m.group(1)
    if not major:
        raise ValueError("无法从文件名识别专业（文件名需含“xxx主修培养方案”）")
    return {"major": major, "year": year, "cats": cats, "codes": codes,
            "courses": courses, "cgroups": cg, "source": os.path.basename(fname)}

def _migrate_plan(p):
    """旧版导入的培养方案没有 courses（课程明细）字段：从 codes + cgroups 重建，
    学分/计划学期/课程性质从成绩分析的 planCourses 补，避免升级后要重新导入。

    归口取「cgroups 里存在且属于本方案 cats 的最长路径」——旧 cgroups 与成绩归口
    都可能有噪声（替代关系误判等），cat 必须能对上方案自己的课组树才可信。
    """
    if p.get("courses") or not p.get("codes"):
        return p
    keys = {c["key"] for c in p.get("cats") or []}
    cg = p.get("cgroups") or {}
    pc = {}
    try:
        with open(ANALYSIS_FILE, encoding="utf-8") as f:
            pc = (json.load(f).get("planCourses") or {}).get(p.get("major") or "", {})
    except Exception:
        pass
    courses = {}
    for code, nm in p["codes"].items():
        ref = pc.get(code) or {}
        cand = [g for g in cg.get(code) or [] if g in keys]
        cand.sort(key=len, reverse=True)
        cat = cand[0] if cand else (ref.get("cat") if ref.get("cat") in keys else "")
        courses[code] = {
            "n": ref.get("n") or nm,
            "c": ref.get("c") or 0.0,
            "t": ref.get("t") or "",
            "k": ref.get("k") or "",
            "cat": cat,
        }
    p["courses"] = courses
    p["_migrated"] = True
    return p


def load_plans():
    if not os.path.exists(PLAN_FILE):
        return {"majors": {}}
    with open(PLAN_FILE, encoding="utf-8") as f:
        d = json.load(f)
    changed = False
    for mj, p in (d.get("majors") or {}).items():
        if not p.get("courses"):
            _migrate_plan(p)
            if p.get("_migrated"):
                p.pop("_migrated", None)
                changed = True
    if changed:
        os.makedirs(DATA, exist_ok=True)
        with open(PLAN_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
    return d


def save_plan(fname, data):
    plan = parse_plan(fname, data)
    plans = load_plans()
    old = plans["majors"].get(plan["major"]) or {}
    # 不允许覆盖已导入的方案：同一专业（本工具按专业存一份，不分年级）已存在时必须先删除，
    # 避免「上传错文件/解析失败的文件」把好方案悄悄换成空方案。
    if plans["majors"].get(plan["major"]):
        raise ValueError(
            "「%s」已导入过培养方案（%s级，%s 导入），不能重复导入；"
            "请先在「培养方案」卡片上点「删除方案」，确认删除后再导入新文件。"
            % (plan["major"], old.get("year") or "?", old.get("importedAt") or "时间未知"))
    # 解析结果兜底：一份正常的方案一定同时有课组和方案内课程。
    # 只认 sharedStrings 的解析器遇到 inlineStr 存文本的 xlsx 会得到空结果，
    # 这里直接拒绝，避免写进一份「零要求」的空方案。
    if not plan["cats"] or not plan["codes"]:
        raise ValueError(
            "这份文件没有解析出任何课组/课程（识别到专业：%s），已停止导入，原有方案未改动；"
            "常见原因：文件以 inlineStr 方式保存，或选的不是「xx主修培养方案」文件。"
            % (plan["major"] or "未知"))
    plan["importedAt"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    plans["majors"][plan["major"]] = plan
    os.makedirs(DATA, exist_ok=True)
    with open(PLAN_FILE, "w", encoding="utf-8") as f:
        json.dump(plans, f, ensure_ascii=False)
    return {"major": plan["major"], "year": plan["year"], "cats": len(plan["cats"]),
            "courses": len(plan["codes"]), "file": plan["source"]}


def set_plan_req(major, key, req):
    plans = load_plans()
    plan = plans["majors"].get(major)
    if not plan:
        raise ValueError("尚未导入该专业的培养方案：%s" % major)
    hit = None
    for c in plan["cats"]:
        if c["key"] == key:
            hit = c
            break
    if hit is None:
        raise ValueError("未找到类别：" + key)
    hit["req"] = float(req)
    hit["reqOverride"] = True
    with open(PLAN_FILE, "w", encoding="utf-8") as f:
        json.dump(plans, f, ensure_ascii=False)
    return True


def del_plan(major):
    plans = load_plans()
    if major in plans.get("majors", {}):
        del plans["majors"][major]
        with open(PLAN_FILE, "w", encoding="utf-8") as f:
            json.dump(plans, f, ensure_ascii=False)
        return True
    return False


# ---------------------------------------------------------------- 成绩分析
def _read_students():
    p = os.path.join(DATA, "01_学生主表.csv")
    if not os.path.exists(p):
        return []
    with open(p, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def _f(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return 0.0


def _norm_name(nm):
    """课程名归一化：只保留汉字/字母/数字，去掉括号与 Ⅰ/Ⅱ/(双语) 等标记。
    用于匹配"重修后课程号变了"的同一门课（如 计算机网络基础(Ⅰ) → 计算机网络基础）。"""
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fa5]", "", (nm or "").strip())


def _row_codes(r):
    """一行成绩对应的课号：实际修读课号 + 「替代课程号」。
    教务导出里「成绩替代」「成绩认定（方案外）」的行，替代课程号往往才是培养方案内的课号
    （152/156 条如此），两者是同一门课的两种编号，必须都算作已修，
    否则会误报“未选课”，学分也会缺口。"""
    out = []
    for k in ("课程号", "替代课程号"):
        c = (r.get(k) or "").strip()
        if c and c not in out:
            out.append(c)
    return out


def _fail(row):
    s = (row.get("总成绩") or "").strip()
    credit = _f(row.get("学分"))
    if s == "不及格":
        return True
    if re.match(r"^[\d.]+$", s):
        val = float(s)
        if val >= 300:
            return False
        if val < 60:
            return True
    if (row.get("是否及格") or "").strip() == "否" and credit > 0:
        return True
    return False


SUB_MAP = {"哲学社会科学": ["哲学社会科学"],
           "艺术体育-艺术类": ["艺术体育", "艺术类"],
           "艺术体育-体育类": ["艺术体育", "体育类"],
           "自然科学与工程科技": ["自然科学与工程技术"],
           "文化交流与国际视野": ["文化交流与国际视野"]}
SIXTY_HISTORY = re.compile(r"党史|新中国史|改革开放史|社会主义发展史|四史")


def _cat_of_row(r, plan_codes=None, plan_names=None):
    """课程归口。传 plan_codes 时，对「培养方案外的课程」改判：

    课号不在本专业培养方案、课名也对不上方案里的课程，一律算「学科交叉课组」
    （不再只认方案里列的那几门交叉课）。修读性质标着「专业选修/专业限选/学科限选」
    但方案里根本没有这门课的，同样算交叉学科——培养方案是唯一判据。
    但以下仍按原体系归口：
    文化素质类（带（文化素质）后缀或校公选课类别）、体育专项/俱乐部、四史类、
    外语类课程、以及通识必修/学科必修/专业必修这类硬必修课。
    课名能对上方案课程的（重修换了课号）不算方案外，仍按方案内课程归口。
    """
    nm = (r.get("课程名") or "").strip()
    cc = (r.get("课程类别") or "").strip()
    gx = (r.get("校公选课类别") or "").strip()
    if plan_codes is not None and not any(c in plan_codes for c in _row_codes(r)):
        if not (plan_names and _norm_name(nm) in plan_names):
            if "（学科交叉）" in nm:
                return ["学科交叉课组"]
            if not ("（文化素质）" in nm or gx
                    or re.match(r"^体育专项|^体育俱乐部", nm)
                    or SIXTY_HISTORY.search(nm) or RE_EN.search(nm)
                    or cc in ("通识必修", "学科必修", "专业必修")):
                return ["学科交叉课组"]
    if "（学科交叉）" in nm:
        return ["学科交叉课组"]
    if "（文化素质）" in nm or gx:
        return ["通识限选", "文化素质类-计科"] + SUB_MAP.get(gx, [])
    if re.match(r"^体育专项", nm):
        return ["通识限选", "体育类-限选", "体育专项"]
    if re.match(r"^体育俱乐部", nm):
        return ["通识限选", "体育类-限选", "体育俱乐部"]
    if SIXTY_HISTORY.search(nm):  # 含「中共党史」「改革开放史」等，修 1 门即可
        return ["通识限选", "四史类"]
    if cc == "通识必修":
        if re.search(r"形势与政策|思想道德|中国近现代史纲要|马克思主义|毛泽东思想|习近平|国家安全", nm):
            return ["通识必修", "思政类课组"]
        if re.match(r"^大学英语|^综合英语|^基础英语|^通用英语|英语|外语", nm):
            return ["通识必修", "外语基础课组I类"]
        if re.search(r"军事理论|军事技能|军训", nm):
            return ["通识必修", "国防教育类"]
        if re.search(r"入学教育|新生研讨|新生导学", nm):
            return ["通识必修", "新生导学类"]
        if re.search(r"心理健康|健康教育", nm):
            return ["通识必修", "健康教育类"]
        if re.search(r"体能训练|体能达标", nm):
            return ["通识必修", "体育类-体能训练"]
        if re.search(r"劳动|社会实践|社会调查", nm):
            return ["通识必修", "实践类"]
        if re.search(r"职业生涯|生涯规划|职业发展|就业指导|就业准备|职业能力实践", nm):
            return ["通识必修", "发展规划类"]
        if re.search(r"素质拓展|综合拓展|创新创业|技术经济|IT项目管理|HSE管理|计算机伦理|信息检索|文献检索", nm):
            return ["通识必修", "综合拓展类"]
        return ["通识必修"]
    if cc == "学科必修":
        return ["学科必修"]
    if cc == "专业必修":
        return ["专业必修"]
    if cc in ("专业限选", "学科限选", "专业选修"):
        return ["专业限选"]
    if cc == "通识限选":
        return ["通识限选", "外语拓展课组I类"]
    if cc == "通识选修":
        return ["通识限选", "文化素质类-计科"]
    return ["学科交叉课组"]


# ---------------------------------------------------------------- 学业警示表（人工）
def read_warnings():
    if not os.path.exists(WARN_FILE):
        return []
    with open(WARN_FILE, encoding="utf-8-sig", newline="") as f:
        return [r for r in csv.DictReader(f)]


def _write_warnings(rows):
    os.makedirs(DATA, exist_ok=True)
    with open(WARN_FILE, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=WARN_HEADER)
        w.writeheader()
        w.writerows(rows)


def add_warning(o):
    action = o.get("action", "add")
    rid = str(o.get("id") or "").strip()
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    if action == "delete":
        rows = [r for r in read_warnings() if r["记录ID"] != rid]
        _write_warnings(rows)
        return {"deleted": rid}
    rec = {"学号": str(o.get("sid") or "").strip(),
           "姓名": str(o.get("name") or "").strip(),
           "警示类型": str(o.get("type") or "学业警示").strip(),
           "级别": str(o.get("level") or "").strip(),
           "原因": str(o.get("reason") or "").strip(),
           "学期": str(o.get("term") or "").strip(),
           "来源": "手动",
           "更新时间": now}
    if not rec["学号"]:
        raise ValueError("缺少学号")
    rows = read_warnings()
    if action == "update":
        hit = None
        for r in rows:
            if r["记录ID"] == rid:
                hit = r
                break
        if hit is None:
            raise ValueError("未找到记录：" + rid)
        hit.update(rec)
    else:
        nums = [int(r["记录ID"]) for r in rows if r["记录ID"].strip().isdigit()]
        rec["记录ID"] = str(max(nums) + 1 if nums else 1)
        rows.append(rec)
    _write_warnings(rows)
    return rec


# ---------------------------------------------------------------- 成绩跟进（人工）
def read_follow():
    if not os.path.exists(FOLLOW_FILE):
        return []
    with open(FOLLOW_FILE, encoding="utf-8-sig", newline="") as f:
        return [r for r in csv.DictReader(f) if (r.get("学号") or "").strip()]


def _write_follow(rows):
    os.makedirs(DATA, exist_ok=True)
    with open(FOLLOW_FILE, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FOLLOW_HEADER)
        w.writeheader()
        w.writerows(rows)


def save_follow(o):
    """登记/撤销「某学生某门课」的跟进结果（补考/重修/免修等）。

    只记录人工跟进结论，不修改教务导出的原始成绩文件；
    分析时把「已通过」的那一条落到成绩行上，学分与挂科随之重算，随时可撤销。
    """
    action = (o.get("action") or "save").strip()
    sid = str(o.get("sid") or "").strip()
    code = str(o.get("code") or "").strip()
    if not sid or not code:
        raise ValueError("缺少学号或课程号")
    rows = [r for r in read_follow()
            if not (r.get("学号") == sid and r.get("课程号") == code)]
    if action == "delete":
        _write_follow(rows)
        return {"deleted": True, "sid": sid, "code": code}
    passed = str(o.get("pass") or "是").strip() in ("是", "1", "true", "True", "通过")
    score = str(o.get("score") or "").strip()
    num = _f(score) if score else 0.0
    if passed and num < 60:
        num = 60.0          # 通过即视为及格，成绩至少 60
    if score or num:
        score = "%g" % num
    rec = {"学号": sid, "课程号": code,
           "课程名": str(o.get("name") or "").strip(),
           "跟进方式": str(o.get("way") or ("补考" if passed else "待处理")).strip(),
           "新成绩": score,
           "是否通过": "是" if passed else "否",
           "备注": str(o.get("note") or "").strip()[:120],
           "更新时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}
    rows.append(rec)
    _write_follow(rows)
    return rec


def follow_index():
    """{学号: [跟进记录...]}，供页面展示。"""
    out = collections.defaultdict(list)
    for r in read_follow():
        out[r.get("学号")].append(r)
    return out


def _apply_follow(rows):
    """把「已通过」的跟进结论落到成绩行上，返回命中条数。

    同一门课若有多条重修记录，只改最近一学期那一行（不追溯历史），
    这样后续的挂科判定、学分累计、未获科目都会自然重算。
    """
    idx = {}
    for r in read_follow():
        if (r.get("是否通过") or "").strip() != "是":
            continue
        idx[(r.get("学号"), r.get("课程号"))] = r
    if not idx:
        return 0
    best = {}
    for r in rows:
        sid = (r.get("学号") or "").strip()
        for code in _row_codes(r):
            k = (sid, code)
            if k not in idx:
                continue
            tk = _term_key(r.get("学年学期")) or (0, 0)
            if k not in best or tk > best[k][0]:
                best[k] = (tk, r)
    for k, (_tk, r) in best.items():
        rec = idx[k]
        r["总成绩"] = (rec.get("新成绩") or "").strip() or "60"
        r["是否及格"] = "是"
        r["重修重考"] = rec.get("跟进方式") or "跟进"
    return len(best)


# ---------------------------------------------------------------- 主分析
def analyze_grades(fname, data):
    import xlsx_io
    header, rows = xlsx_io.read_table(fname, data)
    need = {"学年学期", "学号", "姓名", "课程名", "课程号", "是否及格", "总成绩", "学分", "专业"}
    missing_cols = need - set(header)
    if missing_cols:
        raise ValueError("成绩表缺少必需列：" + "、".join(sorted(missing_cols)))
    students = _read_students()
    roster = {s["学号"]: s for s in students}
    rows = [r for r in rows if (r.get("学号") or "").strip() in roster]
    if not rows:
        raise ValueError("成绩表中没有匹配到学生主表里的学号")
    _apply_follow(rows)                # 人工跟进的「已通过」先落到成绩行上
    fix_by_sid = follow_index()
    terms = sorted({_term_key(r["学年学期"]) for r in rows if _term_key(r["学年学期"])})
    if not terms:
        raise ValueError("成绩表学期列无法识别")
    cur = terms[-1]
    cur_label = term_label(cur)
    completed = terms[:-1]  # 已完成学期
    last_completed = completed[-1] if completed else None
    # 按专业：课程计划学期众数（用「入学后第几学年+秋/春」的相对口径，跨年级才对得齐）
    #          + 同届选课占比 + 课程学分
    byMajor = collections.defaultdict(list)
    for r in rows:
        byMajor[(r.get("专业") or "").strip()].append(r)
    mode_rel, pop, rep_row, code_credit = {}, {}, {}, {}
    natureCnt, name_of = collections.defaultdict(collections.Counter), {}
    # 全校口径（按课号）：某专业的方案课程在本专业没人修过时（新开课/其他专业方案），
    # 用其他专业成绩里的同一门课补「学分 / 课程性质 / 开课学期」，都是成绩表的实测值。
    g_credit, g_nature, g_term = {}, collections.defaultdict(collections.Counter), \
        collections.defaultdict(collections.Counter)
    # 课程号替代映射（成绩替代 / 成绩认定（方案外））：A<->B 互为等效课号。
    # 只取单跳（不做传递闭包），避免「体育专项1~4 全被并成一门」这类跨课串连。
    alias_cnt = collections.defaultdict(collections.Counter)
    for r in rows:
        cs = _row_codes(r)
        if len(cs) > 1:
            alias_cnt[cs[0]][cs[1]] += 1
            alias_cnt[cs[1]][cs[0]] += 1
    alias = {c: cnt.most_common(1)[0][0] for c, cnt in alias_cnt.items()}
    for major, rws in byMajor.items():
        n = max(1, len({r["学号"] for r in rws}))
        termCnt = collections.defaultdict(collections.Counter)
        whoCnt = collections.defaultdict(set)   # 修读该课的学生（重修重复行只算一人）
        for r in rws:
            tk = _term_key(r["学年学期"])
            if not tk:
                continue
            entry = int(((r["学号"] or "")[:4]) or 0)
            for code in _row_codes(r):
                termCnt[code][(tk[0] - entry, tk[1])] += 1
                whoCnt[code].add(r["学号"])
                rep_row.setdefault((major, code), r)
                natureCnt[(major, code)][(r.get("课程性质") or "").strip()] += 1
                name_of.setdefault((major, code), (r.get("课程名") or "").strip())
                code_credit[(major, code)] = max(code_credit.get((major, code), 0.0),
                                                 _f(r.get("学分")))
                g_credit[code] = max(g_credit.get(code, 0.0), _f(r.get("学分")))
                nat = (r.get("课程性质") or "").strip()
                if nat:
                    g_nature[code][nat] += 1
                g_term[code][tk] += 1
        for code, cnt in termCnt.items():
            mode_rel[(major, code)] = cnt.most_common(1)[0][0]
            pop[(major, code)] = min(1.0, len(whoCnt[code]) / float(n))
    plans = load_plans()["majors"]
    # 无选择余地的课组：组内全部课程学分合计 ≈ 要求学分（没有可替换的课）。
    # 这类课组里的「必修」课程必须逐门修，学分凑够不能替代（军事技能、体育俱乐部1 等）。
    leaf_sum, leaf_req = {}, {}
    for _mj, _pl in plans.items():
        _codes = _pl.get("codes") or {}
        if not _codes:
            continue
        _pn = {_norm_name(v) for v in _codes.values()}
        _cg = _pl.get("cgroups") or {}
        for _c in _pl.get("cats", []):
            leaf_req[(_mj, _c["key"])] = _c["req"]
        for _code, _nm in _codes.items():
            # 归口课组以方案 cgroups 为准（不依赖成绩表）；学分用成绩实测，缺时用方案自带
            _gs = _cg.get(_code) or []
            if _gs:
                _k = _gs[0]
            else:
                _rr = rep_row.get((_mj, _code))
                if _rr is None:
                    continue
                _k = " > ".join(_refine_cat(_cat_of_row(_rr, _codes, _pn), _rr, _cg))
            _cr = code_credit.get((_mj, _code), 0.0)
            if _cr <= 0:
                _cr = (_pl.get("courses") or {}).get(_code, {}).get("c") or 0.0
            leaf_sum[(_mj, _k)] = round(leaf_sum.get((_mj, _k), 0.0) + _cr, 2)

    def no_choice(mj, key):
        s, rq = leaf_sum.get((mj, key)), leaf_req.get((mj, key))
        return s is not None and rq is not None and s <= rq + 0.05

    per = {}
    for sid, stu in roster.items():
        major = (stu.get("专业") or "").strip()
        my = [r for r in rows if r["学号"] == sid]
        if not my:
            continue
        entry = int(((stu.get("学号") or "")[:4]) or 0)

        def plan_term(code, min_pop=0.2):
            # 计划学期：培养方案自带的「计划学年学期」是教务数据，最权威；
            # 方案没给时才用同专业同学的实际修读学期众数推断，
            # 且只有"多数同届人都修过"的课，众数才可信（否则视为开课时间未知）。
            #
            # 注意：方案里的学年是「方案那一届」的绝对学年（2023 级方案的
            # 第一学期写 2023-2024 秋），而同一份方案会被 2021~2024 各届学生共用。
            # 所以必须先换算成「第几学年」的相对偏移，再加上该生自己的入学年，
            # 否则 2024 级学生会把方案里"更早的学年"当成已完成，未来课也会被点名。
            tk = _plan_term_of((plan.get("courses") or {}).get(code, {}).get("t"))
            # 方案自带的计划学期是绝对学年，直接用：留级/休学/跨届学生目前一律
            # 按方案那一级算（用户约定），不做入学年换算，避免把新方案才有的课
            # 当成"早就该修却没修"整届误报。
            if tk is not None:
                return tk
            if pop.get((major, code), 0) < min_pop:
                return None
            rel = mode_rel.get((major, code))
            return (entry + rel[0], rel[1]) if rel is not None else None

        def plan_credit(code):
            # 课程学分：本专业成绩实测 > 全校同课号实测 > 方案自带（方案导出表学分列常缺）
            c = code_credit.get((major, code), 0.0) or g_credit.get(code, 0.0)
            if c > 0:
                return c
            return (plan.get("courses") or {}).get(code, {}).get("c") or 0.0

        taken, passed_best, rep_pass = set(), {}, {}
        pass_name = {}  # 归一化课程名 -> 最近通过学期（重修后课程号会变，只能按名匹配）
        pass_rep = {}   # 归口课号 -> 最近通过学期（成绩替代/认定：课号与课名都可能变）
        pass_code = {}  # 课号（含替代课号）-> 最近通过学期，用于识别"换新号重修通过"
        failed, current_rows = [], []
        plan = plans.get(major) or {}
        plan_codes = plan.get("codes") or {}
        plan_names = {_norm_name(v) for v in plan_codes.values()}
        cg = plan.get("cgroups") or {}

        def _refine(r):
            return _refine_cat(_cat_of_row(r, plan_codes, plan_names), r, cg)

        hit_best = {}   # 归口课号 -> (学分, 成绩行)：类别学分按归口课号累计，同一门课只算一次
        for r in my:
            cs = _row_codes(r)
            if not cs:
                continue
            code = cs[0]
            alt = cs[1] if len(cs) > 1 else ""
            # 归口课号：成绩替代/成绩认定（方案外）的行，课程号是实际修读的课，
            # 替代课程号才是培养方案里的课。若只有替代课号在方案内，就按它归口。
            rep = alt if (alt in plan_codes and code not in plan_codes) else code
            tk = _term_key(r["学年学期"])
            nm = (r.get("课程名") or "").strip()
            credit = _f(r.get("学分"))
            sc = (r.get("总成绩") or "").strip()
            isFail = _fail(r)
            if tk == cur:
                # 当前学期：成绩还没出的算"在修"；总成绩已出且不及格的按挂科算
                # （教务成绩表里当前学期的课一旦有了总成绩，就是已出分的真实结果）
                current_rows.append({"n": nm, "s": sc, "c": credit, "ok": not isFail})
                taken.update(cs)
                if credit > 0 and not (isFail and sc):
                    passed_best[code] = credit
                    rep_pass[code] = r
                    if credit > hit_best.get(rep, (0,))[0]:
                        hit_best[rep] = (credit, r)
                    for c in cs:
                        if tk > pass_code.get(c, (0, 0)):
                            pass_code[c] = tk
                elif isFail and sc:
                    cat = _refine(r)
                    if not _is_flex_cat(cat):
                        failed.append({"t": term_label(tk), "n": nm, "s": sc, "c": credit,
                                       "code": code, "rep": rep, "cat": " > ".join(cat),
                                       "way": (r.get("重修重考") or "").strip()})
                continue
            if tk is None or tk > cur:
                continue
            taken.update(cs)
            if isFail:
                cat = _refine(r)
                if _is_flex_cat(cat):
                    continue  # 文化素质/外语/学科交叉：课组内补足即可，不计入挂科、不列明细
                item = {"t": term_label(tk), "n": nm, "s": sc, "c": credit, "code": code,
                        "rep": rep, "cat": " > ".join(cat),
                        "way": (r.get("重修重考") or "").strip()}
                if not any(f["n"] == nm and f["t"] == item["t"] and f["s"] == sc for f in failed):
                    failed.append(item)
            elif credit > 0 and credit > passed_best.get(code, -1):
                passed_best[code] = credit
                rep_pass[code] = r
                if credit > hit_best.get(rep, (0,))[0]:
                    hit_best[rep] = (credit, r)
                nn = _norm_name(nm)
                if nn and tk > pass_name.get(nn, (0, 0)):
                    pass_name[nn] = tk
                if tk > pass_rep.get(rep, (0, 0)):
                    pass_rep[rep] = tk
                for c in cs:
                    if tk > pass_code.get(c, (0, 0)):
                        pass_code[c] = tk
        # 重修/补考/成绩替代已通过的课程不再算挂科（同一门课只留最近一次挂科）
        if failed:
            keep = {}
            for f in failed:
                nn = _norm_name(f["n"])
                ft = _term_key(f["t"]) or (0, 0)
                alt = alias.get(f["code"])
                # 同一门课（成绩替代/认定）只要曾经通过，学分就已获得，后来的重修挂科不算未获学分
                if alt and alt in pass_code:
                    continue
                done = max(pass_name.get(nn, (0, 0)), pass_rep.get(f["rep"], (0, 0)),
                           pass_code.get(alt, (0, 0)))
                if done > ft:
                    continue  # 该课之后已重修/补考通过，或被成绩替代认定，不算挂科
                old = keep.get(nn)
                if old is None or ft > (_term_key(old["t"]) or (0, 0)):
                    keep[nn] = f
            failed = sorted(keep.values(), key=lambda f: (_term_key(f["t"]) or (0, 0)), reverse=True)
        # 累计学分：按归口课号去重（hit_best 的 key），和类别学分同一口径。
        # 不能按实际课号累加 passed_best——「方案课 + 替代课」都通过时同一门课会被算两次。
        earned = round(sum(credit for credit, _ in hit_best.values()), 1)
        # 类别学分累计（按归口课号：替代/认定过的课算到方案内那门课上，同一门课只计一次）
        catEarned = collections.defaultdict(float)
        catCount = collections.defaultdict(int)   # 同一课组已修门数（界面标「已得 X 学分 · N 门」，防把学分看成门数）
        extra = []
        en_credits = 0.0   # 英语学分合池（按归口课号去重，防替代课号重复计分）
        for rep, (credit, r) in hit_best.items():
            cat = _refine(r)
            if cat[0] == "学科交叉课组":
                extra.append({"n": (r.get("课程名") or "").strip(), "c": credit, "t": term_label(tk_of(r))})
                catEarned["学科交叉课组"] += credit
                catCount["学科交叉课组"] += 1
                continue
            # 方案外课程（归口课号不在方案、课名也对不上方案里的课）：只有「自由选课池」
            # 类课组（文化素质/外语/艺术类/交叉）才能按同类计入；固定课组（体育专项、
            # 体育俱乐部、国防教育…）必须修方案里那几门，否则会给缺口课组凑出假学分。
            cat_key = " > ".join(cat)
            in_plan = (rep in plan_codes) or (_norm_name(r.get("课程名") or "") in plan_names)
            if not in_plan and not _is_flex_cat(cat_key):
                continue
            if RE_EN.search((r.get("课程名") or "")):
                en_credits += credit
            for i in range(1, len(cat) + 1):
                catEarned[" > ".join(cat[:i])] += credit
                catCount[" > ".join(cat[:i])] += 1
        # 方案缺口（叶子层）
        reqs = {c["key"]: c["req"] for c in plan.get("cats", [])}
        en_credits = round(en_credits, 1)
        cg_leaf = set()
        for _ps in cg.values():
            cg_leaf.update(_ps)
        short_all = []
        for key, req in sorted(reqs.items()):
            # 课组归属已从培养方案解析出来的子课组（专业限选/专业必修的模块）也能按学分判定
            if not _checkable(key) and key not in cg_leaf:
                continue
            got = round(catEarned.get(key, 0.0), 1)
            # 外语基础组要修的是「同一语种的基础课序列」（如通用英语Ⅰ-1、Ⅰ-2），
            # 拓展课（英语实践2、大学英语写作…）不能顶替缺的那门基础课；
            # 外语拓展组也不能反过来用基础课凑——否则只修了 Ⅰ-1 的学生会被 Ⅰ-4 +
            # 英语实践2 凑满 8 分，基础组缺口永远不提示。两个组都按方案归口统计。
            if got < req - 0.05:
                short_all.append({"k": key, "req": req, "got": got})
        # 父组（汇总组）的缺口若完全由子组造成，就只报子组、不报父组：
        # 否则「通识限选差 2 学分」看不出差在哪（实际是外语拓展组少 2 分），
        # 容易被当成四史/文化素质的学分没算进去。
        short = []
        for s in short_all:
            kids = [o for o in short_all if o["k"].startswith(s["k"] + " > ")]
            if kids and sum(o["req"] - o["got"] for o in kids) >= (s["req"] - s["got"]) - 0.05:
                continue
            short.append(s)
        # 修业年限：默认按学号入学年推算；休学/降级等以学生主表的「在校学年」为准
        yv = max(1, cur[0] - entry + 1)
        yv_manual = ""
        try:
            _sy = int(float(str(stu.get("在校学年") or "").strip() or 0))
        except (TypeError, ValueError):
            _sy = 0
        if 1 <= _sy <= 6:
            yv_manual = _sy
            yv = _sy
        # 应修未修：按培养方案课组口径 —— 课组学分没修满的，才需要再修；
        # 组内没修的课按学期顺序列出来，学分凑够缺口即止（组里修够了的，
        # 剩下的课不属于"必须修"，不再点名）。是否开课不看同届修了多少人。
        missing = []
        if plan.get("codes") and last_completed:
            shortfall_keys = {s["k"] for s in short}
            # 该生已修过但挂科的课（含其替代课号）：挂科明细里会列，这里不重复点名"未选课"
            failed_codes = {f["code"] for f in failed} | {alias.get(f["code"]) for f in failed}

            def cat_of_code(code):
                """方案内课程的归口：以方案自身的课组映射为准（通用：不依赖成绩表，
                全专业没人修过的课一样能归口）；成绩行归口仅作后备。"""
                for g in cg.get(code) or []:
                    if g in reqs:
                        return g.split(" > ")
                rr = rep_row.get((major, code))
                if rr is not None:
                    return _refine(rr)
                g0 = (cg.get(code) or [""])[0]
                return g0.split(" > ") if g0 else None

            bygrp = {}
            for code, nm in plan["codes"].items():
                if code in taken or code in failed_codes:
                    continue
                cat = cat_of_code(code)
                if cat is None or _is_flex_cat(cat):
                    continue  # 文化素质/外语/学科交叉属自由选课凑学分，只记课组学分缺口
                key = " > ".join(cat)
                hit = _short_key_of(key, shortfall_keys)
                if hit is None:
                    continue
                # 叶子组自己有独立学分要求、能按成绩归口核对、且已经达标时，
                # 不拿它的课去凑上层汇总组的缺口：
                # 四史类 2 分修够后，顶层「通识限选」差 2 分不该再点名另外三门四史
                #（四史类是自由选课达标即可；上层缺口由说明里的学分缺口提示）。
                if hit != key and key in reqs and _checkable(key) and not _is_flex_cat(key):
                    continue
                mt = plan_term(code)
                if mt is None or mt > last_completed:
                    continue
                cr = plan_credit(code)
                bygrp.setdefault(key, []).append(
                    {"n": nm, "cat": key, "pop": round(pop.get((major, code), 0) * 100),
                     "term": term_label(mt), "c": round(cr, 1) if cr else "",
                     "_tk": mt, "_cr": cr})
            for key, items in bygrp.items():
                # 必修大类（根类名含「必修」，来自方案本身）或「组内课程学分合计≈要求学分」
                # （没有可选余地，纯数学判定）的课组：方案里列出的课逐门点名，
                # 学分被别的课凑够也不能抵（军事技能、体育俱乐部1 都曾从这里漏掉）。
                # 限选类课组仍有选择余地：按学分缺口列到凑够为止。
                must_all = key.split(" > ")[0] in ("通识必修", "学科必修", "专业必修") \
                           or no_choice(major, key)
                # 缺口取「命中的缺口组」（自身或最近祖先）的差值：
                # 用祖先缺口限额时，列到凑够就停，不再把整组没修的课全列出来。
                # 注意必须按当前 key 重算命中缺口——此前误用了上面课程循环残留的
                # hit（最后一个课号的命中），这里拿到的 gap 基本是 None，
                # 「凑够即止」失效，missing 会把整组没修的课全列出来。
                hit = _short_key_of(key, shortfall_keys)
                gap = next((round(s2["req"] - s2["got"], 1) for s2 in short if s2["k"] == hit), None)
                acc = 0.0
                for x in sorted(items, key=lambda x: (x["_tk"], -x["pop"])):
                    missing.append({k2: v for k2, v in x.items() if not k2.startswith("_")})
                    acc = round(acc + x["_cr"], 1)
                    if not must_all and x["_cr"] > 0 and gap is not None and acc >= gap - 0.05:
                        break   # 限选课组只差 N 学分时，列出能凑够 N 学分的课就够了
            missing.sort(key=lambda x: (-x["pop"], x["cat"]))
            missing = missing[:40]
        # 未获学分科目（截止目前应获未获）
        # 只有文化素质/学科交叉/外语基础/外语拓展 四类以「课程组」形式展示缺口；
        # 其余类别一律展开成具体课程名（挂科未通过 / 该修未选），按缺口学分逐条消费。
        def _leaf_of(k):
            return k.split(" > ")[-1]

        flex_short = [s for s in short if _leaf_of(s["k"]) in FLEX_GROUPS]
        strict_short = [s for s in short if _leaf_of(s["k"]) not in FLEX_GROUPS]
        unearned_g = [{"n": _leaf_of(s["k"]), "k": s["k"], "req": s["req"], "got": s["got"],
                       "cnt": catCount.get(s["k"], 0),
                       "short": round(s["req"] - s["got"], 1)}
                      for s in sorted(flex_short, key=lambda x: -(x["req"] - x["got"]))]
        unearned_c = []
        for s in sorted(strict_short, key=lambda x: x["k"]):
            key, gap = s["k"], round(s["req"] - s["got"], 1)
            cands = []
            for f in failed:
                if f["cat"] == key or f["cat"].startswith(key + " > "):
                    cands.append({"n": f["n"], "cat": f["cat"], "status": "挂科 %s 分" % f["s"],
                                  "t": f["t"], "c": f["c"], "_ord": 0,
                                  "_tk": _term_key(f["t"]) or (0, 0), "_cr": _f(f["c"])})
            for m in missing:
                if (m["cat"] == key or m["cat"].startswith(key + " > ")) \
                        and not any(x["n"] == m["n"] for x in cands):
                    cands.append({"n": m["n"], "cat": m["cat"], "status": "未选课",
                                  "t": m.get("term") or "", "c": m.get("c") or "", "_ord": 1,
                                  "_tk": _term_key(m.get("term")) or (0, 0), "_cr": _f(m.get("c"))})
            cands.sort(key=lambda x: (x["_ord"], x["_tk"]))
            acc = 0.0
            for x in cands:
                out = {k2: v for k2, v in x.items() if not k2.startswith("_")}
                # 「限选类」课组（根类不是必修大类，且组内学分合计>要求学分）：
                # 组内任修哪门都行，点到名的这门只是按修读热度给的推荐 →
                # 前端在类别旁打 * 提示。挂科课必须本人重修，不算推荐。
                if x["status"] == "未选课" \
                        and x["cat"].split(" > ")[0] not in ("通识必修", "学科必修", "专业必修") \
                        and not no_choice(major, x["cat"]):
                    out["rec"] = True
                unearned_c.append(out)
                acc = round(acc + x["_cr"], 1)
                if acc >= gap - 0.05:
                    break
# 挂科就是挂科：方案内已挂的课即使所属课组学分已被别的课凑够，也要列出来
        # （技术经济 55 分这种，不能因为「综合拓展类已满」就不显示）
        _listed_n = {x.get("n") for x in unearned_c}
        for f in failed:
            if f["n"] in _listed_n:
                continue
            unearned_c.append({"n": f["n"], "cat": f["cat"], "status": "挂科 %s 分" % f["s"],
                               "t": f["t"], "c": f["c"]})
        unearned_g.sort(key=lambda x: -x["short"])
        # 警示：三类，从重到轻 —— 退学警示 / 留级警示 / 学业警示
        #   退学、留级：只看累计学分有没有达到该年级的学分线（学校的学业预警标准）
        #   学业      ：按培养方案口径 —— 该修的课没修（应修未修），或挂了方案内的课
        #               且这门课所属课组的学分还没凑够。课组学分已达标（例如文化素质类
        #               用别的课补足了），那么这一门挂科不影响毕业，不予提醒。
        warn = {"type": "", "level": "", "note": ""}
        failed_n = len(failed)
        # 第 6 学年及以上沿用第 5 学年标准（该年级只有退学线），避免高年级无标准可依
        th = YEAR_THRESHOLDS.get(yv) or YEAR_THRESHOLDS.get(5) or (None, None, None)
        notes = []
        # 自由课组（文化素质/外语/交叉）学分没凑够：属"应获未获"，在说明里点出缺口
        for sk in short:
            leaf = sk["k"].split(" > ")[-1]
            if leaf in FLEX_GROUPS:
                notes.append("%s差 %.1f 学分（已得 %.1f/%g 学分 · %d 门课）"
                             % (leaf, sk["req"] - sk["got"], sk["got"], sk["req"],
                                catCount.get(sk["k"], 0)))
        # 方案内仍未拿到的课：挂了、且所属课组仍处于缺口状态
        short_keys = {s["k"] for s in short}
        fail_short = [f for f in failed if _short_hit(f.get("cat"), short_keys)]
        if th[2] is not None and earned < th[2]:
            warn["type"] = "退学警示"
            warn["level"] = "累计 %.1f 学分，低于退学线 %g 分" % (earned, th[2])
        elif th[1] is not None and earned < th[1]:
            warn["type"] = "留级警示"
            warn["level"] = "累计 %.1f 学分，低于留级线 %g 分" % (earned, th[1])
        elif fail_short or missing:
            warn["type"] = "学业警示"
            bits = []
            if fail_short:
                bits.append("挂科 %d 门" % len(fail_short))
            if missing:
                bits.append("应修未修 %d 门" % len(missing))
            warn["level"] = " · ".join(bits)
        # 学分缺口照旧写进说明（表格"说明"列会显示）
        if th[2] is not None and earned < th[2]:
            notes.insert(0, "累计学分 %.1f 低于 %g 分（退学警示线）" % (earned, th[2]))
        if th[1] is not None and earned < th[1]:
            notes.insert(0, "累计学分 %.1f 低于 %g 分（留级线）" % (earned, th[1]))
        if warn["type"] == "学业警示" and not notes:
            todo = [f.get("n") for f in fail_short][:2] + [m.get("n") for m in missing][:2]
            todo = [t for t in todo if t]
            if todo:
                more = (len(fail_short) + len(missing)) > len(todo)
                notes.append("待完成：" + "、".join(todo) + ("…" if more else ""))
        if not warn["type"] and yv >= 4 and not notes:
            notes.append("第四学年：须修满培养方案全部学分方可毕业")
        warn["note"] = "；".join(notes)
        per[sid] = {
            "name": stu.get("姓名", ""), "major": major, "class": stu.get("班级", ""),
            "year": yv, "earned": earned,
            "yearAuto": max(1, cur[0] - entry + 1), "yearManual": yv_manual,
            "failed": failed, "missing": missing, "extra": extra,
            "short": short,
            "unearned": {"g": unearned_g, "c": unearned_c},
            "cats": {k: round(v, 1) for k, v in catEarned.items() if v > 0},
            "cur": current_rows, "warn": warn,
            "fix": fix_by_sid.get(sid, []),
        }
    # 方案课程明细：每门课的归口课组 + 计划学期 + 学分 + 课程性质 + 同专业修读率
    # 全部来自「导入的方案文件」与「导入的成绩表」两个实测来源：
    #   归口/计划学期/课程性质 优先用方案自带（不依赖有没有学生成绩），
    #   学分/修读率/实际修读学期 用成绩表实测补充；两边都没有的字段留空。
    plan_courses = {}
    for major, plan in plans.items():
        codes = plan.get("codes") or {}
        if not codes:
            continue
        pcg = plan.get("cgroups") or {}
        py = int(plan.get("year") or 0) or 0
        d = {}
        for code, nm in codes.items():
            pc = (plan.get("courses") or {}).get(code) or {}
            rr = rep_row.get((major, code))
            rel = mode_rel.get((major, code))
            cat = pc.get("cat") or ""
            if not cat and rr is not None:
                pn = {_norm_name(v) for v in codes.values()}
                on = any(c in codes for c in _row_codes(rr))
                cat = " > ".join(_refine_cat(_cat_of_row(rr, codes if on else None, pn), rr, pcg))
            t = pc.get("t") or ""
            if not t and rel and py:
                t = term_label((py + rel[0], rel[1]))
            if not t and g_term.get(code):
                t = term_label(g_term[code].most_common(1)[0][0])   # 全校同一门课的实际开课学期
            c = round(code_credit.get((major, code), 0.0), 1)
            if not c:
                c = round(g_credit.get(code, 0.0), 1)
            if not c:
                c = pc.get("c") or 0.0
            k = pc.get("k") or ""
            if not k and natureCnt.get((major, code)):
                k = natureCnt[(major, code)].most_common(1)[0][0]
            if not k and g_nature.get(code):
                k = g_nature[code].most_common(1)[0][0]
            d[code] = {
                "n": nm or name_of.get((major, code), ""),
                "c": c, "t": t, "k": k,
                "cat": cat,
                "pop": round(pop.get((major, code), 0) * 100),
            }
        plan_courses[major] = d
    out = {
        "importedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "source": os.path.basename(fname),
        "currentTerm": cur_label,
        "students": per,
        "planCourses": plan_courses,
    }
    os.makedirs(DATA, exist_ok=True)
    with open(ANALYSIS_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    return out


def _is_flex_cat(cat):
    """自由选课凑学分的课组：文化素质类、外语课组、学科交叉。
    挂科不强制重修，课组总学分达标即可，不点名具体课程。"""
    return cat and any(g in cat for g in FLEX_GROUPS)


def tk_of(r):
    tk = _term_key(r.get("学年学期"))
    return tk if tk else (0, 0)


def _checkable(key):
    """该类别能否由成绩单归口学分：专业必修/专业限选只查顶级，子模块成绩行无法区分。"""
    top = key.split(" > ")[0]
    if top in ("专业必修", "专业限选"):
        return " > " not in key
    return top in GRADED_TOPS


def _is_flex_cat(cat):
    """自由选课凑学分的课组：文化素质类、外语课组、学科交叉。
    挂科不强制重修，课组总学分达标即可，不点名具体课程。"""
    return cat and any(g in cat for g in FLEX_GROUPS)


def tk_of(r):
    tk = _term_key(r.get("学年学期"))
    return tk if tk else (0, 0)


def _checkable(key):
    """该类别能否由成绩单归口学分：专业必修/专业限选只查顶级，子模块成绩行无法区分。"""
    top = key.split(" > ")[0]
    if top in ("专业必修", "专业限选"):
        return " > " not in key
    return top in GRADED_TOPS


def _refine_cat(cat, r, cg):
    """用培养方案的课组归属，把「专业限选」这类顶层归口细化到具体课组。

    培养方案里每门课挂在哪个课组（cgroups）是教务的权威归口，直接以它为准：
    成绩行只要课号（或替代课号）在方案里，就按方案的完整路径归类，不再依赖
    按课名猜出来的顶层类别——猜错了会把「体育俱乐部1」挂成学科交叉，
    导致课组学分永远凑不满或永远凑满。
    """
    if not cg:
        return cat
    best = None
    for c in _row_codes(r):
        for g in cg.get(c) or []:
            gp = g.split(" > ")
            if best is None or len(gp) > len(best):
                best = gp
    return best or cat


def _parents(key):
    parts = key.split(" > ")
    return {" > ".join(parts[:i]) for i in range(1, len(parts))}


def _short_hit(cat, short_keys):
    """课程所属课组是否真的缺学分。

    只看「自身课组」和「自己下面的子课组」，不看上层汇总组：
    例如「通识限选 > 体育类-限选 > 体育俱乐部」缺学分时，同属通识限选的
    「四史类」并不缺 —— 四史类四门课修过任意一门（2 学分）即达标，不应被点名。
    """
    return _short_key_of(cat, short_keys) is not None


def _short_key_of(cat, short_keys):
    """返回覆盖该课组的缺口 key（自身或最近的祖先组）；不在任何缺口里则 None。"""
    cat = cat or ""
    if not cat:
        return None
    if cat in short_keys:
        return cat
    parts = cat.split(" > ")
    for i in range(len(parts) - 1, 0, -1):
        k = " > ".join(parts[:i])
        if k in short_keys:
            return k
    return None


def load_academic():
    """供 build.py 汇总：返回嵌入看板的数据。"""
    plans = load_plans()
    ana = {}
    if os.path.exists(ANALYSIS_FILE):
        with open(ANALYSIS_FILE, encoding="utf-8") as f:
            ana = json.load(f)
    return {"plans": plans.get("majors", {}), "analysis": ana, "warnings": read_warnings(),
            "follows": read_follow(),
            "thresholds": {str(k): list(v) for k, v in YEAR_THRESHOLDS.items()}}
