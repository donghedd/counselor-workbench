# -*- coding: utf-8 -*-
"""学业模块：培养方案解析 + 全校成绩导入分析。

数据文件（均在 data/ 下）：
- 12_培养方案.json   各专业培养方案（类别树 + 要求学分 + 方案内课程清单）
- 13_成绩分析.json   最近一次全校成绩导入的逐生分析（挂科/应修未修/学科交叉/学分/警示）
- 14_学业警示.csv    人工登记的学业警示（自动分析不落这张表，前端合并展示）

培养方案 xlsx 是教务系统导出的特殊版式（合并单元格交错），无法按普通表格读。
本模块直接解析 xlsx 的 sharedStrings 与单元格流：
- 类别树来自「类别#g（要求学分：X 已修学分：Y）g#/子类」编码串与面包屑碎片；
- 方案内课程集合来自「课程名[课程号]」形 token。
两者都只取"要求学分"与"课程号集合"，不依赖排版顺序，鲁棒。

成绩分析口径（大四上=当前学期进行中）：
- 挂科：已完成学期中 成绩<60 / 文本"不及格" / 是否及格=否且学分>0（六级等 >=300 大分不算）；
- 成绩替代 / 成绩认定（方案外）：这类行的「替代课程号」才是培养方案内的课号（占 152/156 条），
  视同该课已修读——计入 taken（不再误报未选课）、学分归口到方案内课号；
  同一门课（含替代课号）任一次学期通过即视为已获学分，之后重修再挂不算未获学分；
- 已修学分：已完成/在修学期中通过课程学分之和（同课程号取最大学分一次）；
- 应修未修：方案内课程按同专业同届修读学期众数推断计划学期，落在已完成学期、
  该生从未选过、所在类别学分有缺口、且 >=30% 同专业人选过的课程；
  文化素质类（通识限选>文化素质类-计科）与学科交叉课组属"自由选课凑学分"课组，
  只记学分缺口、不点名具体课程（不必修某一门，达标即可）；
- 学科交叉：培养方案外的课程（课号不在本专业方案、课名也对不上方案课程），
  只要不是文化素质类（含校公选课类别）、体育专项/俱乐部、四史、外语、必修与专业类课程，
  学分一律计入"学科交叉课组"——不限于方案里列出的那几门交叉课；
- 挂科：文化素质类 / 学科交叉课程未通过标记 flex（不强制重修，可在课组内另选补足学分），
  不计入警示挂科门数；其余必修/限选课程挂科才计入警示；
- 学分警示：按修业年限阈值表（1:30/25/15 2:70/55/45 3:110/85/75 4:留级115 退学105）；
- 警示类型：大四挂科=毕业警示，大一~大三挂科=留级警示；累计学分不足线另注在说明里。
"""
import csv, datetime, io, json, os, re, zipfile, collections
import xml.etree.ElementTree as ET

from app_paths import APP_HOME as HERE

DATA = os.path.join(HERE, "data")
PLAN_FILE = os.path.join(DATA, "12_培养方案.json")
ANALYSIS_FILE = os.path.join(DATA, "13_成绩分析.json")
WARN_FILE = os.path.join(DATA, "14_学业警示.csv")
WARN_HEADER = ["记录ID", "学号", "姓名", "警示类型", "级别", "原因", "学期", "来源", "更新时间"]

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
# 英语类课组：所有英语课学分合池，达标即可（外语基础 8 分 / 外语拓展 4 分）
EN_GROUPS = ("外语基础课组I类", "外语拓展课组I类")
RE_EN = re.compile(r"英语|外语")
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
def parse_plan(fname, data):
    strings, cells = _read_strings_and_cells(data)
    tree = {}          # path tuple -> req
    parent_of = {}     # name -> parent path tuple（一名不挂两父）
    # 1) 编码串（sharedStrings 全量，覆盖未被单元格引用的根类别）
    for v in strings:
        sv = v.strip()
        if "要求学分" not in sv or sv.startswith("（要求学分"):
            continue  # 纯学分碎片无名称，交给步骤2
        path = parse_path(v)
        if not path:
            continue
        for k in range(1, len(path) + 1):
            key = tuple(p[0] for p in path[:k])
            req = path[k - 1][1]
            if k > 1:
                par = parent_of.get(key[-1])
                if par is not None and par != key[:-1]:
                    continue  # 同名类挂到不同父级：跳过，防串块噪声
                parent_of[key[-1]] = key[:-1]
            tree[key] = max(tree.get(key, 0), req)
    # 2) 面包屑碎片流：根名称 + （要求学分）配对，补充树（如 专业必修>专业课）
    cur = []
    for v in cells:
        s = v.strip()
        if s in SCAFFOLD or RE_COURSE.match(s) or RE_TERM.match(s) or re.match(r"^[\d.]+$", s) or s in NATURE_SET:
            continue
        m = RE_CREDIT.search(s)
        if m:
            for i in range(1, len(cur) + 1):
                key = tuple(cur[:i])
                if len(key) > 3:
                    break
                if i > 1:
                    par = parent_of.get(key[-1])
                    if par is not None and par != key[:-1]:
                        break
                    parent_of[key[-1]] = key[:-1]
                if key in tree:
                    continue  # 编码串（步骤1）为准，碎片只补缺
                tree[key] = float(m.group(1))
            continue
        name = s.lstrip("/").strip()
        name = _strip_multi(name)
        if not name or name in STOPWORDS:
            continue
        if s.startswith("/") and cur:
            if name not in cur:
                cur.append(name)
        elif RE_NAME.match(name):
            cur = [name]
        # 其余（空串、单字符等）忽略
    # 3) 方案内课程集合
    codes = {}
    for v in list(cells) + strings:
        m = RE_COURSE.match(v.strip())
        if m:
            code = m.group(1)
            name = re.sub(r"\[\d{6,}\]$", "", v.strip()).strip()
            codes.setdefault(code, name)
    cats = [{"key": " > ".join(key), "path": list(key), "req": tree[key]}
            for key in sorted(tree, key=lambda p: (len(p), p))]
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
            "source": os.path.basename(fname)}


def load_plans():
    if not os.path.exists(PLAN_FILE):
        return {"majors": {}}
    with open(PLAN_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_plan(fname, data):
    plan = parse_plan(fname, data)
    plans = load_plans()
    old = plans["majors"].get(plan["major"], {})
    for c in plan["cats"]:
        for oc in old.get("cats", []):
            if oc["key"] == c["key"] and oc.get("reqOverride"):
                c["req"] = oc["req"]
                c["reqOverride"] = True
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
    （不再只认方案里列的那几门交叉课）。但以下仍按原体系归口：
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
                    or cc in ("通识必修", "学科必修", "专业必修", "专业选修", "专业限选", "学科限选")):
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
        for code, cnt in termCnt.items():
            mode_rel[(major, code)] = cnt.most_common(1)[0][0]
            pop[(major, code)] = min(1.0, len(whoCnt[code]) / float(n))
    plans = load_plans()["majors"]
    per = {}
    for sid, stu in roster.items():
        major = (stu.get("专业") or "").strip()
        my = [r for r in rows if r["学号"] == sid]
        if not my:
            continue
        entry = int(((stu.get("学号") or "")[:4]) or 0)

        def plan_term(code, min_pop=0.2):
            # 只有"多数同届人都修过"的课，其学期众数才可信；否则视为开课时间未知/未来课程
            if pop.get((major, code), 0) < min_pop:
                return None
            rel = mode_rel.get((major, code))
            return (entry + rel[0], rel[1]) if rel is not None else None

        def plan_credit(code):
            return code_credit.get((major, code), 0.0)

        taken, passed_best, rep_pass = set(), {}, {}
        pass_name = {}  # 归一化课程名 -> 最近通过学期（重修后课程号会变，只能按名匹配）
        pass_rep = {}   # 归口课号 -> 最近通过学期（成绩替代/认定：课号与课名都可能变）
        pass_code = {}  # 课号（含替代课号）-> 最近通过学期，用于识别"换新号重修通过"
        failed, current_rows = [], []
        plan_codes = (plans.get(major) or {}).get("codes") or {}
        plan_names = {_norm_name(v) for v in plan_codes.values()}
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
                current_rows.append({"n": nm, "s": sc, "c": credit, "ok": not isFail})
                taken.update(cs)
                if credit > 0:  # 在修课程直接计入已得学分（不再当作"未计入"）
                    passed_best[code] = credit
                    rep_pass[code] = r
                    if credit > hit_best.get(rep, (0,))[0]:
                        hit_best[rep] = (credit, r)
                    for c in cs:
                        if tk > pass_code.get(c, (0, 0)):
                            pass_code[c] = tk
                continue
            if tk is None or tk > cur:
                continue
            taken.update(cs)
            if isFail:
                cat = _cat_of_row(r, plan_codes, plan_names)
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
        earned = round(sum(passed_best.values()), 1)
        # 类别学分累计（按归口课号：替代/认定过的课算到方案内那门课上，同一门课只计一次）
        catEarned = collections.defaultdict(float)
        extra = []
        en_credits = 0.0   # 英语学分合池（按归口课号去重，防替代课号重复计分）
        for rep, (credit, r) in hit_best.items():
            cat = _cat_of_row(r, plan_codes, plan_names)
            if cat[0] == "学科交叉课组":
                extra.append({"n": (r.get("课程名") or "").strip(), "c": credit, "t": term_label(tk_of(r))})
                catEarned["学科交叉课组"] += credit
                continue
            if RE_EN.search((r.get("课程名") or "")):
                en_credits += credit
            for i in range(1, len(cat) + 1):
                catEarned[" > ".join(cat[:i])] += credit
        # 方案缺口（叶子层）
        plan = plans.get(major) or {}
        reqs = {c["key"]: c["req"] for c in plan.get("cats", [])}
        en_credits = round(en_credits, 1)
        short_all = []
        for key, req in sorted(reqs.items()):
            if not _checkable(key):
                continue
            got = round(catEarned.get(key, 0.0), 1)
            if key.split(" > ")[-1] in EN_GROUPS:
                got = en_credits
            if got < req - 0.05:
                short_all.append({"k": key, "req": req, "got": got})
        short = [s for s in short_all if not any(o["k"].startswith(s["k"] + " > ") for o in short_all)]
        # 修业年限
        yv = max(1, cur[0] - entry + 1)
        # 应修未修
        missing = []
        if plan.get("codes") and last_completed:
            shortfall_keys = {s["k"] for s in short}
            shortfall_keys |= {p for k in shortfall_keys for p in _parents(k)}
            cand = {}
            # 该生已修过但挂科的课（含其替代课号）：挂科明细里会列，这里不重复点名"未选课"
            failed_codes = {f["code"] for f in failed} | {alias.get(f["code"]) for f in failed}
            for code, nm in plan["codes"].items():
                if code in taken or code in failed_codes:
                    continue
                mt = plan_term(code)
                if mt is None or mt > last_completed:
                    continue
                rr = rep_row.get((major, code))
                if rr is None:
                    continue
                cat = _cat_of_row(rr)
                if _is_flex_cat(cat):
                    continue  # 文化素质/外语/学科交叉属自由选课凑学分，不点名具体课程
                key = " > ".join(cat)
                if key not in shortfall_keys and not (_parents(key) & shortfall_keys):
                    continue
                p = pop.get((major, code), 0)
                if p < 0.3:
                    continue
                cand[code] = {"n": nm, "cat": key, "pop": round(p * 100), "term": term_label(mt)}
            missing = sorted(cand.values(), key=lambda x: (-x["pop"], x["cat"]))[:40]
        # 未获学分科目（截止目前应获未获）
        # 只有文化素质/学科交叉/外语基础/外语拓展 四类以「课程组」形式展示缺口；
        # 其余类别一律展开成具体课程名（挂科未通过 / 该修未选），按缺口学分逐条消费。
        def _leaf_of(k):
            return k.split(" > ")[-1]

        flex_short = [s for s in short if _leaf_of(s["k"]) in FLEX_GROUPS]
        strict_short = [s for s in short if _leaf_of(s["k"]) not in FLEX_GROUPS]
        unearned_g = [{"n": _leaf_of(s["k"]), "k": s["k"], "req": s["req"], "got": s["got"],
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
                                  "t": m.get("term") or "", "c": "", "_ord": 1,
                                  "_tk": _term_key(m.get("term")) or (0, 0), "_cr": 0})
            cands.sort(key=lambda x: (x["_ord"], x["_tk"]))
            acc = 0.0
            for x in cands:
                unearned_c.append({k2: v for k2, v in x.items() if not k2.startswith("_")})
                acc = round(acc + x["_cr"], 1)
                if acc >= gap - 0.05:
                    break
            if acc < gap - 0.05:
                # 已完成学期内已无该修未修的课，但学分仍不足：从方案里"还没开到的课"点名
                pat = _late_hint(key)
                listed = 0
                if pat and plan.get("codes"):
                    got_names = {_norm_name(f["n"]) for f in failed}
                    have = {x["n"] for x in unearned_c}
                    for code, nm in plan["codes"].items():
                        if code in taken or code in failed_codes \
                                or _norm_name(nm) in got_names or nm in have:
                            continue
                        if not pat.search(nm):
                            continue
                        one_of = _leaf_of(key) == "四史类"  # 四史：修 1 门即可
                        mt = plan_term(code, 0.05 if one_of else 0.2)
                        if mt is None or mt > cur:
                            continue  # 开课时间未知或计划学期在未来（毕业设计等）：现在不该修，不列
                        cr = plan_credit(code)
                        unearned_c.append({"n": nm, "cat": key, "status": "待修（未选课）",
                                           "t": term_label(mt) if mt else "",
                                           "c": cr or ""})
                        listed += 1
                        acc = round(acc + cr, 1)
                        if acc >= gap - 0.05 or listed >= (1 if one_of else 4):
                            break
        unearned_g.sort(key=lambda x: -x["short"])
        # 警示
        warn = {"type": "", "level": "", "note": ""}
        failed_n = len(failed)
        th = YEAR_THRESHOLDS.get(yv, (None, None, None))
        notes = []
        # 自由课组（文化素质/外语/交叉）学分没凑够：属"应获未获"，在说明里点出缺口
        for sk in short:
            leaf = sk["k"].split(" > ")[-1]
            if leaf in FLEX_GROUPS:
                notes.append("%s差 %.1f 学分（需补选同类课程）" % (leaf, sk["req"] - sk["got"]))
        if failed_n:
            warn["type"] = "毕业警示" if yv >= 4 else "留级警示"
            warn["level"] = "挂科 %d 门" % failed_n
        if th[1] is not None and earned < th[1]:
            notes.append("累计学分 %.1f 低于 %g 分（留级线）" % (earned, th[1]))
            if not warn["type"]:
                warn["type"] = "毕业警示" if yv >= 4 else "留级警示"
        if th[2] is not None and earned < th[2]:
            notes.append("累计学分 %.1f 低于 %g 分（退学警示线）" % (earned, th[2]))
        if th[0] is not None and earned < th[0] and not warn["type"]:
            warn["type"] = "学业警示"
        if th[0] is not None and earned < th[0]:
            notes.append("累计学分 %.1f 低于 %g 分（学业警示线）" % (earned, th[0]))
        if yv >= 4 and not notes:
            notes.append("第四学年：须修满培养方案全部学分方可毕业")
        warn["note"] = "；".join(notes)
        per[sid] = {
            "name": stu.get("姓名", ""), "major": major, "class": stu.get("班级", ""),
            "year": yv, "earned": earned,
            "failed": failed, "missing": missing, "extra": extra,
            "short": short,
            "unearned": {"g": unearned_g, "c": unearned_c},
            "cats": {k: round(v, 1) for k, v in catEarned.items() if v > 0},
            "cur": current_rows, "warn": warn,
        }
    # 方案课程明细：每门课的实际归口类别 + 计划学期 + 学分 + 课程性质 + 同专业修读率
    # （供「培养方案」页展开成课程表格；方案 xlsx 是合并单元格版式，学分/学期只能由成绩表反推）
    plan_courses = {}
    for major, plan in plans.items():
        codes = plan.get("codes") or {}
        if not codes:
            continue
        pn = {_norm_name(v) for v in codes.values()}
        py = int(plan.get("year") or 0) or 0
        d = {}
        for code, nm in codes.items():
            rr = rep_row.get((major, code))
            rel = mode_rel.get((major, code))
            cat = ""
            if rr is not None:
                on = any(c in codes for c in _row_codes(rr))
                cat = " > ".join(_cat_of_row(rr, codes if on else None, pn))
                # 成绩表把「外语基础」与「外语拓展」都标成通识限选，展示时按课名拆开，
                # 只用于培养方案页分组，不动学分口径（外语两个课组本来就按合池判定）
                if cat == "通识限选 > 外语拓展课组I类" and re.match(r"^通用英语|^大学英语|^基础英语|^综合英语", nm):
                    cat = "通识必修 > 外语基础课组I类"
            d[code] = {
                "n": nm or name_of.get((major, code), ""),
                "c": round(code_credit.get((major, code), 0.0), 1),
                "t": term_label((py + rel[0], rel[1])) if (rel and py) else "",
                "k": (natureCnt[(major, code)].most_common(1)[0][0] if natureCnt.get((major, code)) else ""),
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


# 已完成学期里找不到缺口课程时，用课程名关键词定位「还没开到的课」
LATE_HINTS = (
    ("专业必修", r"毕业设计|毕业论文|毕业实习|生产实习|专业实习|专业认知实践|课程设计"),
    ("体育类-体能训练", r"体能训练|体能达标"),
    ("思政类课组", r"形势与政策|思政课社会实践"),
    ("综合拓展类", r"创新创业|技术经济|IT项目管理|HSE管理|计算机伦理|信息检索|文献检索|素质拓展"),
    ("发展规划类", r"职业生涯|就业指导|就业准备|职业能力实践"),
    ("实践类", r"劳动教育|社会调查|社会实践"),
    ("新生导学类", r"新生研讨|入学教育"),
    ("国防教育类", r"军事理论|军事技能|军训"),
    ("健康教育类", r"心理健康|健康教育"),
    ("四史类", r"四史|党史|新中国史|改革开放史|社会主义发展史"),
    ("体育俱乐部", r"体育俱乐部"),
    ("体育专项", r"体育专项"),
)


def _late_hint(key):
    leaf = key.split(" > ")[-1]
    for name, pat in LATE_HINTS:
        if leaf == name:
            return re.compile(pat)
    return None


def _parents(key):
    parts = key.split(" > ")
    return {" > ".join(parts[:i]) for i in range(1, len(parts))}


def load_academic():
    """供 build.py 汇总：返回嵌入看板的数据。"""
    plans = load_plans()
    ana = {}
    if os.path.exists(ANALYSIS_FILE):
        with open(ANALYSIS_FILE, encoding="utf-8") as f:
            ana = json.load(f)
    return {"plans": plans.get("majors", {}), "analysis": ana, "warnings": read_warnings(),
            "thresholds": {str(k): list(v) for k, v in YEAR_THRESHOLDS.items()}}
