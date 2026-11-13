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


def _plan_scan(data):
    """按教务导出表的视觉顺序还原「组头 → 课程行」结构。

    导出表的共享字符串按首次出现顺序排列，等价于「从上到下、从左到右」：

        (多选4)通识限选#g（要求学分：17.0…）g#/(多选2)体育类-限选#g…g#/体育专项#g…g#   ← 组头（完整层级）
          /(多选2)体育类-限选   （要求学分：3.0 已修学分：3.0）                            ← 组头行的面包屑
        课程名[课程号] / 是否通过 / 成绩 / 学分 / 计划学年学期 / 课程性质 …              ← 表头
        体育专项1[2223001005]  通过  92.0  0.5  2024-2025学年 秋季学期  必修  初修         ← 课程行
        体育专项2[2223002010]  97.0  1.0  2024-2025学年 春季学期                        ← 下一门课

    因此「组头之后、下一个组头之前的课程」就属于该组（挂到最深层子课组），
    组头行可能有多级（父 > 子 > 孙），一行里用 / 分隔的每一段都要入树。

    返回 (tree, cg, courses)：
      tree     {层级元组: 要求学分}
      cg       {课号: [课组 key]}
      courses  {课号: {n 课程名, c 学分, t 计划学年学期, k 课程性质, cat 归口课组}}
    """
    z = zipfile.ZipFile(io.BytesIO(data))
    ss = re.findall(r"<t[^>]*>([^<]*)</t>", z.read("xl/sharedStrings.xml").decode("utf-8"))
    tree, cg, courses = {}, {}, {}
    cur = None      # 当前课组 key（组头里的最深层）
    crumbs = []     # 面包屑栈 [(名, 要求学分)]
    win = None      # 当前课程行的字段窗口
    # 组头编码串自带完整层级时，面包屑碎片只是同一行的视觉重复，不必再入树
    # （碎片没有父级信息，混进去会把「父 > 子 > 孙」串成「子 > 父 > 子」）
    has_chain = any(GROUP_OPEN.search(s) for s in ss)

    def put(path):
        for i in range(1, len(path) + 1):
            k = tuple(n for n, _ in path[:i])
            r = path[i - 1][1]
            if r is None:
                tree.setdefault(k, 0.0)
            else:
                tree[k] = max(tree.get(k, 0.0), r)

    def close():
        nonlocal win
        if win is None:
            return
        d = courses.get(win["code"])
        if d is not None:
            nums = [n for n in win["nums"] if 0 < n <= 10]   # 学分列；成绩（92.0 等）>10 排除
            if nums and not d["c"]:
                d["c"] = round(min(nums), 2)
            if win["terms"] and not d["t"]:
                d["t"] = max(win["terms"], key=len)          # 多学期串取最完整的一条
            if win["nats"] and not d["k"]:
                d["k"] = win["nats"][0]
        win = None

    for raw in ss:
        t = raw.strip()
        if not t:
            continue
        # ① 组头编码串：带 #g/#r 标记与各级「要求学分」，是最权威的层级来源
        if GROUP_OPEN.search(raw):
            seg = parse_path(raw)
            if seg:
                close()
                crumbs = [[n, r] for n, r in seg]
                put(seg)
                cur = " > ".join(n for n, _ in seg)
                continue
        # ② 组头行的面包屑碎片：  /子类  +  （要求学分：X 已修学分：Y）
        #   仅在文件里没有任何组头编码串（老版导出格式）时才作为层级来源
        if not has_chain and t.startswith("/"):
            close()
            nm = _strip_multi(t.lstrip("/").strip())
            if nm:
                idx = None
                for i in range(len(crumbs) - 1, -1, -1):   # 与栈内同名层对齐，避免层层叠加
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
        # ③ 表头文字：出现在课程行之前，用来关闭上一门课的字段窗口
        if t in SCAFFOLD:
            close()
            continue
        # ④ 课程行：课程名[课号]
        m = RE_COURSE.match(t)
        if m:
            close()
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
            win = {"code": code, "nums": [], "terms": [], "nats": []}
            continue
        # ⑤ 课程行后面的字段：计划学年学期 / 课程性质 / 学分
        if win is not None:
            if RE_TERM_ANY.match(t):
                win["terms"].append(t)
                continue
            if t in NATURE_SET:
                win["nats"].append(t)
                continue
            if RE_NUM.match(t):
                win["nums"].append(float(t))
                continue
        # ⑥ 其余（通过 / 未修读 / 优秀 / 初修 / 重修 …）忽略
    close()
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
            tk = _plan_term_of((plan.get("courses") or {}).get(code, {}).get("t"))
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
        earned = round(sum(passed_best.values()), 1)
        # 类别学分累计（按归口课号：替代/认定过的课算到方案内那门课上，同一门课只计一次）
        catEarned = collections.defaultdict(float)
        extra = []
        en_credits = 0.0   # 英语学分合池（按归口课号去重，防替代课号重复计分）
        for rep, (credit, r) in hit_best.items():
            cat = _refine(r)
            if cat[0] == "学科交叉课组":
                extra.append({"n": (r.get("课程名") or "").strip(), "c": credit, "t": term_label(tk_of(r))})
                catEarned["学科交叉课组"] += credit
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
            if key.split(" > ")[-1] in EN_GROUPS:
                got = en_credits
            if got < req - 0.05:
                short_all.append({"k": key, "req": req, "got": got})
        short = [s for s in short_all if not any(o["k"].startswith(s["k"] + " > ") for o in short_all)]
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
                if not _short_hit(key, shortfall_keys):
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
                gap = next((round(s2["req"] - s2["got"], 1) for s2 in short if s2["k"] == key), None)
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
                unearned_c.append({k2: v for k2, v in x.items() if not k2.startswith("_")})
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
                notes.append("%s差 %.1f 学分（需补选同类课程）" % (leaf, sk["req"] - sk["got"]))
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
    cat = cat or ""
    if not cat:
        return False
    for k in short_keys:
        if cat == k or cat.startswith(k + " > "):
            return True
    return False


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
