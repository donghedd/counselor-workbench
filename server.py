# -*- coding: utf-8 -*-
"""本地编辑服务：启动后浏览器访问 http://127.0.0.1:8787/dashboard.html
- 仅监听 127.0.0.1，不联外网，不上传任何数据；
- 看板中点「保存更改」→ POST /api/save → 写回 data/*.csv → 自动重建看板与脱敏版。
只读写 辅导员工作台_输出/ 目录，不碰任何原始文件。"""
import base64, csv, datetime, json, os, re, subprocess, sys, urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler

from app_paths import APP_HOME as HERE
import xlsx_io
import academic

DATA = os.path.join(HERE, "data")
PHOTOS = os.path.join(HERE, "photos")
INDEX = os.path.join(HERE, "photo_index.json")
PORT = 8787

TALKS_FILE = "11_谈话跟进表.csv"
TALK_HEADER = ["记录ID", "学号", "姓名", "谈话时间", "谈话方式", "谈话内容", "更新时间"]
TABLES = {
    "students": "01_学生主表.csv",
    "ycMembers": "02_羽协会员表.csv",
    "ycActivities": "03_羽协活动表.csv",
    "ycEvents": "04_羽协赛事表.csv",
    "ycFinance": "05_羽协经费物资表.csv",
    "ldProjects": "06_劳动教育项目表.csv",
    "ldHours": "07_劳动教育参与学时表.csv",
    "ldSchedule": "08_劳动教育基地排班表.csv",
    "ldMaterials": "09_劳动教育考核材料表.csv",
    "todos": "10_跨域待办表.csv",
    "talks": TALKS_FILE,
}


def _talk_rows():
    """读取谈话记录表；文件不存在或为空时返回 (表头, [])"""
    p = os.path.join(DATA, TALKS_FILE)
    if not os.path.exists(p):
        return list(TALK_HEADER), []
    with open(p, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return list(TALK_HEADER), []
    header = rows[0]
    # 兼容表头缺失/变更：按标准列对齐
    for col in TALK_HEADER:
        if col not in header:
            header.append(col)
    norm = [(r + [""] * len(header))[: len(header)] for r in rows[1:]]
    return header, norm


def _write_talks(header, rows):
    p = os.path.join(DATA, TALKS_FILE)
    os.makedirs(DATA, exist_ok=True)
    with open(p, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def _student_name(sid):
    p = os.path.join(DATA, TABLES["students"])
    if not os.path.exists(p):
        return ""
    with open(p, encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        for r in rd:
            if (r.get("学号") or "").strip() == sid:
                return (r.get("姓名") or "").strip()
    return ""


def save_talk(obj):
    """新增（id 为空）或修改一条谈话记录，返回记录ID"""
    sid = str(obj.get("sid", "")).strip()
    rid = str(obj.get("id", "")).strip()
    when = str(obj.get("time", "")).strip().replace("T", " ")
    way = str(obj.get("way", "")).strip()
    content = str(obj.get("content", "")).strip()
    if not sid:
        raise ValueError("缺少学号")
    if not when:
        raise ValueError("请选择谈话时间")
    if not content:
        raise ValueError("请填写谈话内容")

    header, rows = _talk_rows()
    idx = {c: header.index(c) for c in TALK_HEADER}
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    def setrow(r):
        r[idx["学号"]] = sid
        r[idx["姓名"]] = _student_name(sid)
        r[idx["谈话时间"]] = when
        r[idx["谈话方式"]] = way
        r[idx["谈话内容"]] = content
        r[idx["更新时间"]] = now

    if rid:
        hit = None
        for r in rows:
            if r[idx["记录ID"]] == rid:
                hit = r
                break
        if hit is None:
            raise ValueError("未找到要修改的记录：%s" % rid)
        setrow(hit)
    else:
        nums = [int(r[idx["记录ID"]]) for r in rows if r[idx["记录ID"]].strip().isdigit()]
        rid = str(max(nums) + 1 if nums else 1)
        r = [""] * len(header)
        r[idx["记录ID"]] = rid
        setrow(r)
        rows.append(r)

    _write_talks(header, rows)
    rebuild_dashboard()
    try:
        import db_sync
        db_sync.sync_tables_from_csv(["talks"])
    except Exception:
        pass
    return rid


# ---------------- 批量导入 / 导出（Excel） ----------------
STU_FILE = TABLES["students"]
BACKUP_DIR = os.path.join(HERE, "备份")

EXCEL_ENUMS = {
    "毕业去向": ["考研", "保研", "留学", "就业", "考公考编", "自由职业", "暂无"],
    "去向确认状态": ["已确认", "待确认", "待排查"],
    "升学_结果状态": ["已确认", "待确认", "拟录取", "未通过", "待补充"],
    "升学_材料齐全": ["齐全", "缺失", "待补充"],
    "外住_审批状态": ["已通过", "已退回", "申请中", "已审批", "已到期", "已过期"],
    "外住_安全告知": ["已完成", "未完成", "待补充"],
    "推免_资格": ["有", "无", "待补充"],
    "推免_是否放弃": ["是", "否", "待补充"],
    "风险等级": ["高", "中", "低"],
}
_MATERIAL_FIELDS = ["材料_三方协议", "材料_劳动合同", "材料_录取通知", "材料_外住审批表", "材料_安全承诺书"]

FIELD_NOTES = {
    "学号": "必填，用于匹配学生；表中没有该学号则按新增处理",
    "姓名": "新增学生时必填",
    "毕业去向": "决定编辑卡显示哪些字段",
    "去向确认状态": "学生本人是否已确认去向",
    "升学_报考学校": "考研/保研填写报考或录取学校",
    "升学_结果状态": "考研/保研的录取结果",
    "就业_单位名称": "就业/考公考编填写",
    "就业_地区": "就业/考公考编填写",
    "就业_三方合同状态": "三方协议或劳动合同的签署状态",
    "具体去向": "自由职业填写具体做什么",
    "职务": "班委职务，如：班长、团支书、学习委员（兼多个职务用「、」分隔）",
    "性格": "性格标签或简述",
    "特长": "特长标签或简述",
    "外住_审批状态": "非空即视为外住学生",
    "外住_地址": "脱敏导出时只保留到区/县",
    "待办事项": "考研/保研学生不使用此字段",
    "GPA": "数字或文本均可",
    "备注": "任意补充说明",
}


def _mask_name(n):
    n = (n or "").strip()
    return (n[0] + "*") if n else ""


def _mask_phone(p):
    p = (p or "").strip()
    return p[:3] + "****" + p[7:] if re.fullmatch(r"\d{11}", p) else ("***" if p else "")


def _mask_sid(s):
    s = (s or "").strip()
    return s[-4:] if s else ""


def _mask_region(v, keep_parts=2):
    """籍贯/地址：只保留前 N 段（省-市 / 到区县街道）"""
    v = (v or "").strip()
    if not v:
        return ""
    if "-" in v:
        parts = [p for p in v.split("-") if p]
        kept = "-".join(parts[:keep_parts])
        return kept + ("***" if len(parts) > keep_parts else "")
    m = re.match(r"^(.*?(?:区|县|街道|镇))", v)
    kept = m.group(1) if m else v[:6]
    if len(kept) > 8:
        kept = kept[:8]
    return kept + ("***" if kept != v else "")


def _read_students():
    p = os.path.join(DATA, STU_FILE)
    with open(p, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    return rows[0], rows[1:]


def _export_width(field):
    if field in ("学号",):
        return 16
    if field in ("姓名", "性别", "专业", "班级", "政治面貌", "辅导员"):
        return 14
    if field in ("手机号", "寝室号", "籍贯", "毕业去向", "去向确认状态", "GPA", "专业排名"):
        return 18
    if field in ("风险标签", "待办事项", "备注", "具体去向", "外住_地址", "外住_原因"):
        return 34
    return 22


def build_export(obj):
    """按所选字段 / 范围生成填报表（xlsx），可选脱敏"""
    header, body = _read_students()
    if not header:
        raise ValueError("学生主表为空")
    fields = [f for f in (obj.get("fields") or []) if f in header]
    if not fields:
        raise ValueError("没有选择任何字段")
    # 学号/姓名固定在最前，便于匹配
    order = [f for f in ["学号", "姓名"] if f in fields] + [f for f in fields if f not in ("学号", "姓名")]
    masked = bool(obj.get("masked"))
    keep_sid = obj.get("keepSid", True)
    blank = bool(obj.get("blank"))

    sids = obj.get("sids")
    rows_src = body
    if sids:
        want = set(str(s).strip() for s in sids)
        rows_src = [r for r in body if str(r[header.index("学号")]).strip() in want]

    idx = {h: i for i, h in enumerate(header)}
    names = [r[idx["姓名"]].strip() for r in rows_src if r[idx["姓名"]].strip()]

    def val(row, f):
        v = row[idx[f]] if f in idx else ""
        v = "" if v is None else str(v).strip()
        if not masked or not v:
            return v
        if f == "姓名":
            return _mask_name(v)
        if f == "手机号":
            return _mask_phone(v)
        if f == "学号":
            return v if keep_sid else _mask_sid(v)
        if f == "籍贯":
            return _mask_region(v, 2)
        if f == "外住_地址":
            return _mask_region(v, 3)
        if f == "外住_紧急联系人":
            return _mask_name(v)
        if f in ("备注", "待办事项"):
            t = v
            for n in names:
                if n and n in t:
                    t = t.replace(n, _mask_name(n))
            return t
        return v

    out_rows = [] if blank else [[val(r, f) for f in order] for r in rows_src]

    # 说明页
    notes = []
    for f in order:
        enums = EXCEL_ENUMS.get(f) or (["齐全", "缺失", "待补充", "不适用"] if f in _MATERIAL_FIELDS else [])
        notes.append([f, "必填" if f in ("学号",) else "选填", FIELD_NOTES.get(f, "留空表示不修改该字段"), " / ".join(enums)])

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    if blank:
        fname = "学生信息填报表_空白模板_%s.xlsx" % ts
    else:
        tag = "脱敏_" if masked else ""
        fname = "学生信息填报表_%s%s.xlsx" % (tag, ts)
    sheets = [
        {"name": "学生信息", "header": order, "rows": out_rows, "widths": [_export_width(f) for f in order]},
        {"name": "填写说明", "header": ["字段", "必填", "说明", "可选值"], "rows": notes, "widths": [22, 8, 46, 40]},
    ]
    return xlsx_io.write_xlsx(sheets), fname, len(out_rows)


def _match_col(name, header):
    name = (name or "").strip()
    if name in header:
        return name
    base = re.split(r"[（(]", name)[0].strip()
    cand = [h for h in header if re.split(r"[（(]", h)[0].strip() == base]
    return cand[0] if len(cand) == 1 else None


def apply_import(fname, data):
    """把上传的表格合并进学生主表：有学号且存在→只更新填了值的列；学号不存在→新增一行"""
    up_header, up_rows = xlsx_io.read_table(fname, data)
    header, body = _read_students()
    idx = {h: i for i, h in enumerate(header)}

    colmap, unknown = {}, []
    for h in up_header:
        tgt = _match_col(h, header)
        if tgt:
            colmap[h] = tgt
        else:
            unknown.append(h)
    if "学号" not in colmap.values():
        raise ValueError("表格中缺少「学号」列——请先下载模板再填写")
    sid_key = [k for k, v in colmap.items() if v == "学号"][0]
    name_key = next((k for k, v in colmap.items() if v == "姓名"), None)

    by_sid = {}
    for i, r in enumerate(body):
        s = str(r[idx["学号"]]).strip() if len(r) > idx["学号"] else ""
        if s:
            by_sid.setdefault(s, i)

    updated = added = 0
    skipped, warnings = [], []
    for line, row in enumerate(up_rows, start=2):
        sid = str(row.get(sid_key, "")).strip()
        nm = str(row.get(name_key, "")).strip() if name_key else ""
        cells = {}
        for k, col in colmap.items():
            v = str(row.get(k, "") or "").strip()
            if v and col != "学号":
                cells[col] = v
        if not cells and not nm:
            continue
        # 枚举校验
        for f, v in list(cells.items()):
            allow = EXCEL_ENUMS.get(f)
            if allow and v not in allow:
                warnings.append("第%d行 %s：%s 不在可选值（%s）内，已忽略该单元格" % (line, f, v, "/".join(allow)))
                del cells[f]
        if sid and sid in by_sid:
            r = body[by_sid[sid]]
            while len(r) < len(header):
                r.append("")
            if nm and nm != r[idx["姓名"]]:
                warnings.append("第%d行：学号 %s 在表中是「%s」，上传写的是「%s」，按学号匹配处理"
                                % (line, sid, r[idx["姓名"]], nm))
            diff = {f: v for f, v in cells.items() if r[idx[f]] != v}
            for f, v in diff.items():
                r[idx[f]] = v
            if diff:
                updated += 1
            else:
                skipped.append("第%d行 %s：与现有数据一致，无需更新" % (line, sid))
        else:
            if not nm:
                skipped.append("第%d行：学号「%s」不存在且没有姓名，已跳过" % (line, sid or "空"))
                continue
            if not sid:
                skipped.append("第%d行 %s：缺少学号，无法新增" % (line, nm))
                continue
            r = [""] * len(header)
            r[idx["学号"]] = sid
            r[idx["姓名"]] = nm
            for f, v in cells.items():
                r[idx[f]] = v
            body.append(r)
            by_sid[sid] = len(body) - 1
            added += 1

    result = {"ok": True, "updated": updated, "added": added,
              "skipped": skipped[:12], "skippedCount": len(skipped),
              "warnings": warnings[:12], "unknown": unknown[:12], "rebuilt": False}
    if not updated and not added:
        return result

    # 备份 → 写回 → 重建看板 → 同步数据库
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    import shutil
    shutil.copy2(os.path.join(DATA, STU_FILE), os.path.join(BACKUP_DIR, "01_学生主表_导入前_%s.csv" % ts))
    with open(os.path.join(DATA, STU_FILE), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(body)
    rebuild_dashboard()
    try:
        import db_sync
        db_sync.sync_tables_from_csv(["students"])
    except Exception:
        pass
    result["rebuilt"] = True
    return result


def _backup_file(path):
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(BACKUP_DIR, os.path.basename(path).replace(".csv", "_%s.csv" % ts))
    import shutil
    shutil.copy2(path, dst)
    return dst


def _write_table(fn, header, rows):
    with open(os.path.join(DATA, fn), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def add_student(obj):
    """新增一名学生：学号必填且不能重复，其余字段留空即可"""
    rec = obj.get("record") or {}
    header, body = _read_students()
    if not header:
        raise ValueError("学生主表为空")
    idx = {h: i for i, h in enumerate(header)}
    sid = str(rec.get("学号", "")).strip()
    name = str(rec.get("姓名", "")).strip()
    if not sid or not re.fullmatch(r"\d{6,}", sid):
        raise ValueError("学号必须为 6 位以上数字")
    if not name:
        raise ValueError("姓名必填")
    if any(r[idx["学号"]].strip() == sid for r in body):
        raise ValueError("学号已存在：%s" % sid)
    row = [""] * len(header)
    for k, v in rec.items():
        if k in idx:
            row[idx[k]] = str(v).strip()
    row[idx["学号"]] = sid
    row[idx["姓名"]] = name
    if "最后更新时间" in idx:
        row[idx["最后更新时间"]] = datetime.datetime.now().strftime("%Y-%m-%d")
    body.append(row)
    _backup_file(os.path.join(DATA, STU_FILE))
    _write_table(STU_FILE, header, body)
    rebuild_dashboard()
    try:
        import db_sync
        db_sync.sync_tables_from_csv(["students"])
    except Exception:
        pass
    return {"ok": True, "sid": sid, "name": name}


def delete_students(obj):
    """按学号删除一名或多名学生，同时移除其谈话记录"""
    sids = obj.get("sids") or ([obj.get("sid")] if obj.get("sid") else [])
    sids = [str(s).strip() for s in sids if str(s).strip()]
    if not sids:
        raise ValueError("缺少要删除的学号")
    want = set(sids)

    header, body = _read_students()
    idx = {h: i for i, h in enumerate(header)}
    kept, removed = [], []
    for r in body:
        (removed if r[idx["学号"]].strip() in want else kept).append(r)
    found = {r[idx["学号"]].strip() for r in removed}
    not_found = [s for s in sids if s not in found]
    if not removed:
        raise ValueError("未找到要删除的学生：%s" % "、".join(sids))

    _backup_file(os.path.join(DATA, STU_FILE))
    _write_table(STU_FILE, header, kept)

    # 谈话记录：删除对应学号的记录
    talk_fn = TABLES.get("talks")
    if talk_fn and os.path.exists(os.path.join(DATA, talk_fn)):
        p = os.path.join(DATA, talk_fn)
        with open(p, encoding="utf-8-sig", newline="") as f:
            trows = list(csv.reader(f))
        if len(trows) > 1 and "学号" in trows[0]:
            ti = trows[0].index("学号")
            kept_t = [trows[0]] + [r for r in trows[1:] if r[ti].strip() not in want]
            _backup_file(p)
            _write_table(talk_fn, trows[0], kept_t[1:])
            talk_touched = True
        else:
            talk_touched = False
    else:
        talk_touched = False

    rebuild_dashboard()
    try:
        import db_sync
        keys = ["students"] + (["talks"] if talk_touched else [])
        db_sync.sync_tables_from_csv(keys)
    except Exception:
        pass
    return {"ok": True, "removed": len(removed), "notFound": not_found}


# ---------- 学业模块（培养方案 / 全校成绩 / 学业警示） ----------
GRADES_SRC = os.path.join(DATA, "13_成绩源文件.xlsx")


def _reanalyze():
    """用最近一次上传的成绩源文件重跑分析（培养方案调整后调用）。"""
    if not os.path.exists(GRADES_SRC):
        return None
    with open(GRADES_SRC, "rb") as f:
        data = f.read()
    name = None
    try:
        with open(academic.ANALYSIS_FILE, encoding="utf-8") as f:
            name = json.load(f).get("source")
    except Exception:
        pass
    return academic.analyze_grades(name or "全校成绩查询.xlsx", data)


def academic_plan_import(obj):
    data = base64.b64decode(obj.get("data") or "")
    info = academic.save_plan(obj.get("name") or "培养方案.xlsx", data)
    ana = _reanalyze()
    rebuild_dashboard()
    return {"ok": True, "plan": info,
            "analyzed": bool(ana), "students": len(ana["students"]) if ana else 0}


def academic_grades_import(obj):
    data = base64.b64decode(obj.get("data") or "")
    name = obj.get("name") or "全校成绩查询.xlsx"
    os.makedirs(DATA, exist_ok=True)
    with open(GRADES_SRC, "wb") as f:
        f.write(data)
    ana = academic.analyze_grades(name, data)
    rebuild_dashboard()
    warnN = sum(1 for p in ana["students"].values() if p["warn"]["type"])
    failN = sum(1 for p in ana["students"].values()
                if any(not f.get("flex") for f in p["failed"]))
    missN = sum(1 for p in ana["students"].values() if p["missing"])
    return {"ok": True, "students": len(ana["students"]), "currentTerm": ana["currentTerm"],
            "failedStudents": failN, "missingStudents": missN, "warnStudents": warnN}


def academic_plan_req(obj):
    academic.set_plan_req(str(obj.get("major") or ""), str(obj.get("key") or ""), obj.get("req"))
    ana = _reanalyze()
    rebuild_dashboard()
    return {"ok": True, "analyzed": bool(ana)}


def academic_warning(obj):
    res = academic.add_warning(obj)
    rebuild_dashboard()
    try:
        import db_sync
        db_sync.sync_tables_from_csv(["warnings"])
    except Exception:
        pass
    return {"ok": True, "result": res}


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=HERE, **kw)

    def log_message(self, fmt, *args):
        sys.stdout.write("[访问] %s\n" % (fmt % args))

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urllib.parse.unquote(self.path.split("?")[0])
        if path == "/api/thumbs/plan":
            missing = []
            if os.path.exists(INDEX):
                with open(INDEX, encoding="utf-8") as f:
                    idx = json.load(f)
                for sid in idx:
                    if not os.path.exists(os.path.join(PHOTOS, sid + ".jpg")):
                        missing.append(sid)
            self._json({"missing": missing})
            return
        if path.startswith("/photo/"):
            sid = path[len("/photo/"):]
            try:
                with open(INDEX, encoding="utf-8") as f:
                    idx = json.load(f)
                p = idx.get(sid)
                if not p or not os.path.exists(p):
                    self.send_error(404, "photo not found")
                    return
                with open(p, "rb") as f:
                    body = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "max-age=3600")
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_error(500, str(e))
            return
        super().do_GET()

    def do_POST(self):
        if self.path == "/api/rebuild":
            try:
                rebuild_dashboard()
                self._json({"ok": True})
            except Exception as e:
                self._json({"ok": False, "error": str(e)[-300:]}, 500)
            return
        if self.path == "/api/thumb":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                sid = str(obj["sid"])
                data = obj["data"]
                if "," in data:
                    data = data.split(",", 1)[1]
                os.makedirs(PHOTOS, exist_ok=True)
                with open(os.path.join(PHOTOS, sid + ".jpg"), "wb") as f:
                    f.write(base64.b64decode(data))
                self._json({"ok": True})
            except Exception as e:
                self._json({"ok": False, "error": str(e)}, 500)
            return
        if self.path == "/api/talk":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                rid = save_talk(obj)
                self._json({"ok": True, "id": rid})
            except Exception as e:
                self._json({"ok": False, "error": str(e)[-300:]}, 500)
            return
        if self.path == "/api/student/add":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                self._json(add_student(obj))
            except Exception as e:
                self._json({"ok": False, "error": str(e)[-300:]}, 500)
            return
        if self.path == "/api/student/delete":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                self._json(delete_students(obj))
            except Exception as e:
                self._json({"ok": False, "error": str(e)[-300:]}, 500)
            return
        if self.path == "/api/excel/export":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                data, fname, n = build_export(obj)
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Content-Disposition", "attachment; filename*=UTF-8''%s" % urllib.parse.quote(fname))
                self.send_header("X-Row-Count", str(n))
                self.end_headers()
                self.wfile.write(data)
            except Exception as e:
                self._json({"ok": False, "error": str(e)}, 500)
            return
        if self.path == "/api/excel/import":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                data = base64.b64decode(obj.get("data") or "")
                res = apply_import(obj.get("name") or "upload.xlsx", data)
                self._json(res)
            except Exception as e:
                self._json({"ok": False, "error": str(e)[-300:]}, 500)
            return
        if self.path == "/api/academic/plan":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                self._json(academic_plan_import(obj))
            except Exception as e:
                self._json({"ok": False, "error": str(e)[-300:]}, 500)
            return
        if self.path == "/api/academic/grades":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                self._json(academic_grades_import(obj))
            except Exception as e:
                self._json({"ok": False, "error": str(e)[-300:]}, 500)
            return
        if self.path == "/api/academic/plan/req":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                self._json(academic_plan_req(obj))
            except Exception as e:
                self._json({"ok": False, "error": str(e)[-300:]}, 500)
            return
        if self.path == "/api/academic/warning":
            try:
                ln = int(self.headers.get("Content-Length", 0))
                obj = json.loads(self.rfile.read(ln).decode("utf-8"))
                self._json(academic_warning(obj))
            except Exception as e:
                self._json({"ok": False, "error": str(e)[-300:]}, 500)
            return
        if self.path != "/api/save":
            self._json({"ok": False, "error": "not found"}, 404)
            return
        try:
            ln = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(ln).decode("utf-8"))
            changed = 0
            for t, items in payload.items():
                if t not in TABLES:
                    raise ValueError("未知数据表: %s" % t)
                p = os.path.join(DATA, TABLES[t])
                with open(p, encoding="utf-8-sig", newline="") as f:
                    rows = list(csv.reader(f))
                if not rows:
                    continue
                header = rows[0]
                # 列名容错：CSV 表头可能带枚举说明，如「状态(未开始/进行中/...)」
                def find_col(col):
                    if col in header:
                        return header.index(col)
                    cand = [i for i, h in enumerate(header) if h.split("(")[0].split("（")[0].strip() == col]
                    if len(cand) == 1:
                        return cand[0]
                    raise ValueError("未知列: %s（表 %s）" % (col, t))
                for it in items:
                    ri = int(it["row"]) + 1  # +1 跳过表头
                    if ri < 1 or ri >= len(rows):
                        raise ValueError("行号越界: %s 第%d行" % (t, ri))
                    for col, val in it["cells"].items():
                        rows[ri][find_col(col)] = str(val)
                        changed += 1
                with open(p, "w", encoding="utf-8-sig", newline="") as f:
                    csv.writer(f).writerows(rows)
            # 重建看板与脱敏版
            rebuild_dashboard()
            # 被修改的表同步回 SQLite（只覆盖这几张表，保留数据库里其他表的人工改动）
            try:
                sys.path.insert(0, HERE)
                import db_sync
                db_sync.sync_tables_from_csv(list(payload.keys()))
            except Exception:
                pass
            self._json({"ok": True, "changed": changed})
        except Exception as e:
            self._json({"ok": False, "error": str(e)}, 500)


def rebuild_dashboard():
    """重建看板：直接调用 build.main()。

    打包成 exe 后 sys.executable 指向 exe 本身，不能再起 python 子进程，故改为模块内调用。
    """
    import importlib
    if "build" in sys.modules:
        importlib.reload(sys.modules["build"])
    import build as build_mod
    build_mod.main()


def open_browser(url):
    """Windows 下用默认浏览器打开；os.startfile 比 webbrowser 更可靠。"""
    try:
        os.startfile(url)
        return True
    except Exception:
        pass
    try:
        import webbrowser
        return bool(webbrowser.open(url))
    except Exception:
        return False


def port_in_use(port):
    import socket
    s = socket.socket()
    s.settimeout(0.6)
    try:
        return s.connect_ex(("127.0.0.1", port)) == 0
    finally:
        s.close()


def main():
    url = "http://127.0.0.1:%d/dashboard.html" % PORT

    if port_in_use(PORT):
        # 已经开着一个实例：直接复用，不再重复启动
        print("检测到服务已在运行，直接打开页面：%s" % url)
        open_browser(url)
        print("本窗口可以直接关闭（不会影响已在运行的服务）。")
        return

    print("=" * 58)
    print("  辅导员工作台 - 编辑模式")
    print("=" * 58)
    print("  服务地址：%s" % url)
    print("  仅本机可访问，不联网、不上传任何数据")
    print("  关闭本窗口即停止服务")
    print("=" * 58)

    import threading, time

    def _open():
        time.sleep(1.0)
        if not open_browser(url):
            print("\n[提示] 浏览器没有自动打开，请手动复制上面的地址到浏览器打开")
        else:
            print("\n已在浏览器中打开工作台。")

    threading.Thread(target=_open, daemon=True).start()
    try:
        HTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止。")
    except OSError as e:
        print("\n[错误] 端口 %d 无法监听：%s" % (PORT, e))


if __name__ == "__main__":
    main()

