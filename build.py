# -*- coding: utf-8 -*-
"""一键构建：data/*.csv -> dashboard.html + dashboard_脱敏外发版.html + 汇总xlsx + 脱敏CSV导出。
只读写输出目录，不碰原始文件。"""
import csv, json, os, re, datetime

from app_paths import APP_HOME as HERE, res

DATA = os.path.join(HERE, "data")
TEMPLATE = res("template.html")
_LOCAL_PHOTOS = os.path.join(HERE, "photos_src")
# 照片源目录：环境变量 > 程序目录下 photos_src（开源/换机时放这里）> 原桌面路径（本机老习惯）
PHOTO_DIR = (os.environ.get("WORKBENCH_PHOTOS")
             or (_LOCAL_PHOTOS if os.path.isdir(_LOCAL_PHOTOS) else r"C:\Users\user\Desktop\带班工作\电子照片"))
THUMB_DIR = os.path.join(HERE, "photos")

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
    "talks": "11_谈话跟进表.csv",
}


def read_csv(name):
    p = os.path.join(DATA, name)
    if not os.path.exists(p):
        return []
    with open(p, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def read_table_names():
    """返回 {key: (文件名, 表头列表)}"""
    out = {}
    for k, fn in TABLES.items():
        p = os.path.join(DATA, fn)
        if os.path.exists(p):
            with open(p, encoding="utf-8-sig") as f:
                rd = csv.reader(f)
                header = next(rd, [])
            out[k] = (fn, header)
    return out


def mask_name(n):
    n = (n or "").strip()
    return n[0] + "*" if n else ""


def mask_sid(s):
    s = (s or "").strip()
    return s[-4:] if len(s) >= 4 else s


def mask_phone(p):
    p = (p or "").strip()
    return p[:3] + "****" + p[7:] if re.match(r"^\d{11}$", p) else "***"


def mask_addr(a):
    a = (a or "").strip()
    m = re.match(r"^(.*?(?<!小)(?:区|县|街道|镇))", a)
    return m.group(1) if m else (a[:6] + "…" if a else "")


HIDE_COLS = set()  # 证件/父母电话/QQ/微信未进入主表；手机号/地址/紧急联系人按规则模糊化


def build_masked_students(students):
    out = []
    for x in students:
        y = {}
        for k, v in x.items():
            if k == "学号":
                y[k] = mask_sid(v)
            elif k == "姓名":
                y[k] = mask_name(v)
            elif k == "手机号":
                y[k] = mask_phone(v)
            elif k == "外住_地址":
                y[k] = mask_addr(v)
            elif k == "外住_紧急联系人":
                y[k] = mask_name(v)
            else:
                y[k] = v
        out.append(y)
    return out


def mask_text(text, names):
    t = text or ""
    for n in names:
        if n and n in t:
            t = t.replace(n, mask_name(n))
    return t


def render(data, out_name):
    with open(TEMPLATE, encoding="utf-8") as f:
        tpl = f.read()
    js = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    html = tpl.replace("/*__WORKBENCH_DATA__*/{}", js, 1)
    with open(os.path.join(HERE, out_name), "w", encoding="utf-8") as f:
        f.write(html)
    print("built:", out_name, f"{len(html)//1024} KB")


def write_masked_csvs(students, hours, table_names):
    if "students" not in table_names:
        print("无学生主表，跳过脱敏导出")
        return
    exp = os.path.join(HERE, "导出_脱敏版")
    os.makedirs(exp, exist_ok=True)
    # 学生主表脱敏版：隐藏敏感列
    fn, header = table_names["students"]
    keep = list(header)
    with open(os.path.join(exp, fn), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(keep)
        for x in students:
            y = build_masked_students([x])[0]
            w.writerow([y.get(c, "") for c in keep])
    # 参与学时表脱敏版
    if "ldHours" not in table_names:
        print("masked csv ->", exp)
        return
    fn2, header2 = table_names["ldHours"]
    with open(os.path.join(exp, fn2), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header2)
        for x in hours:
            w.writerow([mask_sid(x.get("学号", "")) if c == "学号" else mask_name(x.get(c, "")) if c == "姓名" else x.get(c, "") for c in header2])
    print("masked csv ->", exp)


def build_xlsx(table_names):
    try:
        import openpyxl
    except ImportError:
        print("openpyxl 未安装，跳过汇总 xlsx")
        return
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for k, (fn, _) in table_names.items():
        ws = wb.create_sheet(title=fn[:-4][:31])
        with open(os.path.join(DATA, fn), encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                ws.append(row)
    wb.save(os.path.join(HERE, "工作台数据汇总.xlsx"))
    print("built: 工作台数据汇总.xlsx")


def scan_photos():
    """扫描电子照片目录：文件名中的长数字串即学号 → {学号: 绝对路径}（只读，不复制不改动原图）"""
    idx = {}
    if not os.path.isdir(PHOTO_DIR):
        print("! 照片目录不存在:", PHOTO_DIR)
        return idx
    for root, _, files in os.walk(PHOTO_DIR):
        for fn in files:
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            import re as _re
            m = _re.findall(r"\d{9,}", fn)
            if not m:
                continue
            sid = max(m, key=len)
            p = os.path.join(root, fn)
            if sid not in idx or len(fn) < len(os.path.basename(idx[sid])):
                idx[sid] = p
    return idx


def build_photo_map(students):
    """生成 {学号: 相对路径}（供 file:// 直开时使用）+ 已有缩略图清单"""
    idx = scan_photos()
    photos, thumbs, index_json = {}, [], {}
    for x in students:
        sid = x.get("学号", "")
        # 看板内导入的照片（photos/<学号>.jpg）没有原图，也要算作已有照片
        if sid and os.path.exists(os.path.join(THUMB_DIR, sid + ".jpg")):
            thumbs.append(sid)
        p = idx.get(sid)
        if not p:
            continue
        photos[sid] = os.path.relpath(p, HERE).replace("\\", "/")
        index_json[sid] = p
        if os.path.exists(os.path.join(THUMB_DIR, sid + ".jpg")) and sid not in thumbs:
            thumbs.append(sid)
    with open(os.path.join(HERE, "photo_index.json"), "w", encoding="utf-8") as f:
        json.dump(index_json, f, ensure_ascii=False)
    print("photos:", len(photos), "matched /", len(students), "students; thumbs:", len(thumbs))
    return photos, thumbs


def main():
    table_names = read_table_names()
    data = {k: read_csv(fn) for k, fn in TABLES.items()}
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    data["generatedAt"] = now
    photos, thumbs = build_photo_map(data["students"])
    data["photos"], data["thumbs"] = photos, thumbs
    # 学业模块：培养方案 + 成绩分析 + 人工学业警示
    try:
        import academic
        data["academic"] = academic.load_academic()
    except Exception as e:
        print("! 学业数据加载失败（忽略）:", e)
        data["academic"] = {"plans": {}, "analysis": {}, "warnings": [], "thresholds": {}}
    render(data, "dashboard.html")
    # 脱敏版
    students = data["students"]
    hours = data["ldHours"]
    names = {x["姓名"] for x in students}
    masked = dict(data)
    masked["students"] = build_masked_students(students)
    # 学业数据脱敏：保留培养方案（无个人信息），清空逐生分析与警示明细
    ac = dict(data.get("academic") or {})
    ac["analysis"] = {"students": {}} if ac.get("analysis") else {}
    ac["warnings"] = []
    masked["academic"] = ac
    masked["photos"], masked["thumbs"] = {}, []  # 照片属敏感信息，脱敏版一律不带
    masked["ldHours"] = [{**h, "学号": mask_sid(h.get("学号", "")), "姓名": mask_name(h.get("姓名", ""))} for h in hours]
    masked["todos"] = [{**t, "对象": mask_text(t.get("对象", ""), names)} for t in data["todos"]]
    # 谈话记录含姓名与谈话内容：脱敏版同样模糊化（内容里出现的学生姓名一并打码）
    masked["talks"] = [
        {**t,
         "学号": mask_sid(t.get("学号", "")),
         "姓名": mask_name(t.get("姓名", "")),
         "谈话内容": mask_text(t.get("谈话内容", ""), names)}
        for t in data.get("talks", [])
    ]
    render(masked, "dashboard_脱敏外发版.html")
    write_masked_csvs(students, hours, table_names)
    build_xlsx(table_names)
    print("ALL BUILT at", now)


if __name__ == "__main__":
    main()
