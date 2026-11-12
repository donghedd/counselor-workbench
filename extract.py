# -*- coding: utf-8 -*-
"""一次性抽取脚本：从原始工作簿(只读)生成 辅导员工作台_输出/data/ 下的 10 张标准 CSV。
原始文件只读不写。后续维护直接改 CSV。"""
import csv, os, re, datetime as dt
import openpyxl, xlrd

BASE = r"C:\Users\user\Desktop"
OUT = os.path.join(BASE, "23", "辅导员工作台_输出")
DATA = os.path.join(OUT, "data")
BAK = os.path.join(OUT, "备份", "首次导入快照")
os.makedirs(DATA, exist_ok=True)
os.makedirs(BAK, exist_ok=True)

P_ROSTER = os.path.join(BASE, "23", "23级软工、网安学生信息.xlsx")
P_BAOKAO = os.path.join(BASE, "23", "报考确认.xlsx")
P_PLAN = os.path.join(BASE, "23", "职业规划问卷.xlsx")
P_GPA_RG = os.path.join(BASE, "带班工作", "学习", "软工GPA.xlsx")
P_GPA_WA = os.path.join(BASE, "带班工作", "学习", "网安GPA.xlsx")
P_YC_BUDGET = os.path.join(BASE, "10.社团指导（羽协）", "羽协2023上半学期预算表.xls")
P_LD_SUM = os.path.join(BASE, "劳动教育", "2026年秋季学期", "加分", "已加", "汇总.xls")
P_LD_W7 = os.path.join(OUT, "_临时解析副本", "第七周活动二课.xlsx")
P_LD_W8 = os.path.join(OUT, "_临时解析副本", "第八周二课加分.xlsx")
P_LD_W10 = os.path.join(OUT, "_临时解析副本", "第十周二课.xlsx")
P_LD_DORM = os.path.join(OUT, "_临时解析副本", "计科院寝室打扫卫生日常收集表（合并）.xlsx")

TODAY = "2026-11-11"
PEND = "待补充"


def s(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d") if isinstance(v, dt.date) and not isinstance(v, dt.datetime) else v.strftime("%Y-%m-%d %H:%M:%S")
    return str(v).strip()


def norm_h(h):
    if not h:
        return ""
    h = re.sub(r"[（(][^）)]*[）)]", "", h.split("\n")[0])
    return h.strip()


def read_sheet_rows(path, sheet_name=None, header_row=0):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]
    rows = []
    for r in ws.iter_rows(values_only=True):
        rows.append([s(c) for c in r])
    wb.close()
    hdr = header_row
    header = [norm_h(h) for h in rows[hdr]]
    out = []
    for r in rows[hdr + 1:]:
        if not any(x for x in r):
            continue
        d = {}
        for i, h in enumerate(header):
            if h:
                d[h] = r[i] if i < len(r) else ""
        out.append(d)
    return out, rows[hdr]


def read_xls_rows(path, sheet_name, header_row=0):
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_name(sheet_name)
    rows = [[s(ws.cell_value(r, c)) for c in range(ws.ncols)] for r in range(ws.nrows)]
    header = [norm_h(h) for h in rows[header_row]]
    out = []
    for r in rows[header_row + 1:]:
        if not any(x for x in r):
            continue
        out.append({header[i]: r[i] if i < len(header) else "" for i in range(len(header)) if header[i]})
    return out


def write_csv(name, header, rows):
    p = os.path.join(DATA, name)
    with open(p, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"{name}: {len(rows)} rows")
    return p


# ---------- 备份原始文件 ----------
import shutil
for src in [P_ROSTER, P_BAOKAO, P_LD_SUM, P_YC_BUDGET]:
    shutil.copy2(src, os.path.join(BAK, os.path.basename(src)))
print("backup done:", BAK)

# ---------- 1. 学生主表 ----------
roster, _ = read_sheet_rows(P_ROSTER, "学生信息", 0)
yibiao, _ = read_sheet_rows(P_ROSTER, "一标三实", 0)
waizhu, _ = read_sheet_rows(P_ROSTER, "外住学生名单", 1)
baokao, _ = read_sheet_rows(P_BAOKAO, None, 1)
gpa = {}
for p in [P_GPA_RG, P_GPA_WA]:
    for r in read_sheet_rows(p, "已修全部课程GPA", 0)[0]:
        gpa[r["学号"]] = (r.get("GPA", ""), r.get("排名", ""))

yibiao_map = {r["学号"]: r for r in yibiao}
waizhu_map = {r["学号"]: r for r in waizhu}
baokao_map = {r["学号"]: r for r in baokao}

# ---------- 职业规划问卷：毕业首选规划 → 毕业意向 ----------
PLAN_MAP_RAW = {"B.直接就业": "直接就业", "D.考公考编": "考公考编",
                "A.国内考研或推免": "考研或推免", "推免": "推免",
                "E.创业或自由职业": "创业或自由职业"}
plan_rows, _ = read_sheet_rows(P_PLAN, None, 0)
name_count = {}
for r in roster:
    name_count[r["姓名"]] = name_count.get(r["姓名"], 0) + 1
plan_by_key = {}
for r in plan_rows:
    nm = r.get("1.姓名", "").strip()
    if not nm:
        continue
    intent = PLAN_MAP_RAW.get(r.get("3.毕业首选规划", "").strip(), r.get("3.毕业首选规划", "").strip())
    key_full = (nm, r.get("2.班级", "").strip())
    plan_by_key[key_full] = intent
    # 姓名在花名册唯一时允许仅按姓名匹配（班级口径差异兜底）
    if name_count.get(nm, 0) == 1:
        plan_by_key.setdefault((nm, "__ANY__"), intent)

matched_plan = 0
plan_unmatched = []

MAIN_HDR = ["学号", "姓名", "性别", "专业", "班级", "手机号", "辅导员", "寝室号", "籍贯",
            "毕业去向", "毕业意向(问卷)", "去向确认状态",
            "升学_报考学校", "升学_结果状态", "升学_复试面试时间", "升学_材料齐全",
            "就业_单位名称", "就业_岗位", "就业_地区", "就业_签约日期", "就业_三方合同状态",
            "外住_类型", "外住_原因", "外住_起止日期", "外住_审批状态", "外住_地址", "外住_紧急联系人", "外住_安全告知",
            "推免_资格", "推免_排名", "推免_公示状态", "推免_录取学校", "推免_是否放弃", "推免_材料状态",
            "GPA", "专业排名",
            "风险标签", "待办事项", "下次跟进日期", "最后更新时间",
            "材料_三方协议", "材料_劳动合同", "材料_录取通知", "材料_外住审批表", "材料_安全承诺书",
            "政治面貌", "备注"]
main_rows = []
risk_rows = []
gender_conflict = []
for r in roster:
    sid = r["学号"]
    name = r["姓名"]
    big, sub = "待定", "待定"
    bk = baokao_map.get(sid)
    if bk:
        if bk["升学类别"] == "国内推免":
            big, sub = "推免", "保研"
        elif bk["升学类别"] == "国内考研":
            big, sub = "升学", "考研"
    wz = waizhu_map.get(sid)
    if big == "待定" and wz:
        big, sub = "外住", wz.get("外宿类型", "")
    # 毕业意向（问卷）与去向确认状态
    intent = plan_by_key.get((name, r["班级"])) or plan_by_key.get((name, "__ANY__")) or ""
    if intent:
        matched_plan += 1
    else:
        plan_unmatched.append(name)
    # 单一「毕业去向」字段（考研/保研/留学/就业/考公考编/自由职业/暂无）：已确认的升学/就业优先于问卷意向
    dest = ("保研" if big == "推免" else
            "留学" if ("出国" in sub or "境外" in sub or "留学" in sub) else
            "考研" if big == "升学" else
            "就业" if big == "就业" else
            {"直接就业": "就业", "考公考编": "考公考编", "创业或自由职业": "自由职业",
             "考研或推免": "考研", "推免": "考研"}.get(intent, "暂无"))
    if big == "推免":
        confirm = "已确认"          # 保研同学去向已全部确认
    elif bk:
        confirm = "待确认"          # 考研已报名，结果未出
    elif intent:
        confirm = "待确认"          # 有意向（就业/考公等），最终去向未定
    else:
        confirm = "待排查"
    # 风险
    risks, todos = [], []
    phone = r["联系电话"]
    if len(phone) != 11 or not phone.isdigit():
        risks.append("手机号位数异常")
        todos.append("核对更正手机号")
    if r["专业"] == "降级":
        risks.append("专业待确认(降级)")
        todos.append("确认降级后班级归属")
    yb = yibiao_map.get(sid)
    if yb and yb["性别"] != r["性别"]:
        risks.append("性别两表不一致")
        todos.append("核对学生信息/一标三实性别")
        gender_conflict.append(sid)
    if wz and wz["审核状态"] == "已退回":
        risks.append("外住审批已退回")
        todos.append("跟进外住审批退回原因")
    if bk and big == "推免":
        todos.append("补充推免录取/公示/放弃信息")
    if big == "待定":
        todos.append("毕业去向排查确认")
    if wz:
        todos.append("补充外住起止日期与安全告知")
    if not risks:
        risks.append("")
    main_rows.append([
        sid, name, r["性别"], r["专业"], r["班级"], phone, "辅导员", r["寝室号"], r["籍贯"],
        dest, intent, confirm,
        bk["报考学校全称"] if bk else "", bk["是否已报名确认"] if bk else "",
        PEND, PEND,
        PEND, PEND, PEND, PEND, PEND,
        wz["外宿类型"] if wz else "", wz["外宿原因"] if wz else "",
        wz["外住时间"] if wz else "", wz["审核状态"] if wz else "",
        wz["详细地址"] if wz else "", PEND, PEND,
        "有" if big == "推免" else "待补充" if bk else "不适用",
        PEND, PEND, PEND, PEND, PEND,
        gpa.get(sid, ("", ""))[0], gpa.get(sid, ("", ""))[1],
        "；".join(x for x in risks if x), "；".join(x for x in todos if x),
        PEND if todos else "", TODAY,
        PEND, PEND, bk["报考学校全称"] if bk and big != "推免" else PEND,
        "齐全" if wz and wz["审核状态"] == "已通过" else "缺失" if wz else "不适用",
        PEND, r["政治面貌"], "两版报考确认一致；来源:学生信息表" if not risks else ""
    ])
write_csv("01_学生主表.csv", MAIN_HDR, main_rows)

# ---------- 2~5 羽协四表 ----------
write_csv("02_羽协会员表.csv",
          ["姓名", "学号", "学院/班级", "角色(会长/部长/干事/会员)", "联系方式", "入会日期", "会员状态(活跃/不活跃/已退会)", "备注"], [])
write_csv("03_羽协活动表.csv",
          ["活动名称", "类型(训练/例会/招新/友谊赛/赛事/宣传)", "日期", "时间", "地点", "负责人", "报名人数", "签到人数",
           "场地审批状态", "安全告知状态", "保险状态", "经费预算", "经费支出", "活动状态(筹备/已审批/进行中/已完成/已取消)", "总结归档状态", "备注"], [])
write_csv("04_羽协赛事表.csv",
          ["赛事名称", "时间", "地点", "参赛队伍/人员", "报名状态", "保险状态", "交通安排", "安全材料", "成绩", "总结归档状态", "备注"], [])

budget = read_xls_rows(P_YC_BUDGET, "Sheet1", 1)
fin_rows = []
cur_act = ""
for r in budget:
    act = r.get("活动", "")
    if act and act != "合计":
        cur_act = act
    typ = r.get("类型", "")
    if not typ or typ in ("合计",) or not re.match(r"^\d", r.get("合计", "")):
        continue
    fin_rows.append([f"{cur_act}-{typ}" if cur_act else typ, "预算(2023上半年历史参考)",
                     r.get("合计", ""), PEND, PEND, r.get("单价", ""), r.get("数量", ""),
                     PEND, PEND, PEND, PEND, PEND, "来源:羽协2023上半学期预算表"])
write_csv("05_羽协经费物资表.csv",
          ["项目", "类别(预算/采购/借用)", "预算金额", "支出", "余额", "单价", "数量",
           "采购/借用", "归还状态", "报销状态", "负责人", "截止日期", "备注"], fin_rows)

# ---------- 6~9 劳动教育四表 ----------
proj_hdr = ["项目名称", "类型(课程劳动/实践劳动/志愿服务/基地劳动/校园服务)", "学期", "负责教师",
            "基地/地点", "起止日期", "学时", "考核方式", "项目状态(筹备/进行中/已完成)", "参与人数(2023级)", "备注"]
projects = [
    ["旗帜活动(每周教室/公共区域打扫)", "日常劳动", "2026年秋季学期", PEND, PEND, PEND, PEND, PEND, "进行中", 0, "来源:第七/十周二课名单"],
    ["网信活动(教室/机房清扫)", "日常劳动", "2026年秋季学期", PEND, PEND, PEND, PEND, PEND, "进行中", 0, "来源:第十四/十七周及本学期名单"],
    ["菜园子翻耕/浇水(青年劳动实践基地)", "基地劳动", "2026年秋季学期", PEND, "青年劳动实践基地", PEND, PEND, PEND, "进行中", 0, "来源:菜园子二课名单及照片"],
    ["寝室卫生打扫(日常)", "日常劳动", "2026年秋季学期", PEND, "学生寝室", PEND, PEND, PEND, "进行中", 0, "来源:寝室打扫日常收集表"],
    ["权益提案活动", "校园服务", "2026年秋季学期", PEND, PEND, PEND, PEND, PEND, "已完成", 0, "来源:第八周二课名单(全院)"],
    ["迎新志愿服务", "志愿服务", "2026年秋季学期", PEND, PEND, PEND, PEND, PEND, "已完成", 0, "来源:已加/2026年迎新志愿者二课"],
    ["除草志愿服务", "志愿服务", "2026年秋季学期", PEND, PEND, PEND, PEND, PEND, "已完成", 0, "来源:已加/9.10除草志愿服务二课"],
    ["学子林劳动", "志愿服务", "2026年秋季学期", PEND, PEND, PEND, PEND, PEND, "已完成", 0, "来源:已加/学子林二课"],
    ["开学典礼志愿服务", "志愿服务", "2026年秋季学期", PEND, PEND, PEND, PEND, PEND, "已完成", 0, "来源:已加/开学典礼二课"],
    ["中秋月饼制作活动", "校园服务", "2026年秋季学期", PEND, PEND, "2026-09(策划书)", PEND, PEND, "已完成", 0, "来源:策划书及新闻稿"],
    ["暑期社会实践", "实践劳动", "2026年暑期", PEND, PEND, PEND, PEND, PEND, "已完成", 0, "来源:暑期社会实践二课加分名单"],
    ["劳动理论课(01-08章节)", "课程劳动", "2024-2025年春季学期", "辅导员等9人", PEND, "2024-03~2025-05", PEND, "理论考核", "已完成", 0, "来源:理论课授课情况统计表21条"],
]
write_csv("06_劳动教育项目表.csv", proj_hdr, projects)

# 参与学时表：收集 roster 学号集合
roster_ids = {r["学号"] for r in roster}
name_by_id = {r["学号"]: r["姓名"] for r in roster}
cls_by_id = {r["学号"]: r["班级"] for r in roster}

hours_hdr = ["学号", "姓名", "班级", "学期", "项目名称", "活动类型", "参与日期", "二课加分",
             "学时", "认定状态(待审核/已认定/不通过)", "材料状态", "审核人", "备注"]
hours_rows = []


def add_hours(rows, proj, typ, term, status, note, date_col=None):
    n = 0
    for r in rows:
        sid = r.get("学号", "")
        if sid not in roster_ids:
            continue
        hours_rows.append([sid, name_by_id.get(sid, r.get("姓名", "")), cls_by_id.get(sid, ""),
                           term, proj, typ, r.get(date_col, "") if date_col else "", r.get("二课加分") or r.get("加分") or r.get("得分") or "",
                           PEND, status, PEND, PEND, note])
        n += 1
    return n


n_sum = 0
for sh, proj, typ in [("日常", "日常劳动(汇总-日常)", "日常劳动"), ("公益", "公益活动(汇总-公益)", "公益服务"),
                      ("生产顶岗", "生产顶岗(汇总)", "生产顶岗"), ("理论", "劳动理论课", "课程劳动")]:
    rows = read_xls_rows(P_LD_SUM, sh, 0)
    n_sum += add_hours(rows, proj, typ, "2024-2025春季起累计", "已认定", "来源:已加/汇总.xls")
n_w7 = add_hours(read_sheet_rows(P_LD_W7, None, 0)[0], "第七周旗帜活动", "公益服务", "2026年秋季学期", "待审核", "来源:第七周活动二课")
n_w8 = add_hours(read_sheet_rows(P_LD_W8, None, 0)[0], "权益提案活动", "公益服务", "2026年秋季学期", "待审核", "来源:第八周二课加分")
n_w10 = add_hours(read_sheet_rows(P_LD_W10, None, 0)[0], "第十周旗帜活动", "日常劳动", "2026年秋季学期", "待审核", "来源:第十周二课")
n_dm = add_hours(read_sheet_rows(P_LD_DORM, None, 0)[0], "寝室卫生打扫", "日常劳动", "2026年秋季学期", "待审核", "来源:寝室打扫日常收集表")
write_csv("07_劳动教育参与学时表.csv", hours_hdr, hours_rows)
print(f"  LD hours kept(2023级): 汇总{n_sum} 周7={n_w7} 周8={n_w8} 周10={n_w10} 寝室={n_dm}")

# 更新项目表参与人数
proj_count = {}
for h in hours_rows:
    proj_count[h[4]] = proj_count.get(h[4], 0) + 1
pm = {"旗帜活动(每周教室/公共区域打扫)": ["第七周旗帜活动", "第十周旗帜活动"],
      "网信活动(教室/机房清扫)": [], "菜园子翻耕/浇水(青年劳动实践基地)": [],
      "寝室卫生打扫(日常)": ["寝室卫生打扫"], "权益提案活动": ["权益提案活动"],
      "劳动理论课(01-08章节)": ["劳动理论课"]}
for row in projects:
    keys = pm.get(row[0], [row[0]])
    cnt = sum(proj_count.get(k, 0) for k in keys)
    if row[0].startswith("旗帜活动"):
        cnt = proj_count.get("第七周旗帜活动", 0) + proj_count.get("第十周旗帜活动", 0)
    elif row[0].startswith("寝室卫生"):
        cnt = proj_count.get("寝室卫生打扫", 0)
    elif row[0].startswith("权益提案"):
        cnt = proj_count.get("权益提案活动", 0)
    elif row[0].startswith("劳动理论"):
        cnt = proj_count.get("劳动理论课", 0)
    else:
        cnt = 0
    row[9] = cnt if cnt else "待统计"
# rewrite projects csv with counts
write_csv("06_劳动教育项目表.csv", proj_hdr, projects)

write_csv("08_劳动教育基地排班表.csv",
          ["基地名称", "联系人", "地址", "安全条件", "排班日期", "班级", "人数", "交通方式",
           "保险状态", "安全告知状态", "审批状态", "备注"], [])
write_csv("09_劳动教育考核材料表.csv",
          ["项目名称", "日期", "签到表", "照片", "劳动记录", "总结", "评价", "公示材料",
           "材料是否齐全", "归档位置", "备注"], [])

# ---------- 10. 跨域待办表 ----------
todo_hdr = ["编号", "工作域", "对象", "事项", "风险等级(高/中/低)", "截止日期", "负责人", "下次跟进日期", "状态(未开始/进行中/已完成/已关闭)", "来源"]
ret = [r for r in waizhu if r["审核状态"] == "已退回"]
bad_phone = [r for r in roster if len(r["联系电话"]) != 11 or not r["联系电话"].isdigit()]
downgraded = [r for r in roster if r["专业"] == "降级"]
todos = [
    ["T01", "毕业班", "全体238人", "就业签约信息表未建立：单位/岗位/三方状态等全部待补录", "高", PEND, "辅导员", PEND, "未开始", "扫描发现"],
    ["T02", "毕业班", "全体238人", "毕业去向排查：毕业意向已从职业规划问卷回填，待定学生（含就业/考公意向者）需确认最终去向", "高", PEND, "辅导员", PEND, "未开始", "自动生成"],
    ["T03", "毕业班", f"{ret[0]['姓名']}({ret[0]['学号'][-4:]})" if ret else "-", "外住审核状态为[已退回]但仍在名单中，需跟进", "高", PEND, "辅导员", PEND, "未开始", "外住学生名单"],
    ["T04", "毕业班", "24名推免学生", "推免去向已确认；录取学校/公示状态/是否放弃/材料状态待补充", "中", PEND, "辅导员", PEND, "进行中", "报考确认表"],
    ["T05", "毕业班", "49名外住学生", "外住起止日期规范化、紧急联系人、安全告知待补充；审批到期提醒无法生成", "中", PEND, "辅导员", PEND, "未开始", "外住学生名单"],
    ["T06", "毕业班", f"{len(bad_phone)}人", "手机号位数异常待核对更正", "中", PEND, "辅导员", PEND, "未开始", "数据质量检查"],
    ["T07", "毕业班", f"{len(downgraded)}人", "专业为[降级]的学生班级归属待确认", "中", PEND, "辅导员", PEND, "未开始", "数据质量检查"],
    ["T08", "毕业班", "2人", "学生信息与一标三实两份名单不一致：一标三实含2022******32、学生信息含2022******11（均为2022级学号），需确认降转学生归属", "中", PEND, "辅导员", PEND, "未开始", "数据质量检查"],
    ["T09", "羽毛球协会", "协会整体", "会员/活动/赛事结构化数据全部缺失，需建立台账", "高", PEND, "辅导员", PEND, "未开始", "扫描发现"],
    ["T10", "羽毛球协会", "2023上半年预算", "预算余额、支出、报销状态待补充核对", "低", PEND, "辅导员", PEND, "未开始", "羽协预算表"],
    ["T11", "劳动教育", "项目表12个项目", "负责教师/基地/学时/考核方式等字段待补充", "中", PEND, "辅导员", PEND, "进行中", "扫描发现"],
    ["T12", "劳动教育", "汇总表-公益sheet", "8条学号异常记录(如4/8等)待核对修正", "中", PEND, "辅导员", PEND, "未开始", "数据质量检查"],
    ["T13", "劳动教育", "2023级参与学生", "学时数未记录(仅二课加分)，学时统计口径待定并补录", "中", PEND, "辅导员", PEND, "未开始", "扫描发现"],
    ["T14", "劳动教育", "基地排班/考核材料", "基地排班表与考核材料表为空，需建立", "低", PEND, "辅导员", PEND, "未开始", "扫描发现"],
    ["T15", "毕业班", "5人", "职业规划问卷意向与报考确认不一致：4人已报名考研但意向为就业/考公，1人推免但意向未填推免，需逐一核对", "中", PEND, "辅导员", PEND, "未开始", "数据交叉检查"],
]
write_csv("10_跨域待办表.csv", todo_hdr, todos)

print("ALL DONE. gender_conflict:", gender_conflict)
print(f"问卷意向匹配: {matched_plan}/{len(roster)} 未匹配: {plan_unmatched[:10]}")
