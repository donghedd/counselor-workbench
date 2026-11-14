# -*- coding: utf-8 -*-
"""CSV <-> SQLite 双向同步。
- python db_sync.py csv2db   ：data/*.csv → 辅导员工作台.db（以 CSV 为准，适合在 Excel 改完 CSV 后）
- python db_sync.py db2csv   ：辅导员工作台.db → data/*.csv（以数据库为准，适合用 DB Browser/SQL 改完后）
只读写输出目录，不动任何原始文件。
"""
import csv, os, sqlite3, sys

from app_paths import APP_HOME as HERE

DATA = os.path.join(HERE, "data")
DB = os.path.join(HERE, "辅导员工作台.db")

# key -> (CSV 文件名, 数据库表名)
TABLES = [
    ("students", "01_学生主表.csv", "学生主表"),
    ("todos", "10_跨域待办表.csv", "跨域待办表"),
    ("talks", "11_谈话跟进表.csv", "谈话跟进表"),
    ("warnings", "14_学业警示.csv", "学业警示表"),
]
KEY2TABLE = {k: (fn, tn) for k, fn, tn in TABLES}


def q(name):
    """SQLite 标识符加引号（列名/表名可能含中文、括号）"""
    return '"' + name.replace('"', '""') + '"'


def read_csv(fn):
    p = os.path.join(DATA, fn)
    if not os.path.exists(p):
        return [], []
    with open(p, encoding="utf-8-sig", newline="") as f:
        rd = list(csv.reader(f))
    if not rd:
        return [], []
    return rd[0], rd[1:]


def csv2db(only_keys=None):
    con = sqlite3.connect(DB)
    cur = con.cursor()
    n = 0
    for key, fn, tn in TABLES:
        if only_keys and key not in only_keys:
            continue
        header, rows = read_csv(fn)
        if not header:
            continue
        cur.execute("DROP TABLE IF EXISTS %s" % q(tn))
        cols = ", ".join("%s TEXT" % q(h) for h in header)
        cur.execute("CREATE TABLE %s (%s)" % (q(tn), cols))
        ph = ", ".join("?" * len(header))
        # 行长度对齐
        norm = [(r + [""] * len(header))[: len(header)] for r in rows]
        cur.executemany("INSERT INTO %s VALUES (%s)" % (q(tn), ph), norm)
        n += len(norm)
        print("  → %s：%d 行" % (tn, len(norm)))
    con.commit()
    con.close()
    print("csv2db 完成，共 %d 行 → %s" % (n, DB))


def db2csv():
    if not os.path.exists(DB):
        print("数据库不存在，请先运行：python db_sync.py csv2db")
        return
    con = sqlite3.connect(DB)
    cur = con.cursor()
    n = 0
    for key, fn, tn in TABLES:
        try:
            cur.execute("SELECT * FROM %s" % q(tn))
        except sqlite3.OperationalError:
            print("  ! 数据库缺少表：%s，跳过" % tn)
            continue
        header = [d[0] for d in cur.description]
        rows = cur.fetchall()
        with open(os.path.join(DATA, fn), "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(header)
            for r in rows:
                w.writerow(["" if v is None else str(v) for v in r])
        n += len(rows)
        print("  → %s：%d 行" % (fn, len(rows)))
    con.close()
    print("db2csv 完成，共 %d 行 → data/*.csv" % n)


def sync_tables_from_csv(keys):
    """看板保存后，仅把被修改的表从 CSV 刷回数据库，避免覆盖其他表在库里的人工改动"""
    csv2db(only_keys=set(keys))


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "csv2db"
    if mode == "csv2db":
        csv2db()
    elif mode == "db2csv":
        db2csv()
    else:
        print("用法：python db_sync.py [csv2db|db2csv]")
