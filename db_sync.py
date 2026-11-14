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


def check_header(fn, header):
    """表头体检：重复列名 / 空表头会让 CREATE TABLE 失败，
    而重建表时旧表已经被拆掉（DDL 在 sqlite 里默认自动提交，回滚不了）。
    所以先体检，有问题直接报错，数据库一个字节都不动。"""
    seen, dup = {}, []
    for h in header:
        k = str(h).strip()
        if not k:
            raise ValueError("%s：有空表头列（第 %d 列），请先在 Excel 里补上列名再重跑"
                             % (fn, header.index(h) + 1))
        if k in seen:
            dup.append(k)
        seen[k] = 1
    if dup:
        raise ValueError("%s：表头有重复列名（%s），建表会失败；"
                         "常见原因是 Excel 里复制粘贴多出一列，改好后重跑。数据库未改动。"
                         % (fn, "、".join(dup)))


def csv2db(only_keys=None):
    todo = [(key, fn, tn) for key, fn, tn in TABLES if not only_keys or key in only_keys]
    # 先全部体检，再动数据库：避免前几张表建成、后面失败，留下半套数据
    for key, fn, tn in todo:
        header, _ = read_csv(fn)
        if header:
            check_header(fn, header)
    con = sqlite3.connect(DB)
    con.isolation_level = None      # 自己管事务：显式 BEGIN 后 DDL 也受事务保护，失败能整体回滚
    cur = con.cursor()
    n = 0
    for key, fn, tn in todo:
        header, rows = read_csv(fn)
        if not header:
            continue
        cols = ", ".join("%s TEXT" % q(h) for h in header)
        ph = ", ".join("?" * len(header))
        # 行长度对齐
        norm = [(r + [""] * len(header))[: len(header)] for r in rows]
        cur.execute("BEGIN IMMEDIATE")
        try:
            cur.execute("DROP TABLE IF EXISTS %s" % q(tn))
            cur.execute("CREATE TABLE %s (%s)" % (q(tn), cols))
            cur.executemany("INSERT INTO %s VALUES (%s)" % (q(tn), ph), norm)
            cur.execute("COMMIT")
        except Exception:
            cur.execute("ROLLBACK")     # 旧表原样保留
            con.close()
            raise
        n += len(norm)
        print("  → %s：%d 行" % (tn, len(norm)))
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
        p = os.path.join(DATA, fn)
        tmp = p + ".tmp"
        with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(header)
            for r in rows:
                w.writerow(["" if v is None else str(v) for v in r])
        os.replace(tmp, p)      # 写完才替换，中途失败不会留下半截 CSV
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
