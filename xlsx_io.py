# -*- coding: utf-8 -*-
"""极简 xlsx 读写：只用标准库 zipfile + xml，不依赖 openpyxl / pandas。

- write_xlsx(sheets)  : 生成 .xlsx 字节流；sheet = (名称, 表头, 行数据, 列宽, 说明行)
- read_table(name, data): 读取 .xlsx（或 .csv）→ [ {列名: 值}, ... ]，兼容 Excel / WPS 保存的文件
"""
import csv
import io
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def col_letter(idx):
    """0 → A, 25 → Z, 26 → AA"""
    s, idx = "", idx + 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _tag(el):
    return el.tag.split("}")[-1]


def _cell(ref, value, style=0):
    """生成一个单元格 XML；数字用数值型，其余用内联字符串（避免公式/类型问题）"""
    if value is None:
        value = ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, style, value)
    text = escape(str(value))
    if not text:
        return '<c r="%s" s="%d"/>' % (ref, style)
    return '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, style, text)


_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="%s">
<fonts count="3">
<font><sz val="11"/><name val="微软雅黑"/><color rgb="FF1F2735"/></font>
<font><b/><sz val="11"/><name val="微软雅黑"/><color rgb="FFFFFFFF"/></font>
<font><sz val="10"/><name val="微软雅黑"/><color rgb="FF8A94A6"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF2A5BD7"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFDCE3EE"/></left><right style="thin"><color rgb="FFDCE3EE"/></right><top style="thin"><color rgb="FFDCE3EE"/></top><bottom style="thin"><color rgb="FFDCE3EE"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>""" % _NS_MAIN


def _sheet_xml(table):
    header, rows, widths = table["header"], table["rows"], table.get("widths") or []
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="%s">' % _NS_MAIN,
             '<sheetViews><sheetView tabSelected="1" workbookViewId="0">',
             '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
             '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="19"/>']
    if widths:
        cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w)
                       for i, w in enumerate(widths))
        parts.append("<cols>%s</cols>" % cols)
    parts.append("<sheetData>")
    cells = "".join(_cell(col_letter(i) + "1", h, 1) for i, h in enumerate(header))
    parts.append('<row r="1" ht="22" customHeight="1">%s</row>' % cells)
    for ri, row in enumerate(rows, start=2):
        cells = "".join(_cell(col_letter(i) + str(ri), row[i] if i < len(row) else "", 2)
                        for i in range(len(header)))
        parts.append('<row r="%d">%s</row>' % (ri, cells))
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


def write_xlsx(sheets):
    """sheets: [{"name":..., "header":[...], "rows":[[...]], "widths":[...]}]"""
    n = len(sheets)
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(n):
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1))
    ct.append("</Types>")

    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')

    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<workbook xmlns="%s" xmlns:r="%s">' % (_NS_MAIN, _NS_REL), "<sheets>"]
    wb_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<Relationships xmlns="%s">' % _NS_PKG_REL]
    for i, sh in enumerate(sheets):
        wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(sh["name"]), i + 1, i + 1))
        wb_rels.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>'
                       % (i + 1, _NS_REL, i + 1))
    wb.append("</sheets></workbook>")
    wb_rels.append('<Relationship Id="rIdS" Type="%s/styles" Target="styles.xml"/>' % _NS_REL)
    wb_rels.append("</Relationships>")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", "".join(wb))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
        z.writestr("xl/styles.xml", _STYLES)
        for i, sh in enumerate(sheets):
            z.writestr("xl/worksheets/sheet%d.xml" % (i + 1), _sheet_xml(sh))
    return buf.getvalue()


# ---------------- 读取 ----------------

def _shared_strings(z):
    out = []
    name = "xl/sharedStrings.xml"
    if name not in z.namelist():
        return out
    root = ET.fromstring(z.read(name))
    for si in root:
        if _tag(si) != "si":
            continue
        out.append("".join(t.text or "" for t in si.iter() if _tag(t) == "t"))
    return out


def _first_sheet_path(z):
    """按 workbook.xml + rels 定位第一个工作表；失败则回退 sheet1.xml"""
    try:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rid = None
        for el in wb.iter():
            if _tag(el) == "sheet":
                for k, v in el.attrib.items():
                    if k.endswith("}id") or k == "r:id":
                        rid = v
                        break
                break
        if rid:
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            for r in rels:
                if r.attrib.get("Id") == rid:
                    t = r.attrib.get("Target", "").lstrip("/")
                    return t if t.startswith("xl/") else "xl/" + t
    except Exception:
        pass
    return "xl/worksheets/sheet1.xml"


_NUM = re.compile(r"^-?\d+(\.\d+)?$")


def _cell_value(c, shared):
    t = c.attrib.get("t", "")
    if t == "inlineStr":
        return "".join(x.text or "" for x in c.iter() if _tag(x) == "t")
    v = None
    for el in c:
        if _tag(el) == "v":
            v = el.text
            break
    if v is None:
        return ""
    if t == "s":
        try:
            return shared[int(v)]
        except Exception:
            return ""
    if t == "b":
        return "TRUE" if v.strip() in ("1", "TRUE", "true") else "FALSE"
    if t in ("str", "e"):
        return v
    return v


def read_table(fname, data):
    """返回 [ {表头: 值} ]；支持 .xlsx/.xlsm 与 .csv/.txt"""
    if not data:
        raise ValueError("文件为空")
    if data[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(data))
        shared = _shared_strings(z)
        path = _first_sheet_path(z)
        root = ET.fromstring(z.read(path))
        grid = []
        for row in root.iter():
            if _tag(row) != "row":
                continue
            cells = {}
            for c in row:
                if _tag(c) != "c":
                    continue
                ref = c.attrib.get("r", "")
                m = re.match(r"^([A-Z]+)(\d+)$", ref)
                if not m:
                    continue
                letters = m.group(1)
                ci = 0
                for ch in letters:
                    ci = ci * 26 + (ord(ch) - 64)
                cells[ci - 1] = _cell_value(c, shared).strip()
            if not cells:
                continue
            width = max(cells) + 1
            grid.append([cells.get(i, "") for i in range(width)])
        # 去掉整体空行
        grid = [r for r in grid if any(str(v).strip() for v in r)]
    else:
        text = None
        for enc in ("utf-8-sig", "utf-8", "gb18030"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError("无法识别文件编码，请另存为 UTF-8 CSV 或 .xlsx")
        grid = [r for r in csv.reader(io.StringIO(text)) if any(str(v).strip() for v in r)]

    if not grid:
        raise ValueError("表格里没有内容")
    header = [str(h).strip() for h in grid[0]]
    out = []
    for r in grid[1:]:
        obj = {}
        for i, h in enumerate(header):
            if not h:
                continue
            obj[h] = str(r[i]).strip() if i < len(r) else ""
        out.append(obj)
    return header, out
