#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""安全员履职考评表生成核心逻辑 v1.0.0"""

import os
import re
import sys
from datetime import datetime

import win32com.client

TEMPLATE_NAME = "安全员安全生产责任制履职清单考评表（模板）.doc"

ITEM_ROWS = {
    1: 3, 2: 5, 3: 6, 4: 7, 5: 8, 6: 9,
    7: 10, 8: 11, 9: 12, 10: 13, 11: 14, 12: 15,
}
TOTAL_ROW = 16
HEADER_ROW = 1
COL_SELF_DESC = 6
COL_SELF_SCORE = 7
COL_MATERIAL = 8
COL_EVAL_DESC = 9
COL_EVAL_SCORE = 10
IMG_MAX_WIDTH_CM = 2.5
IMG_MAX_HEIGHT_CM = 2.6
XLSX_SHEET = "安全员月度履职评价表"
XLSX_SCORE_COLS = {1: "E", 2: "G", 3: "I", 4: "K", 5: "M", 6: "O",
                   7: "Q", 8: "S", 9: "U", 10: "W", 11: "Y", 12: "AA"}
XLSX_DESC_COLS = {1: "F", 2: "H", 3: "J", 4: "L", 5: "N", 6: "P",
                  7: "R", 8: "T", 9: "V", 10: "X", 11: "Z", 12: "AB"}
XLSX_NAME_COL = "D"
XLSX_TOTAL_COL = "AC"
DEFAULT_NAME_PATTERN = "安全员安全生产责任制履职清单考评表({X}月{XXX}).doc"


def find_template(template_path=None):
    """定位模板文件：显式路径 > 程序目录 > PyInstaller 资源目录"""
    candidates = []
    if template_path:
        candidates.append(template_path)
    base = os.path.dirname(os.path.abspath(__file__))
    meipass = getattr(sys, "_MEIPASS", None)
    candidates += [
        os.path.join(base, TEMPLATE_NAME),
        os.path.join(base, "..", TEMPLATE_NAME),
    ]
    if meipass:
        candidates.insert(0, os.path.join(meipass, TEMPLATE_NAME))
    for c in candidates:
        if c and os.path.exists(c):
            return c
    raise FileNotFoundError("找不到模板文件：" + TEMPLATE_NAME)


def build_filename(pattern, year, month, name, default_ext=".doc"):
    """按命名模板生成文件名。占位符：{Y}=年, {X}=月, {XXX}=姓名"""
    fn = pattern.strip()
    if not fn:
        fn = DEFAULT_NAME_PATTERN
    fn = fn.replace("{Y}", str(year)).replace("{年}", str(year))
    fn = fn.replace("{X}", str(month)).replace("{月份}", str(month))
    fn = fn.replace("{XXX}", str(name)).replace("{姓名}", str(name))
    if not fn.lower().endswith((".doc", ".docx")):
        fn += default_ext
    fn = re.sub(r'[\\/:*?"<>|]', "_", fn)
    return fn


def read_xlsx_scores(xlsx_path):
    """读取履职履责表.xlsx「安全员月度履职评价表」，返回人员列表。
    每项: {name, total, items: {1..12: {'score': str, 'desc': str}}}
    """
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if XLSX_SHEET not in wb.sheetnames:
        raise ValueError("xlsx 中找不到工作表：" + XLSX_SHEET)
    ws = wb[XLSX_SHEET]
    name_col_idx = ws[XLSX_NAME_COL][0].column
    persons = []
    for r in range(5, ws.max_row + 1):
        name = ws.cell(row=r, column=name_col_idx).value
        if name is None or not str(name).strip():
            continue
        name = str(name).strip()
        items = {}
        for idx in range(1, 13):
            sc = ws[XLSX_SCORE_COLS[idx] + str(r)].value
            dc = ws[XLSX_DESC_COLS[idx] + str(r)].value
            items[idx] = {
                "score": "" if sc is None else str(sc).strip(),
                "desc": "" if dc is None else str(dc).strip(),
            }
        total_cell = ws[XLSX_TOTAL_COL + str(r)].value
        persons.append({
            "name": name,
            "total": "" if total_cell is None else str(total_cell).strip(),
            "items": items,
        })
    if not persons:
        raise ValueError("「安全员月度履职评价表」中没有读取到人员数据")
    return persons


def _set_cell_text(cell, text):
    """写入单元格文本，保留模板原有字体/段落格式（Range 替换法）。"""
    text = "" if text is None else str(text)
    rng = cell.Range
    rng.End = rng.End - 1
    rng.Text = text


def _clear_cell_content(cell):
    """清空单元格内文字与图片，保留段落结构。"""
    for shp in list(cell.Range.InlineShapes):
        shp.Delete()
    rng = cell.Range
    rng.End = rng.End - 1
    rng.Text = ""


def _cm2pt(v):
    return v * 28.35


IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"}


def _is_image(path):
    """判断文件是否为可内嵌图片。"""
    return os.path.splitext(path)[1].lower() in IMAGE_EXTS


def _insert_image(cell, path):
    """在单元格末尾插入单张图片（内嵌），自动等比缩放。"""
    from PIL import Image

    with Image.open(path) as im:
        ow, oh = im.size
    if ow <= 0 or oh <= 0:
        return False
    tw = _cm2pt(IMG_MAX_WIDTH_CM)
    th = oh * tw / ow
    if th > _cm2pt(IMG_MAX_HEIGHT_CM):
        th = _cm2pt(IMG_MAX_HEIGHT_CM)
        tw = ow * th / oh
    rng = cell.Range
    rng.End = rng.End - 1      # 排除单元格结束符 \x07
    rng.Collapse(0)            # 折叠到单元格末尾内部
    shp = cell.Range.InlineShapes.AddPicture(FileName=path, LinkToFile=False,
                                             SaveWithDocument=True, Range=rng)
    shp.Width = round(tw, 1)
    shp.Height = round(th, 1)
    return True


def _insert_ole(cell, path):
    """在单元格末尾插入文件为 OLE 嵌入对象（图标+文件名，可双击打开）。
    保持 Word 自然显示尺寸（图标+文件名标签），单元格内自动换行适配。
    """
    name = os.path.basename(path)
    rng = cell.Range
    rng.End = rng.End - 1
    rng.Collapse(0)
    shp = cell.Range.InlineShapes.AddOLEObject(
        FileName=path, LinkToFile=False, DisplayAsIcon=True, IconLabel=name,
        Range=rng)
    return True


def _insert_materials(cell, material_paths):
    """向单元格末尾依次插入支撑材料：
    图片 -> 内嵌图片；其他文件 -> OLE 嵌入对象；失败则退化为文件名文字。
    """
    valid = [p for p in material_paths if p and os.path.exists(p)]
    for idx, path in enumerate(valid):
        try:
            if _is_image(path):
                _insert_image(cell, path)
            else:
                _insert_ole(cell, path)
        except Exception:
            rng = cell.Range
            rng.End = rng.End - 1
            rng.Collapse(0)
            rng.Text = os.path.basename(path)
        if idx < len(valid) - 1:
            r2 = cell.Range
            r2.End = r2.End - 1
            r2.Collapse(0)
            r2.InsertParagraphAfter()


def _fill_header(cell, name, month):
    """替换表头 R1 中的姓名与评价月份。
    用标签区间截取法，避免正则组引用歧义。
    """
    raw = cell.Range.Text
    body = raw.replace("\x07", "").rstrip("\r")

    # 姓名：考评对象（安全员）：...管理者姓名
    m = re.search(r"(考评对象（安全员）：)\s*", body)
    if m:
        start = m.end()
        rest = body[start:]
        next_label = rest.find("管理者姓名")
        seg = rest if next_label < 0 else rest[:next_label]
        trailing = re.search(r"\s*$", seg)
        trailing_ws = trailing.group(0) if trailing else ""
        body = body[:start] + str(name).strip() + trailing_ws + rest[len(seg):]

    # 月份：评价月份：...评价人员签字
    m = re.search(r"(评价月份：)\s*", body)
    if m:
        start = m.end()
        rest = body[start:]
        next_label = rest.find("评价人员签字")
        seg = rest if next_label < 0 else rest[:next_label]
        trailing = re.search(r"\s*$", seg)
        trailing_ws = trailing.group(0) if trailing else ""
        body = body[:start] + str(month).strip() + "月" + trailing_ws + rest[len(seg):]

    rng = cell.Range
    rng.End = rng.End - 1
    rng.Text = body


def _sum_scores(items, key):
    """对某字段的数值求和（用于合计），忽略非数值；无有效数值返回空。"""
    total = 0
    has_val = False
    for idx in range(1, 13):
        v = (items.get(idx) or {}).get(key)
        if v is None:
            continue
        try:
            total += float(str(v).strip().rstrip("分"))
            has_val = True
        except (ValueError, TypeError):
            continue
    if not has_val:
        return ""
    return str(int(total)) if total == int(total) else str(total)


def generate_doc(template_path, output_path, name, month, items, year=None):
    """核心生成：打开模板副本 -> 填充 -> 另存为 .doc。
    items: {1..12: {'desc','score','material_text','material_images':[path],
                    'eval_desc','super_score'}}
    """
    if year is None:
        year = datetime.now().year
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(template_path))
        tb = doc.Tables(1)

        _fill_header(tb.Cell(HEADER_ROW, 1), name, month)

        for idx in range(1, 13):
            row = ITEM_ROWS[idx]
            item = items.get(idx) or {}
            _set_cell_text(tb.Cell(row, COL_SELF_DESC), item.get("desc", ""))
            _set_cell_text(tb.Cell(row, COL_SELF_SCORE), item.get("score", ""))
            mat_cell = tb.Cell(row, COL_MATERIAL)
            _clear_cell_content(mat_cell)
            mat_text = (item.get("material_text") or "").strip()
            materials = [p for p in (item.get("materials")
                                     or item.get("material_images") or [])
                         if p and os.path.exists(p)]
            if mat_text:
                _set_cell_text(mat_cell, mat_text)
            if materials:
                if mat_text:
                    r2 = mat_cell.Range
                    r2.Collapse(0)
                    r2.InsertParagraphAfter()
                _insert_materials(mat_cell, materials)
            _set_cell_text(tb.Cell(row, COL_EVAL_DESC), item.get("eval_desc", ""))
            _set_cell_text(tb.Cell(row, COL_EVAL_SCORE), item.get("super_score", ""))

        self_total = _sum_scores(items, "score")
        super_total = _sum_scores(items, "super_score")
        _set_cell_text(tb.Cell(TOTAL_ROW, COL_SELF_DESC), self_total)
        _set_cell_text(tb.Cell(TOTAL_ROW, COL_EVAL_DESC), super_total)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.SaveAs2(os.path.abspath(output_path), 0)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        word.Quit()
    return output_path


if __name__ == "__main__":
    tpl = find_template()
    print("模板:", tpl)
    items = {i: {"desc": "", "score": "", "material_text": "",
                 "material_images": [], "eval_desc": "", "super_score": ""}
             for i in range(1, 13)}
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_test_output.doc")
    generate_doc(tpl, out, "测试员", 8, items, year=2026)
    print("已生成:", out, os.path.getsize(out), "bytes")




