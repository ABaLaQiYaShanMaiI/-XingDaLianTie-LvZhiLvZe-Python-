#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""安全员履职考评表生成核心逻辑 v1.0.4"""

import json
import os
import re
import sys
from datetime import datetime

import win32com.client
from win32com.client import gencache

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
IMG_MAX_SIZE_CM = 2.0      # 图片/OLE 对象显示长宽上限（等比缩放，长宽均不超过该值）
DEFAULT_NAME_PATTERN = "安全员安全生产责任制履职清单考评表({X}月{XXX}).doc"


def find_template(template_path=None):
    """定位模板文件：显式路径 > exe/脚本目录 > 上一级目录 > _MEIPASS(旧版兼容)。

    模板为外置文件：打包成 exe 后需与 exe 同目录（或上一级目录）分发，
    不再内置到 exe 中；源码运行时按脚本所在目录定位。
    """
    if template_path and os.path.exists(template_path):
        return template_path

    # 打包(exe)时以 exe 所在目录为基准；源码运行时以脚本所在目录为基准
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    candidates = [
        os.path.join(base, TEMPLATE_NAME),
        os.path.join(os.path.dirname(base), TEMPLATE_NAME),
    ]
    # 兼容旧版打包：模板仍被内置在 exe 资源目录(_MEIPASS)中的情况
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, TEMPLATE_NAME))
    for c in candidates:
        if c and os.path.exists(c):
            return c
    raise FileNotFoundError(
        f"未找到模板文件 '{TEMPLATE_NAME}'，请确保模板在以下目录之一：{chr(10)}"
        f"  - {base}{chr(10)}"
        f"  - {os.path.dirname(base)}"
    )


def build_filename(pattern, year, month, name, default_ext=".doc"):
    """按命名模板生成文件名。占位符：{Y}=年, {X}=月, {XXX}=姓名"""
    fn = pattern.strip()
    if not fn:
        fn = DEFAULT_NAME_PATTERN
    fn = fn.replace("{Y}", str(year)).replace("{年}", str(year))
    fn = fn.replace("{X}", str(month)).replace("{月份}", str(month))
    fn = fn.replace("{XXX}", str(name)).replace("{姓名}", str(name))
    fn = re.sub(r'[\\/:*?"<>|]', "_", fn)
    # 分离扩展名，只对主干做 Windows 消毒（去掉结尾点/空格）
    if fn.lower().endswith(".docx"):
        stem, ext = fn[:-5], fn[-5:]
    elif fn.lower().endswith(".doc"):
        stem, ext = fn[:-4], fn[-4:]
    else:
        stem, ext = fn, default_ext
    stem = stem.rstrip(" .") or "未命名"
    fn = stem + ext
    # Windows 保留设备名（CON/PRN/AUX/NUL/COM1~9/LPT1~9），按第一个点前的主名判断
    stem0 = fn.split(".", 1)[0].upper()
    reserved = {"CON", "PRN", "AUX", "NUL"} | {f"COM{i}" for i in range(1, 10)} \
        | {f"LPT{i}" for i in range(1, 10)}
    if stem0 in reserved:
        fn = "_" + fn
    return fn


# ============ 支撑材料文件夹自动匹配 ============
# 每个考评项的关键词表（按文件名/相对路径中的子文件夹名匹配）。
# 第一词为该考评项标题词（权重 2），其余为扩展词（权重 1）；
# 界面“编辑权重”按钮可打开本地 material_rules.json 调整（修改后重启生效）。
# 关键词基于各月实际文件夹命名归纳（如 扫雷→⑩、安全月报→⑫、安监网/安检网录入→⑪）。
ITEM_MATCH_RULES = {
    1: ["安全绩效", "绩效", "考核", "事故", "工伤", "轻伤", "月度会"],
    2: ["管理体系", "体系", "制度", "责任制", "标准化"],
    3: ["教育培训", "培训", "教育", "学习", "课件", "交底", "考试", "双考", "应知", "安全活动"],
    4: ["会议活动", "会议", "班会", "班前会", "活动", "纪要", "部署", "发言"],
    5: ["检查改进", "检查", "巡查", "督查", "整改", "找茬", "大检查", "自查", "排查表"],
    6: ["应急演习", "应急", "演练", "演习", "预案"],
    7: ["监督执行", "监督", "值守", "带班", "值班", "旁站", "应急值守"],
    8: ["安全能力", "能力", "取证", "特种作业", "资格证", "上岗证", "特种工", "证书", "规程", "复审"],
    9: ["危险源", "风险", "危害", "辨识", "评估"],
    10: ["隐患排查", "隐患", "排查", "扫雷", "有限空间"],
    11: ["违章管理", "违章", "违规", "三违", "违纪", "反思", "安监网", "安检网", "禁令"],
    12: ["其它", "其他", "主题", "总结", "小结", "月报", "台账", "人数", "体检", "协议", "资料",
         "汇报", "保险", "方案", "通知", "清单"],
}
# 扫描时跳过：临时文件、考评表/履责表自身、无意义扩展名
SKIP_FILE_PARTS = ("~$",)
SKIP_NAME_PARTS = ("考评表", "履职履责表", "履职评价表", "履责表", "履职履责")
SKIP_EXTS = {".lnk", ".url", ".ini", ".tmp", ".db", ".py", ".pyc", ".md"}

# 本地文件名权重配置文件（与 exe/脚本同目录，首次运行自动生成）
MATERIAL_RULES_FILE = "material_rules.json"


def ensure_material_rules_file(base_dir=None):
    """确保本地文件名权重配置文件存在（首次自动生成），返回其路径。"""
    base_dir = base_dir or os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base_dir, MATERIAL_RULES_FILE)
    if not os.path.exists(path):
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump({str(k): list(v) for k, v in ITEM_MATCH_RULES.items()},
                          f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return path


def load_material_rules(path=None):
    """读取本地文件名权重配置；文件缺失/非法时回退内置 ITEM_MATCH_RULES。"""
    if path and os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            rules = {}
            for k, v in data.items():
                idx = int(str(k).split(".")[0])          # 兼容 "1" / "1. 安全绩效"
                if isinstance(v, dict):                  # 兼容 {"weight":2,"keywords":[...]}
                    kw = v.get("keywords") or v.get("关键词") or []
                else:
                    kw = v
                kw = [str(x).strip() for x in (kw or []) if str(x).strip()]
                if kw:
                    rules[idx] = kw
            if rules:
                return rules
        except Exception:
            pass
    return dict(ITEM_MATCH_RULES)


def score_materials_file(rel_path, rules):
    """按权重打分并返回 (考评项编号, 得分)；无匹配返回 (None, 0)。
    每个考评项第 1 词权重 2、其余词权重 1；命中分高者优先，同分取编号小者。"""
    text = rel_path.rsplit(".", 1)[0]                    # 去扩展名，保留相对路径（含子文件夹名）
    scores = {}
    for idx, kws in rules.items():
        s = 0
        for i, kw in enumerate(kws):
            if kw and kw in text:
                s += 2 if i == 0 else 1
        if s:
            scores[idx] = s
    if not scores:
        return None, 0
    idx = max(scores, key=lambda k: (scores[k], -k))
    return idx, scores[idx]


def match_materials_file(rel_path, rules=None):
    """按关键词把文件相对路径匹配到考评项，返回编号(1~12)；无匹配返回 None。"""
    idx, _score = score_materials_file(rel_path, rules or load_material_rules())
    return idx


def find_month_subfolder(folder, year, month):
    """在年度/根目录中自动定位当月子文件夹（如 2026.08 / 2026.8 / 8月）。
    找到返回其路径，否则原样返回 folder。"""
    targets = {f"{year}.{month:02d}", f"{year}.{month}", f"{year}-{month:02d}",
               f"{year}-{month}", f"{year}年{month}月", f"{month}月"}
    try:
        for name in os.listdir(folder):
            full = os.path.join(folder, name)
            if os.path.isdir(full) and name in targets:
                return full
    except Exception:
        pass
    return folder


def scan_materials_folder(folder, max_per_item=15, rules=None):
    """递归扫描支撑材料文件夹，按文件名权重自动匹配到各考评项。

    返回 (result, unmatched)：
      result   = {1..12: [文件绝对路径...]}（每项最多 max_per_item 个，按权重分降序）
      unmatched= [未匹配(或超限)的文件绝对路径...]
    """
    folder = os.path.abspath(folder)
    if not os.path.isdir(folder):
        raise ValueError("支撑材料文件夹不存在：" + folder)
    rules = rules or load_material_rules()
    buckets = {i: [] for i in range(1, 13)}              # {idx: [(score, path)...]}
    unmatched = []
    for root, _dirs, files in os.walk(folder):
        for fn in sorted(files):
            if fn.startswith(SKIP_FILE_PARTS):
                continue
            if any(p in fn for p in SKIP_NAME_PARTS):
                continue
            if os.path.splitext(fn)[1].lower() in SKIP_EXTS:
                continue
            full = os.path.join(root, fn)
            rel = os.path.relpath(full, folder)
            idx, score = score_materials_file(rel, rules)
            if idx is None:
                unmatched.append(full)
            elif len(buckets[idx]) < max_per_item:
                buckets[idx].append((score, full))
            else:
                unmatched.append(full)  # 超过每项上限，按未匹配提示用户手动处理
    result = {i: [p for _s, p in sorted(bucket, key=lambda t: -t[0])]
              for i, bucket in buckets.items()}
    return result, unmatched


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


def is_image(path):
    """判断文件是否为可内嵌图片。"""
    return os.path.splitext(path)[1].lower() in IMAGE_EXTS


def _insert_image(cell, path):
    """在单元格末尾插入单张图片（内嵌），自动等比缩放，长宽均不超过 IMG_MAX_SIZE_CM。"""
    from PIL import Image

    with Image.open(path) as im:
        ow, oh = im.size
    if ow <= 0 or oh <= 0:
        return False
    max_pt = _cm2pt(IMG_MAX_SIZE_CM)
    scale = min(max_pt / ow, max_pt / oh)
    tw = round(ow * scale, 1)
    th = round(oh * scale, 1)
    rng = cell.Range
    rng.End = rng.End - 1      # 排除单元格结束符 \x07
    rng.Collapse(0)            # 折叠到单元格末尾内部
    shp = cell.Range.InlineShapes.AddPicture(FileName=path, LinkToFile=False,
                                             SaveWithDocument=True, Range=rng)
    shp.Width = tw
    shp.Height = th
    return True


def _fit_inline_size(shp):
    """将内嵌对象（图片/OLE）等比缩放，使长宽均不超过 IMG_MAX_SIZE_CM。

    实测：图片/Excel 等 OLE 可通过 Width/Height 赋值缩放，
    而 Word 文档类 OLE 图标会忽略 Width/Height 赋值，需改用
    ScaleWidth/ScaleHeight（相对原始尺寸的百分比）。
    """
    max_pt = _cm2pt(IMG_MAX_SIZE_CM)
    try:
        w, h = shp.Width, shp.Height
    except Exception:
        return
    if w <= 0 or h <= 0 or (w <= max_pt and h <= max_pt):
        return
    scale = min(max_pt / w, max_pt / h)
    # 方式一：直接设置宽高（对图片/Excel OLE 有效）
    try:
        shp.Width = round(w * scale, 1)
        shp.Height = round(h * scale, 1)
    except Exception:
        pass
    # 方式二：若宽高赋值未生效（Word 文档类 OLE），改用百分比缩放
    try:
        w2 = shp.Width
    except Exception:
        w2 = w
    if w2 <= 0 or w2 >= w - 0.5:
        try:
            shp.ScaleWidth = round(scale * 100, 1)
            shp.ScaleHeight = round(scale * 100, 1)
        except Exception:
            pass


def _insert_ole(cell, path):
    """在单元格末尾插入文件为 OLE 嵌入对象（图标+文件名，可双击打开）。
    保持 Word 自然显示尺寸（图标+文件名标签），随后等比缩小到 IMG_MAX_SIZE_CM 以内。
    """
    name = os.path.basename(path)
    rng = cell.Range
    rng.End = rng.End - 1
    rng.Collapse(0)
    shp = cell.Range.InlineShapes.AddOLEObject(
        FileName=path, LinkToFile=False, DisplayAsIcon=True, IconLabel=name,
        Range=rng)
    _fit_inline_size(shp)
    return True


def _insert_materials(cell, material_paths):
    """向单元格末尾依次插入支撑材料：
    图片 -> 内嵌图片；其他文件 -> OLE 嵌入对象；失败则退化为文件名文字。
    """
    valid = [p for p in material_paths if p and os.path.exists(p)]
    for idx, path in enumerate(valid):
        try:
            if is_image(path):
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


def _replace_slot(body, label, next_label, value):
    """替换 body 中「label ... next_label」区间内的值，保留区间首尾空白。
    - 若区间原值存在（如旧姓名），新值顶替原值，首尾空白不变；
    - 若区间全为空白（模板留空待填），新值居中放置。
    """
    li = body.find(label)
    if li < 0:
        return body
    start = li + len(label)
    ni = body.find(next_label, start)
    end = ni if ni >= 0 else len(body)
    seg = body[start:end]
    if not seg.strip():
        # 全空白：把空白对半分，值放中间
        n = len(seg)
        head, tail = n // 2, n - n // 2
        return body[:start] + seg[:head] + value + seg[n - tail:] + body[end:]
    head = len(seg) - len(seg.lstrip())
    tail = len(seg) - len(seg.rstrip())
    new_seg = seg[:head] + value + (seg[len(seg) - tail:] if tail else "")
    return body[:start] + new_seg + body[end:]


def _fill_header(cell, name, month):
    """替换表头 R1 中的姓名与评价月份。
    用标签区间截取法，避免正则组引用歧义；保留空白占位区，值居中写入。
    """
    raw = cell.Range.Text
    body = raw.replace("\x07", "").rstrip("\r")

    # 姓名：考评对象（安全员）：...管理者姓名
    body = _replace_slot(body, "考评对象（安全员）：", "管理者姓名", str(name).strip())
    # 月份：评价月份：...评价人员签字
    body = _replace_slot(body, "评价月份：", "评价人员签字", str(month).strip() + "月")

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


def _new_word_app():
    """创建 Word COM 实例。

    必须用静态类型库分派(gencache.EnsureDispatch)：动态分派(Dispatch)下
    InlineShapes.AddOLEObject 在无界面 Word 环境会挂起（实测复现）。
    PyInstaller 打包环境自动回退到动态分派，并优先把类型库缓存指向临时目录。
    """
    if getattr(sys, "frozen", False):
        try:
            gencache.is_readonly = False
            gencache.GetGeneratePath()
        except Exception:
            pass
    try:
        return gencache.EnsureDispatch("Word.Application")
    except Exception:
        return win32com.client.Dispatch("Word.Application")


def generate_doc(template_path, output_path, name, month, items, year=None,
                 progress_cb=None):
    """核心生成：打开模板副本 -> 填充 -> 另存为 .doc。
    items: {1..12: {'desc','score','material_text','material_images':[path],
                    'eval_desc','super_score'}}
    progress_cb: 可选回调，每处理完一个考评项调用 progress_cb(idx)，idx=1..12。
    """
    if year is None:
        year = datetime.now().year
    word = _new_word_app()
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(template_path))
        tb = doc.Tables(1)

        # 模板结构校验：防止模板被替换/改版后错格写入
        rows, cols = tb.Rows.Count, tb.Columns.Count
        if rows != TOTAL_ROW or cols != 10:
            raise ValueError(
                f"模板表格结构异常：期望 {TOTAL_ROW} 行×10 列，实际 {rows} 行×{cols} 列，"
                f"请更换正确模板（{TEMPLATE_NAME}）"
            )
        header = tb.Cell(HEADER_ROW, 1).Range.Text
        if "考评对象" not in header or "评价月份" not in header:
            raise ValueError(
                "模板表头缺少「考评对象（安全员）」或「评价月份」标签，请更换正确模板"
            )

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
            if progress_cb:
                try:
                    progress_cb(idx)
                except Exception:
                    pass

        self_total = _sum_scores(items, "score")
        super_total = _sum_scores(items, "super_score")
        _set_cell_text(tb.Cell(TOTAL_ROW, COL_SELF_DESC), self_total)
        _set_cell_text(tb.Cell(TOTAL_ROW, COL_EVAL_DESC), super_total)

        # 保存前统一缩放 C8 全部内嵌对象（图片/OLE），确保长宽均 ≤ IMG_MAX_SIZE_CM。
        # OLE 对象在插入瞬间尺寸可能未稳定，故在此兜底再缩放一次。
        for idx in range(1, 13):
            try:
                mat_cell = tb.Cell(ITEM_ROWS[idx], COL_MATERIAL)
                for shp in list(mat_cell.Range.InlineShapes):
                    _fit_inline_size(shp)
            except Exception:
                pass

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




