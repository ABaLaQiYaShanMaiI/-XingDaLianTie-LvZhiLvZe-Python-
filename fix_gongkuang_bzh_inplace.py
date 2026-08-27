# -*- coding: utf-8 -*-
"""原地修正供矿班组长考评表中的跨作业区/错别字描述（不重新抽取/嵌入支撑材料）。

直接打开目标目录里的 28 份 .doc，只对「自评描述(C6)」「评价描述(C9)」两列做文字替换，
保留支撑材料与格式原样不变。
"""

import os
import sys
import time
import json
import io

import pythoncom
import win32com.client
from win32com.client import gencache

KAOPING_DIR = r"C:\Users\ABaLaQiYaShanMaiI\OneDrive\Desktop\-XingDaLianTie-LvZhiLvZe-Python-\班组长安全生产责任制履职"
sys.path.insert(0, KAOPING_DIR)
import kaoping_core as k  # noqa: E402

OUT_ROOT = r"E:\2.武钢兴达工作\2.月度固定工作\2026\2026.08\2.总包单位和供应商检查\1.2用工管理\履职履责"

PEOPLE = {
    1: ["刘安平", "李耀东", "王从虎", "王贻文"],
    2: ["刘安平", "李耀东", "王从虎", "王贻文"],
    3: ["刘安平", "李耀东", "王从虎", "王贻文"],
    4: ["刘安平", "李耀东", "王从虎", "王贻文"],
    5: ["刘安平", "李耀东", "王从虎", "王贻文"],
    6: ["曹光奇", "李耀东", "王从虎", "王贻文"],
    7: ["曹光奇", "李耀东", "王从虎", "王贻文"],
}

FIXES = [
    ("原料分厂皮带工", "皮带机"),
    ("组织卸煤机司机开展起重设备安全培训", "组织翻车机司机开展设备安全培训"),
    ("煤库作业区", "供矿作业区"),
    ("隐患排，", "隐患排查，"),
    ("传答", "传达"),
]


def _fix_text(t):
    if not t:
        return t
    for old, new in FIXES:
        t = t.replace(old, new)
    return t


def _cell_text(tb, row, col):
    try:
        return tb.Cell(row, col).Range.Text.replace("\x07", "").replace("\r", "").strip()
    except Exception:
        return ""


def _set_cell_text(tb, row, col, text):
    cell = tb.Cell(row, col)
    rng = cell.Range
    rng.End = rng.End - 1
    rng.Text = "" if text is None else str(text)


def main():
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    changed = []
    errors = []
    t_start = time.time()
    try:
        for month in range(1, 8):
            for name in PEOPLE[month]:
                p = os.path.join(OUT_ROOT, "2026年%d月" % month, "班组长",
                                 "班组长安全生产责任制履职清单考评表（%s%d月）.doc" % (name, month))
                t0 = time.time()
                try:
                    d = word.Documents.Open(p)
                    tb = d.Tables(1)
                    hits = []
                    for idx in range(1, k.N_ITEMS + 1):
                        for row in k._item_rows(idx):
                            for col, label in ((k.COL_SELF_DESC, "desc"), (k.COL_EVAL_DESC, "eval")):
                                old = _cell_text(tb, row, col)
                                new = _fix_text(old)
                                if new != old:
                                    _set_cell_text(tb, row, col, new)
                                    hits.append("第%d项-%s: %s → %s" % (idx, label, old, new))
                    d.Save()
                    d.Close(False)
                    changed.append((month, name, hits))
                    print("[OK] %d月 %s 修正 %d 处 (%.1fs)" % (month, name, len(hits), time.time() - t0))
                    for h in hits:
                        print("     " + h)
                except Exception as e:
                    errors.append((month, name, str(e)))
                    print("[FAIL] %d月 %s : %s" % (month, name, e))
    finally:
        word.Quit()

    summary = {
        "fixed_files": len(changed),
        "errors": [{"month": m, "name": n, "error": e} for m, n, e in errors],
        "elapsed_sec": round(time.time() - t_start, 1),
        "detail": [{"month": m, "name": n, "hits": h} for m, n, h in changed],
    }
    with io.open(os.path.join(OUT_ROOT, "_fix_bzh_summary.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print("=" * 60)
    print("完成：修正 %d 个文件，错误 %d 个，总耗时 %.1fs" % (len(changed), len(errors), summary["elapsed_sec"]))


if __name__ == "__main__":
    main()
