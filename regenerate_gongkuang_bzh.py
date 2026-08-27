# -*- coding: utf-8 -*-
"""供矿作业区 4 位班组长 1-7 月履职履责考评表 —— 无界面批量重新生成。

复用了班组长工具（v2.0.2）的核心逻辑 kaoping_core.py，等价于在 GUI 里依次执行：
  读旧考评表（取自评描述 / 材料说明 / 支撑材料）→ 读月度履职履责表 xlsx（取评分）
  → 一键生成（安全绩效 20 分单行共用、上级评分 = 自评得分）。

数据口径（与用户确认）：
  1. 6 月 / 7 月「线上翻车机清底班」班组长以 xlsx 为准为 曹光奇（1-5 月为 刘安平）；
  2. 各考评项评分全部以月度履职履责表 xlsx 为准；
  3. 自评描述 / 材料说明 / 支撑材料从现有 .doc 考评表读取保留；
     （曹光奇 6/7 月沿用同岗位前任 刘安平 6/7 月考评表的描述与材料）
"""

import os
import sys
import time
import json
import io
import shutil
import tempfile

import pythoncom

KAOPING_DIR = r"C:\Users\ABaLaQiYaShanMaiI\OneDrive\Desktop\-XingDaLianTie-LvZhiLvZe-Python-\班组长安全生产责任制履职"
sys.path.insert(0, KAOPING_DIR)
import kaoping_core as k  # noqa: E402

BASE = r"E:\2.武钢兴达工作\2.月度固定工作\2026\2026.08\2.总包单位和供应商检查\1.3安全管理组织机构\06_履职履责（月度）\供矿作业区1-7月履职履责"
OUT_ROOT = os.path.join(BASE, "重新生成_供矿班组长")
TEMPLATE_SRC = os.path.join(KAOPING_DIR, "班组长安全生产责任制履职清单考评表（模板）.doc")
EXTRACT_DIR = os.path.join(BASE, "_tmp_regenerate_materials")

# Word COM 无法直接打开 OneDrive 路径下的模板（报“命令失败”），
# 生成前先复制到本地临时目录，用本地副本打开。
TEMPLATE = os.path.join(tempfile.gettempdir(), "_bzh_kaoping_template.doc")

YEAR = 2026
PATTERN = "班组长安全生产责任制履职清单考评表（{XXX}{X}月）.doc"

# 每月 4 位班组长（6/7 月线上翻车机清底班换为曹光奇）
PEOPLE = {
    1: ["刘安平", "李耀东", "王从虎", "王贻文"],
    2: ["刘安平", "李耀东", "王从虎", "王贻文"],
    3: ["刘安平", "李耀东", "王从虎", "王贻文"],
    4: ["刘安平", "李耀东", "王从虎", "王贻文"],
    5: ["刘安平", "李耀东", "王从虎", "王贻文"],
    6: ["曹光奇", "李耀东", "王从虎", "王贻文"],
    7: ["曹光奇", "李耀东", "王从虎", "王贻文"],
}

LOG_PATH = os.path.join(OUT_ROOT, "_regenerate_log.txt")


def _src_name(name):
    """曹光奇无历史考评表，沿用同岗位前任刘安平的考评表取描述/材料。"""
    return "刘安平" if name == "曹光奇" else name


def _find_xlsx(month):
    d = os.path.join(BASE, "%d月" % month)
    for fn in sorted(os.listdir(d)):
        if fn.lower().endswith(".xlsx"):
            return os.path.join(d, fn)
    raise FileNotFoundError("第%d月目录未找到履职履责表 xlsx" % month)


def _doc_path(name, month):
    return os.path.join(BASE, "%d月" % month, "班组长",
                        "班组长安全生产责任制履职清单考评表（%s%d月）.doc" % (name, month))


def _clean_paths(paths):
    return [p for p in (paths or []) if p and os.path.exists(p)]


def build_items(scores, extracted):
    """按『评分取 xlsx、描述/材料取旧表』的口径组装 17 项 items。"""
    items = {}
    for idx in range(1, k.N_ITEMS + 1):
        s = extracted.get("items", {}).get(idx) or {}
        sc = (scores.get(idx) or {}).get("score", "")
        item = {
            "desc": s.get("desc", ""),
            "score": sc,                                   # 自评得分 <- xlsx
            "material_text": s.get("material_text", ""),
            "materials": _clean_paths(s.get("materials")),
            "eval_desc": s.get("eval_desc", ""),           # 评价描述 <- 旧表保留
            "super_score": sc,                             # 上级评分 <- xlsx(=自评得分)
        }
        if idx == 1:                                       # 安全绩效双行：第 2 行仅材料
            sub = s.get("sub") or {}
            item["sub"] = {"materials": _clean_paths(sub.get("materials"))}
        items[idx] = item
    return items


def main():
    pythoncom.CoInitialize()
    os.makedirs(OUT_ROOT, exist_ok=True)
    # 模板复制到本地临时目录，规避 OneDrive 路径下 Word COM 打不开的问题
    shutil.copy2(TEMPLATE_SRC, TEMPLATE)
    log = []
    ok, fail = [], []
    t_start = time.time()

    for month in range(1, 8):
        xlsx = _find_xlsx(month)
        for name in PEOPLE[month]:
            t0 = time.time()
            tag = "%d月 %s" % (month, name)
            try:
                scores = k.read_eval_scores(xlsx, name)["items"]
                src_doc = _doc_path(_src_name(name), month)
                if not os.path.exists(src_doc):
                    raise FileNotFoundError("缺少描述来源考评表：" + src_doc)
                extracted = k.extract_kaoping_doc(src_doc, out_dir=EXTRACT_DIR)
                items = build_items(scores, extracted)
                out_dir = os.path.join(OUT_ROOT, "%d月" % month)
                os.makedirs(out_dir, exist_ok=True)
                out_name = k.build_filename(PATTERN, YEAR, month, name)
                out_path = os.path.join(out_dir, out_name)
                k.generate_doc(TEMPLATE, out_path, name, month, items, year=YEAR)
                ok.append((month, name, out_path))
                msg = "[OK] %s -> %s (%.1fs)" % (tag, os.path.basename(out_path), time.time() - t0)
                print(msg)
                log.append(msg)
            except Exception as e:
                fail.append((month, name, str(e)))
                msg = "[FAIL] %s : %s" % (tag, e)
                print(msg)
                log.append(msg)

    summary = {
        "total": len(ok) + len(fail),
        "ok": len(ok),
        "fail": len(fail),
        "ok_list": [{"month": m, "name": n, "path": p} for m, n, p in ok],
        "fail_list": [{"month": m, "name": n, "error": e} for m, n, e in fail],
        "elapsed_sec": round(time.time() - t_start, 1),
    }
    with io.open(os.path.join(OUT_ROOT, "_regenerate_summary.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    with io.open(LOG_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(log) + "\n")
    print("=" * 60)
    print("完成：成功 %d / 失败 %d，总耗时 %.1fs" % (len(ok), len(fail), summary["elapsed_sec"]))
    for m, n, e in fail:
        print("  失败：%d月 %s -> %s" % (m, n, e))


if __name__ == "__main__":
    main()
