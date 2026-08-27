# -*- coding: utf-8 -*-
"""为安全员 王雪梅 1-7 月考评表补充支撑材料并重新生成。

支撑材料来源：供矿作业区各月支撑材料文件夹
  （供矿作业区1-7月履职履责\{月}月\班组长\对标甲方层级的履职履责\{月}月支撑材料），
  按安全员 material_rules.json 的 12 项关键词自动匹配。

筛选规则：
  1. 排除其他作业区的材料（文件名/路径含「原料」「煤库」）；
  2. 排除临时/副本文件（~$、副本）；
  3. 第1项「安全绩效」不配材料（无事故，描述自明）；
  4. 其余每项取匹配分最高的 1 个文件。
"""

import os
import sys
import time
import json
import io
import shutil
import tempfile

import pythoncom

KAOPING_DIR = r"C:\Users\ABaLaQiYaShanMaiI\OneDrive\Desktop\-XingDaLianTie-LvZhiLvZe-Python-\安全员安全生产责任制履职"
sys.path.insert(0, KAOPING_DIR)
import kaoping_core as k  # noqa: E402

BASE = r"E:\2.武钢兴达工作\2.月度固定工作\2026\2026.08\2.总包单位和供应商检查\1.3安全管理组织机构\06_履职履责（月度）\供矿作业区1-7月履职履责"
OUT_ROOT = r"E:\2.武钢兴达工作\2.月度固定工作\2026\2026.08\2.总包单位和供应商检查\1.2用工管理\履职履责"
TEMPLATE_SRC = os.path.join(KAOPING_DIR, "安全员安全生产责任制履职清单考评表（模板）.doc")
TEMPLATE = os.path.join(tempfile.gettempdir(), "_aqy_kaoping_template.doc")
RULES_PATH = os.path.join(KAOPING_DIR, "material_rules.json")

YEAR = 2026
PATTERN = "安全员安全生产责任制履职清单考评表（{XXX}{X}月）.doc"

DESCS = {
    1: "未发生轻伤及以上生产安全事故",
    2: "安全生产责任制和管理制度体系建立健全并有效运行",
    3: "组织开展供矿作业区安全教育培训并跟踪考核",
    4: "按时组织召开安全会议，开展安全主题活动",
    5: "定期开展安全检查，落实隐患整改闭环",
    6: "组织开展应急演练，提升应急处置能力",
    7: "监督检查安全生产制度和岗位规程执行情况",
    8: "具备岗位要求的安全生产知识和管理能力",
    9: "组织开展危险源辨识、风险评估与分级管控",
    10: "开展隐患排查治理，落实整改措施",
    11: "未完成违章查处",
    12: "完成上级交办的其他安全管理工作",
}


def _find_xlsx(month):
    d = os.path.join(BASE, "%d月" % month)
    for fn in sorted(os.listdir(d)):
        if fn.lower().endswith(".xlsx"):
            return os.path.join(d, fn)
    raise FileNotFoundError("第%d月目录未找到履职履责表 xlsx" % month)


def _mat_folder(month):
    return os.path.join(BASE, "%d月" % month, "班组长", "对标甲方层级的履职履责", "%d月支撑材料" % month)


def _filter_materials(files):
    out = []
    for f in files:
        b = os.path.basename(f)
        if b.startswith("~$") or "副本" in b:
            continue
        if "原料" in f or "煤库" in f:
            continue
        out.append(f)
    return out


def main():
    pythoncom.CoInitialize()
    shutil.copy2(TEMPLATE_SRC, TEMPLATE)
    rules = k.load_material_rules(RULES_PATH)
    log = []
    ok, fail = [], []
    t_start = time.time()
    for month in range(1, 8):
        xlsx = _find_xlsx(month)
        folder = _mat_folder(month)
        t0 = time.time()
        try:
            scores = k.read_eval_scores(xlsx, "王雪梅")["items"]
            result, _unmatched = k.scan_materials_folder(folder, rules=rules)
            materials = {}
            for idx in range(1, k.N_ITEMS + 1):
                if idx == 1:            # 安全绩效不配材料
                    materials[idx] = []
                    continue
                picked = _filter_materials(result.get(idx, []))[:1]
                materials[idx] = picked
            items = {}
            for idx in range(1, k.N_ITEMS + 1):
                sc = (scores.get(idx) or {}).get("score", "")
                kf = (scores.get(idx) or {}).get("desc", "")
                is_deduct = idx == 11 and kf
                items[idx] = {
                    "desc": (kf if is_deduct else DESCS.get(idx, "")),
                    "score": sc,
                    "material_text": "",
                    "materials": materials[idx],
                    "eval_desc": (kf if is_deduct else "已完成"),
                    "super_score": sc,
                }
            out_dir = os.path.join(OUT_ROOT, "2026年%d月" % month, "安全员")
            os.makedirs(out_dir, exist_ok=True)
            out_name = k.build_filename(PATTERN, YEAR, month, "王雪梅")
            out_path = os.path.join(out_dir, out_name)
            k.generate_doc(TEMPLATE, out_path, "王雪梅", month, items, year=YEAR)
            nmat = sum(len(v) for v in materials.values())
            mat_names = {idx: [os.path.basename(p) for p in v] for idx, v in materials.items() if v}
            ok.append((month, out_path))
            print("[OK] %d月 王雪梅 -> %s，材料 %d 个 (%.1fs)" % (month, os.path.basename(out_path), nmat, time.time() - t0))
            for idx, names in mat_names.items():
                print("     第%d项: %s" % (idx, "、".join(names)))
            log.append({"month": month, "materials": mat_names})
        except Exception as e:
            fail.append((month, str(e)))
            print("[FAIL] %d月 : %s" % (month, e))
    summary = {"total": len(ok) + len(fail), "ok": len(ok), "fail": len(fail),
               "fail_list": [{"month": m, "error": e} for m, e in fail],
               "elapsed_sec": round(time.time() - t_start, 1)}
    with io.open(os.path.join(BASE, "_wxm_mat_summary.json"), "w", encoding="utf-8") as f:
        json.dump({"summary": summary, "log": log}, f, ensure_ascii=False, indent=2)
    print("=" * 60)
    print("完成：成功 %d / 失败 %d，总耗时 %.1fs" % (len(ok), len(fail), summary["elapsed_sec"]))


if __name__ == "__main__":
    main()
