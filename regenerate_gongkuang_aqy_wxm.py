# -*- coding: utf-8 -*-
"""安全员 王雪梅（供矿作业区安全员）1-7 月履职履责考评表 —— 无界面批量生成。

复用安全员工具（v1.1.6）的核心逻辑 kaoping_core.py：
  评分以月度履职履责表 xlsx「安全员月度履职评价表」为准；
  自评描述为新拟写（王雪梅无历史考评表），按 12 个考评项 + 供矿作业区场景撰写。

评分口径（来自 xlsx）：
  ①安全绩效20 ②管理体系10 ③~⑩各5 ⑫其它10 全满分；
  ⑪违章管理(20) 1/2 月 10 分、3~7 月 15 分（扣分说明：未完成违章查处）。
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
# Word COM 打不开 OneDrive 路径，复制到本地临时目录
TEMPLATE = os.path.join(tempfile.gettempdir(), "_aqy_kaoping_template.doc")

YEAR = 2026
PATTERN = "安全员安全生产责任制履职清单考评表（{XXX}{X}月）.doc"

# 12 项自评描述（供矿作业区安全员，新拟写）
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
    11: "未完成违章查处",          # 违章管理：按 xlsx 扣分说明
    12: "完成上级交办的其他安全管理工作",
}


def _find_xlsx(month):
    d = os.path.join(BASE, "%d月" % month)
    for fn in sorted(os.listdir(d)):
        if fn.lower().endswith(".xlsx"):
            return os.path.join(d, fn)
    raise FileNotFoundError("第%d月目录未找到履职履责表 xlsx" % month)


def build_items(scores):
    items = {}
    for idx in range(1, k.N_ITEMS + 1):
        sc = (scores.get(idx) or {}).get("score", "")
        kf = (scores.get(idx) or {}).get("desc", "")   # 扣分说明
        is_deduct = idx == 11 and kf
        items[idx] = {
            "desc": (kf if is_deduct else DESCS.get(idx, "")),
            "score": sc,
            "material_text": "",
            "materials": [],
            "eval_desc": (kf if is_deduct else "已完成"),
            "super_score": sc,
        }
    return items


def main():
    pythoncom.CoInitialize()
    shutil.copy2(TEMPLATE_SRC, TEMPLATE)
    log = []
    ok, fail = [], []
    t_start = time.time()
    for month in range(1, 8):
        xlsx = _find_xlsx(month)
        t0 = time.time()
        try:
            scores = k.read_eval_scores(xlsx, "王雪梅")["items"]
            items = build_items(scores)
            out_dir = os.path.join(OUT_ROOT, "2026年%d月" % month, "安全员")
            os.makedirs(out_dir, exist_ok=True)
            out_name = k.build_filename(PATTERN, YEAR, month, "王雪梅")
            out_path = os.path.join(out_dir, out_name)
            k.generate_doc(TEMPLATE, out_path, "王雪梅", month, items, year=YEAR)
            ok.append((month, out_path))
            msg = "[OK] %d月 王雪梅 -> %s (%.1fs)" % (month, os.path.basename(out_path), time.time() - t0)
            print(msg)
            log.append(msg)
        except Exception as e:
            fail.append((month, str(e)))
            msg = "[FAIL] %d月 王雪梅 : %s" % (month, e)
            print(msg)
            log.append(msg)

    summary = {
        "total": len(ok) + len(fail), "ok": len(ok), "fail": len(fail),
        "ok_list": [{"month": m, "path": p} for m, p in ok],
        "fail_list": [{"month": m, "error": e} for m, e in fail],
        "elapsed_sec": round(time.time() - t_start, 1),
    }
    with io.open(os.path.join(OUT_ROOT, "_王雪梅_regenerate_summary.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    with io.open(os.path.join(OUT_ROOT, "_王雪梅_regenerate_log.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(log) + "\n")
    print("=" * 60)
    print("完成：成功 %d / 失败 %d，总耗时 %.1fs" % (len(ok), len(fail), summary["elapsed_sec"]))


if __name__ == "__main__":
    main()
