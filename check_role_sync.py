#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""四岗位副本一致性检查（防漂移）。

背景：四个岗位目录各存一份 kaoping_core.py / generate_kaoping.py / tests，
历史教训是「同一修改要手动落到 4 个岗位各一遍」（最近一次提交一次性改了 17 个文件），
漏改/改串即产生静默漂移，日后模板改版或换模板时错格写入。

本脚本逐文件解析每个岗位副本的顶层函数，做「规范化后源码」跨岗位对比：
去掉注释、docstring 与字符串字面量（岗位差异大多藏在字符串里，如表头标签、报错文案），
剩下的代码若在岗位间不一致，即为「分歧函数」。脚本与入库的基线 role_sync_baseline.json
对比，只报告**相对基线的新增分歧**，避免刷屏历史已知差异。

用法（在仓库根目录执行）：
    python check_role_sync.py            # 检查：与基线对比，发现新增漂移则退出码 1
    python check_role_sync.py --init     # 首次使用：把当前差异固化为基线并入库
    python check_role_sync.py --update   # 有意的同步改动后：刷新基线（4 岗位已同步好再刷）
"""
import argparse
import ast
import json
import os
import re
import sys

ROLE_DIRS = [
    "安全员安全生产责任制履职",
    "班组长安全生产责任制履职",
    "作业长安全生产责任制履职",
    "主要负责人安全生产责任制履职",
]
RELEVANT_FILES = [
    "kaoping_core.py",
    "generate_kaoping.py",
    "tests/test_core.py",
    "tests/test_gui.py",
    "test_ole.py",
]
BASELINE_FILE = "role_sync_baseline.json"
HEAD = os.path.dirname(os.path.abspath(__file__))


def _norm_source(src):
    """规范化函数源码：去注释/docstring/字符串字面量，折叠空白。

    岗位差异（表头标签、报错文案、命名模板等）绝大多数藏在字符串里，
    去掉后逻辑代码应当完全一致；若仍不一致即为真正需要关注的分歧。
    """
    src = re.sub(r"#[^\n]*", "", src)
    src = re.sub(r'"""(?:[^"\\]|\\.|"(?!""))*"""', '""', src, flags=re.S)
    src = re.sub(r"'(?:[^'\\\n]|\\.)*'", "''", src)
    src = re.sub(r'"(?:[^"\\\n]|\\.)*"', '""', src)
    return re.sub(r"\s+", " ", src).strip()


def _extract_functions(text):
    """返回 {函数名: 规范化源码}。解析失败（语法异常等）时返回 None 交由调用方报错。"""
    try:
        tree = ast.parse(text)
    except SyntaxError as e:
        return None, e
    result = {}
    for node in tree.body:
        if isinstance(node, ast.FunctionDef):
            result[node.name] = _norm_source(ast.get_source_segment(text, node) or "")
    return result, None


def _read_role_file(role_dir, rel):
    p = os.path.join(HEAD, role_dir, rel)
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return f.read()


def _collect():
    """返回 {(role_dir, rel, func): normalized_source_or_MISSING}。"""
    data = {}
    for role in ROLE_DIRS:
        for rel in RELEVANT_FILES:
            text = _read_role_file(role, rel)
            if text is None:
                data[(role, rel, "<文件缺失>")] = "MISSING"
                continue
            funcs, err = _extract_functions(text)
            if err is not None:
                data[(role, rel, "<解析失败:%s>" % err.lineno)] = "PARSE_ERROR:%s" % err.msg
                continue
            for name, src in funcs.items():
                data[(role, rel, name)] = src
    return data


def _divergences(data):
    """计算当前跨岗位分歧集合：{文件: {函数: {岗位: 状态摘要}}}。"""
    by_key = {}
    for (role, rel, func), src in data.items():
        by_key.setdefault((rel, func), {})[role] = src
    result = []
    for (rel, func), roles in sorted(by_key.items()):
        values = {s for s in roles.values() if not s.startswith("MISSING")}
        # 至少两个岗位都真正存在该函数时才算"函数级"对比（避免把单岗位独有函数当分歧）
        present = [r for r, s in roles.items() if not s.startswith("MISSING")]
        if len(present) >= 2 and len(values) > 1:
            summary = {}
            for role, s in roles.items():
                summary[role] = "缺失" if s.startswith("MISSING") else "不一致"
            result.append((rel, func, summary))
        elif len(present) < 2:
            for role, s in roles.items():
                if s.startswith("MISSING"):
                    result.append((rel, func, {role: "缺失"}))
    return result


def _pretty(report):
    lines = []
    for rel, func, summary in report:
        parts = " / ".join("%s=%s" % (r.replace("安全生产责任制履职", ""), v)
                           for r, v in summary.items())
        lines.append("  %s :: %s  [%s]" % (rel, func, parts))
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="四岗位副本一致性检查（防漂移）")
    ap.add_argument("--init", action="store_true", help="首次使用：生成基线文件")
    ap.add_argument("--update", action="store_true", help="有意的同步改动后：刷新基线")
    args = ap.parse_args()

    data = _collect()
    current = _divergences(data)
    current_keys = set((rel, func) for rel, func, _s in current)

    baseline_path = os.path.join(HEAD, BASELINE_FILE)
    if args.init or args.update:
        payload = {"files": RELEVANT_FILES,
                   "divergences": [{"file": rel, "func": func, "summary": summary}
                                   for rel, func, summary in current]}
        with open(baseline_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print("已%s基线：%s（当前分歧 %d 项）" % ("生成" if args.init else "刷新",
                                            BASELINE_FILE, len(current)))
        return 0

    if not os.path.exists(baseline_path):
        print("未找到基线文件 %s。首次使用请先执行：python check_role_sync.py --init" % BASELINE_FILE)
        return 2

    with open(baseline_path, encoding="utf-8") as f:
        baseline = { (d["file"], d["func"]) for d in json.load(f)["divergences"] }

    new_ones = sorted(current_keys - baseline)
    gone = sorted(baseline - current_keys)
    if not new_ones and not gone:
        print("[OK] 四岗位副本与基线一致，无新增漂移。")
        return 0
    if new_ones:
        report = [c for c in current if (c[0], c[1]) in new_ones]
        print("发现 %d 处新增漂移（相对基线）：" % len(new_ones))
        print(_pretty(report))
        print("若已确认 4 个岗位同步到位，请执行：python check_role_sync.py --update")
    if gone:
        print("基线中的 %d 处分歧已消失（4 岗位已对齐），建议执行 --update 刷新基线："
              % len(gone))
        print("  " + "\n  ".join("%s :: %s" % (f, n) for f, n in sorted(gone)))
    return 1


if __name__ == "__main__":
    sys.exit(main())
