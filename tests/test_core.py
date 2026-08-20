# -*- coding: utf-8 -*-
"""kaoping_core 纯函数自动化测试（无需第三方测试框架，python tests/test_core.py 即可运行）"""
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import kaoping_core as kc  # noqa: E402


def test_build_filename():
    # 默认模板
    assert kc.build_filename(kc.DEFAULT_NAME_PATTERN, 2026, 8, "张三") == \
        "安全员安全生产责任制履职清单考评表(8月张三).doc"
    # 全角/自定义模板
    assert kc.build_filename("{XXX}（{Y}年{X}月）", 2026, 12, "李四") == \
        "李四（2026年12月）.doc"
    # 非法字符替换
    assert kc.build_filename("{XXX}", 2026, 8, '王五/李六*') == "王五_李六_.doc"
    # Windows 保留设备名
    assert kc.build_filename("CON", 2026, 8, "x").startswith("_")
    assert kc.build_filename("NUL.{XXX}", 2026, 8, "x").startswith("_")
    # 结尾点/空格消毒
    assert kc.build_filename("{XXX}.doc", 2026, 8, "赵六. ") == "赵六.doc"


def test_replace_slot():
    body = "考评对象（安全员）：  孙忠      管理者姓名：x评价月份：                 评价人员签字"
    b = kc._replace_slot(body, "考评对象（安全员）：", "管理者姓名", "测试员")
    assert "：  测试员      管理者姓名" in b
    b = kc._replace_slot(b, "评价月份：", "评价人员签字", "8月")
    seg = b[b.find("评价月份：") + 5:b.find("评价人员签字")]
    assert seg.strip() == "8月", f"月份区间异常: {seg!r}"
    # 值居中：前后都有空白
    assert seg.startswith(" ") and seg.endswith(" ")


def test_sum_scores():
    items = {
        1: {"score": "20", "super_score": "18分"},
        2: {"score": "10.5", "super_score": ""},
        3: {"score": "abc", "super_score": "5"},
    }
    assert kc._sum_scores(items, "score") == "30.5"
    assert kc._sum_scores(items, "super_score") == "23"
    assert kc._sum_scores({}, "score") == ""


def test_is_image():
    assert kc.is_image("a.JPG")
    assert kc.is_image("a.png")
    assert not kc.is_image("a.docx")
    assert not kc.is_image("a.xlsx")


def test_match_materials_file():
    assert kc.match_materials_file("扫雷/2026年扫雷统计表.xlsx") == 10
    assert kc.match_materials_file("综合大检查/安全综合检查通报.doc") == 5
    assert kc.match_materials_file("（炼铁）8月主题工作/1.高危人群清查/xxx.jpg") == 12
    assert kc.match_materials_file("应急演练/皮带机演练方案.doc") == 6
    assert kc.match_materials_file("危险源/危险源辨识记录.xlsx") == 9
    assert kc.match_materials_file("安全培训课件.pdf") == 3
    assert kc.match_materials_file("违章查处记录.xls") == 11
    # 无匹配
    assert kc.match_materials_file("高温天气预警/8.3.jpg") is None
    assert kc.match_materials_file("新员工入职/电动车审批蒋千林.docx") is None


def test_scan_materials_folder():
    tmp = tempfile.mkdtemp()
    try:
        os.makedirs(os.path.join(tmp, "扫雷"))
        os.makedirs(os.path.join(tmp, "应急演练"))
        # 匹配文件
        open(os.path.join(tmp, "扫雷", "扫雷统计表.xlsx"), "w").close()
        open(os.path.join(tmp, "应急演练", "演练方案.doc"), "w").close()
        # 不匹配文件
        open(os.path.join(tmp, "高温预警.jpg"), "w").close()
        # 应被跳过的考评表文件
        open(os.path.join(tmp, "安全员安全生产责任制履职清单考评表（张三8月）.doc"), "w").close()
        # 应被跳过的临时文件
        open(os.path.join(tmp, "~$临时.docx"), "w").close()

        result, unmatched = kc.scan_materials_folder(tmp)
        assert len(result[10]) == 1 and result[10][0].endswith("扫雷统计表.xlsx")
        assert len(result[6]) == 1 and result[6][0].endswith("演练方案.doc")
        assert len(unmatched) == 1, f"未匹配应为1个，实际 {len(unmatched)}"
        assert unmatched[0].endswith("高温预警.jpg")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failed = 0
    for t in tests:
        try:
            t()
            print(f"[PASS] {t.__name__}")
        except AssertionError as e:
            failed += 1
            print(f"[FAIL] {t.__name__}: {e}")
        except Exception as e:
            failed += 1
            print(f"[ERROR] {t.__name__}: {type(e).__name__}: {e}")
    print(f"\n共 {len(tests)} 项测试，失败 {failed} 项")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
