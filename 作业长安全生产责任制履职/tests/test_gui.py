# -*- coding: utf-8 -*-
"""GUI 自动化冒烟测试（无第三方测试框架，python tests/test_gui.py 即可运行）。

覆盖：ImageItem 材料增删/预览绑定、App 构建布局、得分校验、后台扫描异步回填、
权重配置、清空全部、生成完成状态。不启动 Word、不弹真实对话框。
"""
import os
import shutil
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import kaoping_core as kc      # noqa: E402
import generate_kaoping as gk  # noqa: E402
from tkinterdnd2 import TkinterDnD  # noqa: E402


class _FakeBox:
    """替代 messagebox：记录调用并返回 False，避免弹窗阻塞测试。"""

    def __init__(self):
        self.calls = []

    def _rec(self, *args, **kw):
        self.calls.append((args, kw))
        return False

    askyesno = _rec
    askokcancel = _rec
    askretrycancel = _rec
    showinfo = _rec
    showwarning = _rec
    showerror = _rec


def _patch_env(tmpdir):
    """屏蔽真实对话框/文件对话框/配置读写，保证测试可重复且不污染。"""
    gk.messagebox = _FakeBox()
    gk.filedialog.askdirectory = lambda **kw: tmpdir
    gk.filedialog.askopenfilenames = lambda **kw: []
    gk.filedialog.askopenfilename = lambda **kw: ""
    gk._load_config = lambda: {"out_dir": tmpdir, "pattern": kc.DEFAULT_NAME_PATTERN}
    gk._save_config = lambda cfg: True
    kc.find_template = lambda *a, **kw: os.path.join(tmpdir, "不存在模板.doc")


def _new_app(tmpdir):
    _patch_env(tmpdir)
    root = TkinterDnD.Tk()
    root.withdraw()
    app = gk.App(root)
    return root, app


def _touch(*parts):
    os.makedirs(os.path.dirname(os.path.join(*parts)), exist_ok=True)
    open(os.path.join(*parts), "w").close()
    return os.path.join(*parts)


def test_imageitem_materials():
    tmp = tempfile.mkdtemp()
    try:
        root, app = _new_app(tmp)
        try:
            it = app.items_widgets[0]
            f1 = _touch(tmp, "材料a.docx")
            f2 = _touch(tmp, "材料b.pdf")
            it.add_materials([f1, f1, f2])          # 重复文件应去重
            assert it.material_paths == [f1, f2], it.material_paths
            it._remove_material(f1)
            assert it.material_paths == [f2]
            for i in range(20):                     # 超过上限后不再追加
                it.add_materials([_touch(tmp, f"填充{i}.txt")])
            assert len(it.material_paths) <= it.MAX_MATERIALS
            opened = []                             # 双击打开回调（monkeypatch）
            orig = gk.os.startfile
            gk.os.startfile = lambda p: opened.append(p)
            try:
                it._open_material(f2)
            finally:
                gk.os.startfile = orig
            assert opened == [f2]
            it.clear()
            assert it.material_paths == []
        finally:
            root.destroy()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_imageitem_single_row():
    """作业长无双行项：第 1 项 sub_max 为 None、to_item 不含 sub、clear 单清。"""
    tmp = tempfile.mkdtemp()
    try:
        root, app = _new_app(tmp)
        try:
            it = app.items_widgets[0]                       # 第 1 项 安全绩效（单行项）
            assert it.sub_max is None
            assert it.material_paths == []
            f1 = _touch(tmp, "材料1.doc")
            it.add_materials([f1])
            assert it.material_paths == [f1]
            item = it.to_item()
            assert "sub" not in item
            it.clear()
            assert it.desc_var.get() == ""
            assert it.score_var.get() == "20"               # 恢复标准分
            assert it.material_paths == []
        finally:
            root.destroy()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_material_preview_bindings():
    tmp = tempfile.mkdtemp()
    try:
        root, app = _new_app(tmp)
        try:
            it = app.items_widgets[0]
            it.add_materials([_touch(tmp, "photo.png")])   # 空图触发降级分支，仍生成预览
            labs = it.thumb_frame.winfo_children()
            assert len(labs) == 1, "预览标签未生成"
            lab = labs[0]
            for seq in ("<Double-Button-1>", "<Button-3>", "<Enter>", "<Leave>"):
                assert lab.bind(seq), f"预览未绑定 {seq}"
        finally:
            root.destroy()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def test_app_build_layout():
    tmp = tempfile.mkdtemp()
    try:
        root, app = _new_app(tmp)
        try:
            assert len(app.items_widgets) == 13
            assert hasattr(app, "import_label")        # v1.1.2 读取已生成考评表(.doc)
            assert not hasattr(app, "ref_label")
            assert not hasattr(app, "ref_materials")
            # 作业长模板无双行项：第 1 项不渲染 sub 控件
            it1 = app.items_widgets[0]
            assert it1.sub_max is None
            assert not hasattr(it1, "sub_desc_var")
            bottom = app.btn_generate.master           # “一键生成”最先布局，缩小不被遮挡
            assert bottom.winfo_children()[0] is app.btn_generate
            assert app.canvas.bind("<MouseWheel>")     # 滚轮绑定
            btns = [c.cget("text") for c in app.mat_folder_label.master.winfo_children()
                    if c.winfo_class() == "TButton"]
            assert "编辑权重" in btns, btns
            btns2 = [c.cget("text") for c in app.import_label.master.winfo_children()
                     if c.winfo_class() == "TButton"]
            assert "读取考评表" in btns2, btns2

        finally:
            root.destroy()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_validate_items():
    tmp = tempfile.mkdtemp()
    try:
        root, app = _new_app(tmp)
        try:
            ok = {i: {"score": "", "super_score": ""} for i in range(1, 14)}
            assert app._validate_items(ok) is True
            ok[1]["score"] = "25"                     # 超出 ①标准分 20
            assert app._validate_items(ok) is False
            ok[1]["score"] = "abc"                    # 非数字
            assert app._validate_items(ok) is False
            ok[1]["score"] = "15分"                   # 中文“分”后缀合法（≤20）
            assert app._validate_items(ok) is True
            ok[1]["score"] = "12"                      # 未超标准分 20 → 合法
            assert app._validate_items(ok) is True
        finally:
            root.destroy()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_scan_async_fill():
    """年度根目录定位当月 + 后台线程扫描 + 异步回填材料。"""
    tmp = tempfile.mkdtemp()
    try:
        root, app = _new_app(tmp)
        try:
            month_dir = os.path.join(tmp, "2026.08")
            mf = _touch(month_dir, "扫雷", "扫雷记录.xlsx")
            app.year_var.set("2026")
            app.month_var.set("8")
            app._load_material_folder(tmp)            # 选年度根目录 → 自动定位 2026.08
            root.after(600, root.quit)                # 等待 worker 线程的 after 回调
            root.mainloop()
            assert app.items_widgets[1].material_paths == [mf], \
                app.items_widgets[1].material_paths
        finally:
            root.destroy()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_material_rules_and_weights():
    tmp = tempfile.mkdtemp()
    try:
        path = kc.ensure_material_rules_file(tmp)
        assert os.path.exists(path)
        rules = kc.load_material_rules(path)
        assert "扫雷" in rules[2]
        assert kc.match_materials_file("扫雷/统计.xlsx") == 2
        assert kc.match_materials_file("六个一/安全专题会纪要.doc") == 13
        bad = os.path.join(tmp, "bad.json")           # 非法配置回退内置
        with open(bad, "w", encoding="utf-8") as f:
            f.write("{oops")
        assert kc.load_material_rules(bad) == kc.ITEM_MATCH_RULES
        assert kc.find_month_subfolder(tmp, 2026, 8) == tmp   # 无当月子文件夹原样返回
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_clear_all():
    tmp = tempfile.mkdtemp()
    try:
        root, app = _new_app(tmp)
        try:
            it = app.items_widgets[0]
            it.add_materials([_touch(tmp, "m.docx")])
            it.desc_var.set("测试")
            it.score_var.set("18")
            app._clear_all()
            assert it.material_paths == []
            assert it.desc_var.get() == ""
            assert it.score_var.get() == str(it.std_score)
        finally:
            root.destroy()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_generate_done_status():
    tmp = tempfile.mkdtemp()
    try:
        root, app = _new_app(tmp)
        try:
            app._gen_started = time.time()
            app._on_generate_done(os.path.join(tmp, "out.doc"), tmp)
            assert "生成完成" in app.status_var.get()
            assert "秒" in app.status_var.get()
        finally:
            root.destroy()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_apply_kaoping_doc_fill():
    """读取已生成考评表(.doc) 后回填姓名/月份/评分评价/材料（mock 核心提取）。"""
    tmp = tempfile.mkdtemp()
    try:
        fake_doc = os.path.join(tmp, "廖辉7月.doc")
        open(fake_doc, "w").close()
        fake_mat = _touch(tmp, "材料", "证据.xlsx")
        fake_data = {
            "name": "廖辉", "month": "7", "warnings": ["某警告"],
            "items": {i: {"desc": "", "score": "", "material_text": "",
                          "materials": [], "eval_desc": "", "super_score": ""}
                      for i in range(1, 14)},
        }
        fake_data["items"][1] = {"desc": "无事故", "score": "20",
                                 "material_text": "及时传达事故通报", "materials": [],
                                 "eval_desc": "本月无事故", "super_score": "20"}
        fake_data["items"][9] = {"desc": "未按违章查处", "score": "15",
                                 "material_text": "", "materials": [fake_mat],
                                 "eval_desc": "未按", "super_score": "15"}
        orig = kc.extract_kaoping_doc
        kc.extract_kaoping_doc = lambda p, out_dir=None, progress_cb=None: fake_data
        try:
            root, app = _new_app(tmp)
            try:
                app._apply_kaoping_doc(fake_doc)
                assert app.name_var.get() == "廖辉"
                assert app.month_var.get() == "7"
                it1 = app.items_widgets[0]
                assert it1.desc_var.get() == "无事故"
                assert it1.score_var.get() == "20" and it1.super_var.get() == "20"
                it9 = app.items_widgets[8]
                assert it9.score_var.get() == "15"
                assert it9.material_paths == [fake_mat]
                assert "13 项" in app.status_var.get()
            finally:
                root.destroy()
        finally:
            kc.extract_kaoping_doc = orig
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_eval_scores_fill():
    """读取履职履责表(xlsx)：姓名列表写入下拉、各项评分/描述回填、状态栏显示总分。"""
    tmp = tempfile.mkdtemp()
    try:
        fake = os.path.join(tmp, "履职履责.xlsx")
        open(fake, "w").close()
        orig_list, orig_read = kc.list_eval_names, kc.read_eval_scores
        kc.list_eval_names = lambda p: ["廖辉", "王旺"]
        kc.read_eval_scores = lambda p, name: {
            "name": name, "total": "93",
            "items": {i: {"score": "5", "desc": "描述%d" % i} for i in range(1, 14)},
        }
        try:
            root, app = _new_app(tmp)
            try:
                app._pick_eval_name = lambda names: "王旺"     # 当前姓名为空 → 走选择框
                app._read_eval_scores(fake)
                assert app.name_var.get() == "王旺"
                assert list(app.name_combo["values"]) == ["廖辉", "王旺"]
                assert app.items_widgets[0].score_var.get() == "5"
                assert app.items_widgets[0].desc_var.get() == ""
                assert app.items_widgets[12].score_var.get() == "5"
                assert "93" in app.status_var.get()
            finally:
                root.destroy()
        finally:
            kc.list_eval_names, kc.read_eval_scores = orig_list, orig_read
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_auto_fill_super():
    """自评得分已填的考评项：自动补齐上级评分(=自评得分)与评语(已完成)；已填不覆盖；未打分不生成。"""
    items = {
        1: {"score": "10", "super_score": "", "eval_desc": "", "materials": []},
        2: {"score": "5", "super_score": "5分", "eval_desc": "人工已填", "materials": []},
        3: {"score": "", "super_score": "", "eval_desc": "", "materials": []},
        4: {"score": "8", "super_score": "", "eval_desc": "", "materials": [],
            "sub": {"score": "10", "super_score": "", "eval_desc": "", "materials": []}},
    }
    gk._auto_fill_super(items)
    assert items[1]["super_score"] == "10" and items[1]["eval_desc"] == "已完成"
    assert items[2]["super_score"] == "5分" and items[2]["eval_desc"] == "人工已填"   # 已填不覆盖
    assert items[3]["super_score"] == "" and items[3]["eval_desc"] == ""            # 未打分不生成
    assert items[4]["super_score"] == "8" and items[4]["eval_desc"] == "已完成"     # 主行
    assert items[4]["sub"]["super_score"] == "10" and items[4]["sub"]["eval_desc"] == "已完成"  # 双行sub



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
