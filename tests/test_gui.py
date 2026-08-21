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
            assert len(app.items_widgets) == 12
            assert hasattr(app, "import_label")        # v1.3.0 读取已生成考评表(.doc)
            assert not hasattr(app, "ref_label")
            assert not hasattr(app, "ref_materials")
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
            ok = {i: {"score": "", "super_score": ""} for i in range(1, 13)}
            assert app._validate_items(ok) is True
            ok[1]["score"] = "25"                     # 超出 ①标准分 20
            assert app._validate_items(ok) is False
            ok[1]["score"] = "abc"                    # 非数字
            assert app._validate_items(ok) is False
            ok[1]["score"] = "15分"                   # 中文“分”后缀合法
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
            root.after(300, root.quit)
            root.mainloop()                           # 等待 worker 线程的 after 回调
            assert app.items_widgets[9].material_paths == [mf], \
                app.items_widgets[9].material_paths
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
        assert "扫雷" in rules[10]
        assert kc.match_materials_file("扫雷/统计.xlsx") == 10
        assert kc.match_materials_file("（炼铁）8月主题工作/x.jpg") == 12
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
        fake_doc = os.path.join(tmp, "孙忠7月.doc")
        open(fake_doc, "w").close()
        fake_mat = _touch(tmp, "材料", "证据.xlsx")
        fake_data = {
            "name": "孙忠", "month": "7", "warnings": ["某警告"],
            "items": {i: {"desc": "", "score": "", "material_text": "",
                          "materials": [], "eval_desc": "", "super_score": ""}
                      for i in range(1, 13)},
        }
        fake_data["items"][1] = {"desc": "未发生事故", "score": "20",
                                 "material_text": "未发生事故", "materials": [],
                                 "eval_desc": "未发生", "super_score": "20"}
        fake_data["items"][11] = {"desc": "未按违章查处", "score": "5",
                                  "material_text": "", "materials": [fake_mat],
                                  "eval_desc": "未按", "super_score": "5"}
        orig = kc.extract_kaoping_doc
        kc.extract_kaoping_doc = lambda p, out_dir=None, progress_cb=None: fake_data
        try:
            root, app = _new_app(tmp)
            try:
                app._apply_kaoping_doc(fake_doc)
                assert app.name_var.get() == "孙忠"
                assert app.month_var.get() == "7"
                it1 = app.items_widgets[0]
                assert it1.desc_var.get() == "未发生事故"
                assert it1.score_var.get() == "20" and it1.super_var.get() == "20"
                it11 = app.items_widgets[10]
                assert it11.score_var.get() == "5"
                assert it11.material_paths == [fake_mat]
                assert "12 项" in app.status_var.get()
            finally:
                root.destroy()
        finally:
            kc.extract_kaoping_doc = orig
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
