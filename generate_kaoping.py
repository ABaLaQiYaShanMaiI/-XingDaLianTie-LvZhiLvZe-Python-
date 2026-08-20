#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""安全员安全生产责任制履职清单考评表 自动生成工具 GUI v1.1.0"""
import os
import sys
import json
import threading
import ctypes
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image, ImageTk
import pythoncom

import kaoping_core as kc

VERSION = "1.1.0"

ITEM_LABELS = [
    (1, "安全绩效", 20),
    (2, "管理体系", 10),
    (3, "教育培训", 5),
    (4, "会议活动", 5),
    (5, "检查改进", 5),
    (6, "应急演习", 5),
    (7, "监督执行", 5),
    (8, "安全能力", 5),
    (9, "危险源", 5),
    (10, "隐患排查", 5),
    (11, "违章管理", 20),
    (12, "其它", 10),
]

# 打包(exe)后 __file__ 指向 _MEIPASS 临时目录，配置文件/默认输出目录需以 exe 目录为准
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "kaoping_config.json")
THUMB_SIZE = (90, 90)


def _load_config():
    cfg = {"out_dir": BASE_DIR, "pattern": kc.DEFAULT_NAME_PATTERN}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg.update(json.load(f))
    except Exception:
        pass
    return cfg


def _save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


class ImageItem:
    """单个考评项的录入控件：自评描述/得分/材料文字/材料拖拽/评价描述/上级评分"""

    MAX_MATERIALS = 15

    def __init__(self, parent, root, index, title, std_score):
        self.root = root
        self.index = index
        self.title = title
        self.std_score = std_score
        self.material_paths = []
        self._thumb_refs = []

        self.frame = ttk.LabelFrame(parent, text=f"  {index}、{title}（标准分 {std_score}分） ", padding=6)
        self.frame.pack(fill=tk.X, pady=4)

        grid = ttk.Frame(self.frame)
        grid.pack(fill=tk.X)

        # 第一行：自评描述
        ttk.Label(grid, text="自评描述:").grid(row=0, column=0, sticky="e", padx=2, pady=2)
        self.desc_var = tk.StringVar()
        self.desc_entry = ttk.Entry(grid, textvariable=self.desc_var, width=45)
        self.desc_entry.grid(row=0, column=1, sticky="we", padx=2, pady=2)
        ttk.Label(grid, text="自评得分:").grid(row=0, column=2, sticky="e", padx=(8, 2))
        self.score_var = tk.StringVar(value=str(std_score))
        self.score_entry = ttk.Entry(grid, textvariable=self.score_var, width=6)
        self.score_entry.grid(row=0, column=3, padx=2, pady=2)
        ttk.Label(grid, text="材料说明:").grid(row=0, column=4, sticky="e", padx=(8, 2))
        self.mat_var = tk.StringVar()
        self.mat_entry = ttk.Entry(grid, textvariable=self.mat_var, width=20)
        self.mat_entry.grid(row=0, column=5, sticky="we", padx=2, pady=2)

        # 第二行：评价描述/上级评分
        ttk.Label(grid, text="评价描述:").grid(row=1, column=0, sticky="e", padx=2, pady=2)
        self.eval_var = tk.StringVar()
        self.eval_entry = ttk.Entry(grid, textvariable=self.eval_var, width=45)
        self.eval_entry.grid(row=1, column=1, sticky="we", padx=2, pady=2)
        ttk.Label(grid, text="上级评分:").grid(row=1, column=2, sticky="e", padx=(8, 2))
        self.super_var = tk.StringVar()
        self.super_entry = ttk.Entry(grid, textvariable=self.super_var, width=6)
        self.super_entry.grid(row=1, column=3, padx=2, pady=2)

        # 第三行：材料拖拽区（图片/文档/表格均可）
        img_row = ttk.Frame(self.frame)
        img_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(img_row, text="支撑材料:").pack(side=tk.LEFT, padx=(2, 4))
        self.drop_label = tk.Label(
            img_row, text=f"拖入材料到此（图片/文档/表格等，最多{self.MAX_MATERIALS}个）或点击选择",
            bg="#E8F0FE", fg="#666", relief=tk.GROOVE,
            height=2, cursor="hand2")
        self.drop_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.drop_label.bind("<Button-1>", lambda e: self._choose_materials())
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self._on_drop)
        self.btn_add = ttk.Button(img_row, text="选择材料", command=self._choose_materials)
        self.btn_add.pack(side=tk.LEFT, padx=2)
        self.btn_clear = ttk.Button(img_row, text="清空材料", command=self._clear_materials)
        self.btn_clear.pack(side=tk.LEFT, padx=2)

        # 材料预览区（图片缩略图 / 文件名称）
        self.thumb_frame = ttk.Frame(self.frame)
        self.thumb_frame.pack(fill=tk.X, pady=(4, 0))

        grid.columnconfigure(1, weight=3)
        grid.columnconfigure(5, weight=2)

    def _on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        self._add_materials(files)

    def _choose_materials(self):
        files = filedialog.askopenfilenames(
            title=f"选择第{self.index}项支撑材料",
            filetypes=[("所有文件", "*.*"),
                       ("图片文件", "*.jpg;*.jpeg;*.png;*.bmp;*.gif"),
                       ("Word 文档", "*.doc;*.docx"),
                       ("Excel 表格", "*.xls;*.xlsx"),
                       ("PDF", "*.pdf")])
        if files:
            self._add_materials(list(files))

    def _add_materials(self, files):
        for p in files:
            if len(self.material_paths) >= self.MAX_MATERIALS:
                messagebox.showwarning("提示", f"每个考评项最多 {self.MAX_MATERIALS} 个材料")
                break
            if not os.path.exists(p):
                continue
            if p not in self.material_paths:
                self.material_paths.append(p)
        self._refresh_previews()

    def _clear_materials(self):
        self.material_paths = []
        self._refresh_previews()

    def _remove_material(self, path):
        if path in self.material_paths:
            self.material_paths.remove(path)
            self._refresh_previews()

    def _refresh_previews(self):
        for w in self.thumb_frame.winfo_children():
            w.destroy()
        self._thumb_refs = []
        if not self.material_paths:
            return
        cols = 6
        for i, p in enumerate(self.material_paths):
            r, c = i // cols, i % cols
            short = os.path.basename(p)
            if len(short) > 12:
                short = short[:11] + "…"
            if kc.is_image(p):
                try:
                    with Image.open(p) as im:
                        im.thumbnail(THUMB_SIZE)
                        photo = ImageTk.PhotoImage(im)
                    self._thumb_refs.append(photo)
                    lab = tk.Label(self.thumb_frame, image=photo, text=short,
                                   compound=tk.TOP, relief=tk.RIDGE,
                                   cursor="hand2", fg="#333")
                except Exception:
                    lab = tk.Label(self.thumb_frame, text="🖼️\n" + short,
                                   relief=tk.RIDGE, bg="#f0f4ff", fg="#333",
                                   cursor="hand2", justify=tk.CENTER)
            else:
                lab = tk.Label(self.thumb_frame, text="📄\n" + short,
                               relief=tk.RIDGE, bg="#FFF8E1", fg="#8a5a00",
                               cursor="hand2", justify=tk.CENTER)
            lab.grid(row=r, column=c, padx=3, pady=2, sticky="n")
            full = os.path.basename(p)
            lab.bind("<Button-1>", lambda e, path=p: self._remove_material(path))
            lab.bind("<Enter>", lambda e, ww=lab, full=full: ww.config(text=full))
            lab.bind("<Leave>", lambda e, ww=lab, s=short: ww.config(text=s))

    def to_item(self):
        return {
            "desc": self.desc_var.get().strip(),
            "score": self.score_var.get().strip(),
            "material_text": self.mat_var.get().strip(),
            "materials": list(self.material_paths),
            "eval_desc": self.eval_var.get().strip(),
            "super_score": self.super_var.get().strip(),
        }

    def fill_from_xlsx(self, item):
        """xlsx 自动填充：得分 + 扣分说明（作为自评描述）"""
        self.score_var.set(item.get("score", ""))
        self.desc_var.set(item.get("desc", ""))

    def clear(self):
        self.desc_var.set("")
        self.score_var.set(str(self.std_score))
        self.mat_var.set("")
        self.eval_var.set("")
        self.super_var.set("")
        self._clear_materials()


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = _load_config()
        self.persons = []
        self.items_widgets = []
        root.title(f"安全员安全生产责任制履职清单考评表自动生成工具 v{VERSION}")
        root.geometry("1180x860")
        root.minsize(1000, 700)
        self._build_ui()
        self._load_template_path()

    # ============ 界面构建 ============
    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)

        # ---- 顶部：基本信息 ----
        top = ttk.LabelFrame(outer, text=" 基本信息 ", padding=8)
        top.pack(fill=tk.X, pady=(0, 6))

        r1 = ttk.Frame(top); r1.pack(fill=tk.X, pady=2)
        ttk.Label(r1, text="年份:").pack(side=tk.LEFT)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Spinbox(r1, from_=2020, to=2040, textvariable=self.year_var, width=6).pack(side=tk.LEFT, padx=(2, 12))
        ttk.Label(r1, text="月份:").pack(side=tk.LEFT)
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        ttk.Spinbox(r1, from_=1, to=12, textvariable=self.month_var, width=5).pack(side=tk.LEFT, padx=(2, 12))
        ttk.Label(r1, text="姓名:").pack(side=tk.LEFT)
        self.name_var = tk.StringVar()
        self.name_combo = ttk.Combobox(r1, textvariable=self.name_var, width=12)
        self.name_combo.pack(side=tk.LEFT, padx=(2, 12))
        self.name_combo.bind("<<ComboboxSelected>>", self._on_name_selected)
        ttk.Label(r1, text="履职履责表:").pack(side=tk.LEFT)
        self.xlsx_label = tk.Label(
            r1, text="拖入当月履职履责表.xlsx 到此自动填充 或 点击选择",
            bg="#DFF0D8", fg="#2c6e49", relief=tk.GROOVE,
            height=1, cursor="hand2")
        self.xlsx_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.xlsx_label.bind("<Button-1>", lambda e: self._load_xlsx())
        self.xlsx_label.drop_target_register(DND_FILES)
        self.xlsx_label.dnd_bind("<<Drop>>", self._drop_xlsx)

        r2 = ttk.Frame(top); r2.pack(fill=tk.X, pady=2)
        ttk.Label(r2, text="命名模板:").pack(side=tk.LEFT)
        self.pattern_var = tk.StringVar(value=self.cfg.get("pattern", kc.DEFAULT_NAME_PATTERN))
        ttk.Entry(r2, textvariable=self.pattern_var, width=55).pack(side=tk.LEFT, padx=(2, 12))
        ttk.Label(r2, text="(占位符: {Y}年 {X}月 {XXX}姓名)").pack(side=tk.LEFT)

        r3 = ttk.Frame(top); r3.pack(fill=tk.X, pady=2)
        ttk.Label(r3, text="模板文件:").pack(side=tk.LEFT)
        self.tpl_var = tk.StringVar()
        self.tpl_label = tk.Label(r3, text="自动定位", bg="#E8F0FE", fg="#555",
                                  relief=tk.GROOVE, cursor="hand2")
        self.tpl_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.tpl_label.bind("<Button-1>", lambda e: self._sel_template())
        self.tpl_label.drop_target_register(DND_FILES)
        self.tpl_label.dnd_bind("<<Drop>>", self._drop_template)
        ttk.Button(r3, text="选择", width=8, command=self._sel_template).pack(side=tk.LEFT, padx=2)
        ttk.Label(r3, text="输出目录:").pack(side=tk.LEFT, padx=(12, 2))
        self.out_var = tk.StringVar(value=self.cfg.get("out_dir", BASE_DIR))
        self.out_label = tk.Label(r3, text=self.out_var.get(), bg="#FCE4D6", fg="#555",
                                  relief=tk.GROOVE, cursor="hand2")
        self.out_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.out_label.bind("<Button-1>", lambda e: self._sel_outdir())
        ttk.Button(r3, text="选择", width=8, command=self._sel_outdir).pack(side=tk.LEFT, padx=2)

        # ---- 第4行：支撑材料文件夹自动匹配 ----
        r4 = ttk.Frame(top); r4.pack(fill=tk.X, pady=2)
        ttk.Label(r4, text="支撑材料文件夹:").pack(side=tk.LEFT)
        self.mat_folder_label = tk.Label(
            r4, text="选择月份文件夹，按文件名关键词自动匹配填入12个考评项",
            bg="#E8F0FE", fg="#555", relief=tk.GROOVE, cursor="hand2")
        self.mat_folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.mat_folder_label.bind("<Button-1>", lambda e: self._load_material_folder())
        self.mat_folder_label.drop_target_register(DND_FILES)
        self.mat_folder_label.dnd_bind("<<Drop>>", self._drop_material_folder)
        ttk.Button(r4, text="选择", width=8, command=self._load_material_folder).pack(side=tk.LEFT, padx=2)
        ttk.Label(r4, text="(未匹配文件会提示，可手动拖入)").pack(side=tk.LEFT, padx=(8, 0))

        # ---- 中部：12 项滚动列表 ----
        mid_lf = ttk.LabelFrame(outer, text=" 各考评项填写（可拖拽图片/文档/表格材料） ", padding=4)
        mid_lf.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(mid_lf, highlightthickness=0)
        vsb = ttk.Scrollbar(mid_lf, orient=tk.VERTICAL, command=canvas.yview)
        self.list_frame = ttk.Frame(canvas)
        self.list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas = canvas
        canvas.bind("<MouseWheel>", self._on_mousewheel)

        for idx, title, std in ITEM_LABELS:
            w = ImageItem(self.list_frame, self.root, idx, title, std)
            self.items_widgets.append(w)

        # ---- 底部工具栏 ----
        bottom = ttk.Frame(outer)
        bottom.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(bottom, text="读取履职履责表.xlsx 自动填充",
                   command=self._load_xlsx).pack(side=tk.LEFT, padx=2)
        ttk.Button(bottom, text="清空全部", command=self._clear_all).pack(side=tk.LEFT, padx=2)
        self.prog = ttk.Progressbar(bottom, length=260, mode="determinate")
        self.prog.pack(side=tk.LEFT, padx=10)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bottom, textvariable=self.status_var, foreground="#555").pack(side=tk.LEFT, padx=8)
        self.btn_generate = ttk.Button(bottom, text="一键生成", command=self._generate)
        self.btn_generate.pack(side=tk.RIGHT, padx=2)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    # ============ 模板 / 输出目录 ============
    def _load_template_path(self):
        try:
            tpl = kc.find_template()
            self.tpl_var.set(tpl)
            self.tpl_label.config(text=tpl, fg="#1a7f37")
        except FileNotFoundError as e:
            self.tpl_var.set("")
            self.tpl_label.config(text=str(e), fg="#c0392b")

    def _sel_template(self):
        p = filedialog.askopenfilename(title="选择考评表模板(.doc)",
                                       filetypes=[("Word 文档", "*.doc;*.docx"), ("所有文件", "*.*")])
        if p:
            self.tpl_var.set(p)
            self.tpl_label.config(text=p, fg="#1a7f37")

    def _drop_template(self, event):
        files = self.root.tk.splitlist(event.data)
        if files:
            self.tpl_var.set(files[0])
            self.tpl_label.config(text=files[0], fg="#1a7f37")

    def _sel_outdir(self):
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            self.out_var.set(p)
            self.out_label.config(text=p)

    # ============ 支撑材料文件夹自动匹配 ============
    def _drop_material_folder(self, event):
        for p in self.root.tk.splitlist(event.data):
            if os.path.isdir(p):
                self._load_material_folder(p)
                return

    def _load_material_folder(self, path=None):
        if not path:
            path = filedialog.askdirectory(
                title="选择支撑材料文件夹（按文件名关键词自动匹配到12个考评项）")
        if not path:
            return
        self.status_var.set("正在扫描支撑材料文件夹...")
        self.root.update()
        try:
            result, unmatched = kc.scan_materials_folder(path)
        except Exception as e:
            messagebox.showerror("扫描失败", str(e))
            return
        filled = 0
        for w in self.items_widgets:
            paths = result.get(w.index, [])
            if paths:
                before = len(w.material_paths)
                w._add_materials(paths)
                filled += len(w.material_paths) - before
        self.mat_folder_label.config(text=os.path.basename(path), fg="#2c6e49")
        detail = "，".join(f"{i}项{len(result[i])}" for i in range(1, 13) if result[i])
        if len(detail) > 80:
            detail = detail[:77] + "…"
        suffix = (f"，{len(unmatched)} 个未匹配（可手动拖入）" if unmatched
                  else "，全部匹配")
        self.status_var.set(f"已自动填入 {filled} 个支撑材料{suffix}：{detail}")
        # 未匹配文件过多时仅状态栏提示，避免弹窗刷屏
        if 0 < len(unmatched) <= 200:
            names = "\n".join(os.path.basename(u) for u in unmatched[:40])
            if len(unmatched) > 40:
                names += f"\n……共 {len(unmatched)} 个未匹配"
            messagebox.showinfo("未匹配文件",
                                "以下文件未匹配到任何考评项，请手动拖入对应项：\n\n" + names)

    # ============ xlsx 自动填充 ============
    def _drop_xlsx(self, event):
        files = self.root.tk.splitlist(event.data)
        if files:
            self._load_xlsx(files[0])

    def _load_xlsx(self, path=None):
        if not path:
            path = filedialog.askopenfilename(
                title="选择履职履责表.xlsx",
                filetypes=[("Excel 文件", "*.xlsx;*.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            self.persons = kc.read_xlsx_scores(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        names = [x["name"] for x in self.persons]
        self.name_combo["values"] = names
        self.xlsx_label.config(text=os.path.basename(path), fg="#2c6e49")
        self.status_var.set(f"已读取 {len(names)} 人：{'、'.join(names)}")
        if names:
            self.name_var.set(names[0])
            self._apply_person(names[0])

    def _on_name_selected(self, _event=None):
        name = self.name_var.get().strip()
        if name:
            self._apply_person(name)

    def _apply_person(self, name):
        person = next((x for x in self.persons if x["name"] == name), None)
        if not person:
            return
        for w in self.items_widgets:
            w.fill_from_xlsx(person["items"].get(w.index, {}))
        self.status_var.set(f"已按「{name}」填充 12 项自评得分与扣分说明")

    def _clear_all(self):
        for w in self.items_widgets:
            w.clear()
        self.status_var.set("已清空全部填写内容")

    # ============ 一键生成 ============
    def _generate(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "请填写姓名")
            return
        try:
            year = int(self.year_var.get().strip())
            month = int(self.month_var.get().strip())
        except ValueError:
            messagebox.showwarning("提示", "年份/月份必须为数字")
            return
        if not (2000 <= year <= 2100):
            messagebox.showwarning("提示", "年份超出合理范围（2000~2100）")
            return
        if not (1 <= month <= 12):
            messagebox.showwarning("提示", "月份必须为 1~12")
            return
        tpl = self.tpl_var.get().strip()
        if not tpl or not os.path.exists(tpl):
            messagebox.showwarning("提示", "请先指定有效的模板文件(.doc)")
            return
        out_dir = self.out_var.get().strip() or BASE_DIR

        items = {}
        for w in self.items_widgets:
            items[w.index] = w.to_item()
        if not self._validate_items(items):
            return

        out_name = kc.build_filename(self.pattern_var.get().strip(), year, month, name)
        out_path = os.path.join(out_dir, out_name)
        if os.path.exists(out_path):
            if not messagebox.askyesno("文件已存在",
                                       f"文件已存在：\n{out_path}\n\n是否覆盖？"):
                return

        self.prog["value"] = 10
        self.status_var.set("正在生成（后台进行，可继续操作界面）...")
        self.btn_generate.config(state=tk.DISABLED)

        def worker():
            pythoncom.CoInitialize()
            try:
                kc.generate_doc(tpl, out_path, name, month, items, year=year)
            except Exception as e:
                self.root.after(0, lambda: self._on_generate_failed(str(e)))
            else:
                self.root.after(0, lambda: self._on_generate_done(out_path, out_dir))
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=worker, daemon=True).start()

    def _validate_items(self, items):
        """校验自评得分/上级评分：必须为 0~标准分 之间的数字。"""
        std = {idx: s for idx, _t, s in ITEM_LABELS}
        for idx, it in items.items():
            for key, cn in (("score", "自评得分"), ("super_score", "上级评分")):
                v = (it.get(key) or "").strip()
                if not v:
                    continue
                try:
                    num = float(v.rstrip("分"))
                except ValueError:
                    messagebox.showwarning("提示",
                                           f"第 {idx} 项「{cn}」=「{v}」不是有效数字")
                    return False
                if num < 0 or num > std[idx]:
                    messagebox.showwarning("提示",
                                           f"第 {idx} 项「{cn}」= {num} 超出 0~{std[idx]} 分范围")
                    return False
        return True

    def _on_generate_done(self, out_path, out_dir):
        self.btn_generate.config(state=tk.NORMAL)
        self.prog["value"] = 100
        self.status_var.set("生成完成：" + out_path)
        self.cfg["out_dir"] = out_dir
        self.cfg["pattern"] = self.pattern_var.get().strip()
        if not _save_config(self.cfg):
            self.status_var.set("生成完成（配置记忆保存失败）：" + out_path)
        if messagebox.askyesno("生成完成",
                               f"已生成：\n{out_path}\n\n是否打开文件？"):
            self._open_file(out_path)
        elif messagebox.askyesno("生成完成",
                                 f"已生成：\n{out_path}\n\n是否打开所在文件夹？"):
            self._open_folder(out_dir)

    def _on_generate_failed(self, err):
        self.btn_generate.config(state=tk.NORMAL)
        self.status_var.set("生成失败：" + err)
        messagebox.showerror("生成失败", err)

    @staticmethod
    def _open_file(path):
        try:
            os.startfile(path)
        except Exception:
            pass

    @staticmethod
    def _open_folder(path):
        try:
            os.startfile(path)
        except Exception:
            pass


def main():
    # 高 DPI 显示器适配（Win8.1+）：让 tkinter 按物理分辨率缩放，避免界面偏小
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()




