#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""主要负责人安全生产责任制履职清单考评表 自动生成工具 GUI v4.0.3（兼容 Windows 7 SP1 ~ Windows 11）"""
import logging
import os
import sys
import json
import threading
import time
import ctypes
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image, ImageTk
import pythoncom

import kaoping_core as kc

VERSION = "4.0.3"

# 主要负责人考评表 13 个考评项（序号/名称/标准分，合计 100 分）
ITEM_LABELS = [
    (1, "安全绩效", 20),
    (2, "隐患排查", 5),
    (3, "安全检查", 5),
    (4, "安全教育", 5),
    (5, "危险源管理", 5),
    (6, "应急管理", 5),
    (7, "违章查处", 10),
    (8, "机构人员", 5),
    (9, "制度规程", 5),
    (10, "费用投入", 5),
    (11, "履职评价", 5),
    (12, "“六个一”工作实绩", 10),
    (13, "甲方契合度", 15),
]

# 双行考评项：主要负责人模板无双行独立填写项，保持空。
SUB_ROW_MAX = {}

# 打包(exe)后 __file__ 指向 _MEIPASS 临时目录，配置文件/默认输出目录需以 exe 目录为准
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "kaoping_config.json")
LOG_PATH = os.path.join(BASE_DIR, "kaoping.log")
THUMB_SIZE = (90, 90)

# ---- 界面配色（统一柔和中性色，减少视觉杂乱）----
C_INPUT_BG = "#F3F6FA"   # 可点击 / 可拖拽区域的底色
C_HINT_FG = "#5A6472"    # 提示文字颜色
C_OK_FG = "#1E7D3C"      # 成功 / 已定位
C_ERR_FG = "#C0392B"     # 失败 / 错误


class ToolTip:
    """简易悬停提示（纯 tkinter 实现，兼容 Windows 7/8/10/11）。"""

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self._tip = None
        self._after_id = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._hide()
        self._after_id = self.widget.after(350, self._show)

    def _show(self):
        self._after_id = None
        if self._tip is not None:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
            self._tip = tk.Toplevel(self.widget)
            self._tip.wm_overrideredirect(True)
            self._tip.wm_geometry("+%d+%d" % (x, y))
            label = tk.Label(self._tip, text=self.text, justify=tk.LEFT,
                             bg="#FFFFF0", fg="#333", relief=tk.SOLID,
                             borderwidth=1, font=("Microsoft YaHei", 9),
                             padx=8, pady=5, wraplength=430)
            label.pack()
        except Exception:
            pass

    def _hide(self, _event=None):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        if self._tip is not None:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


def _setup_logging():
    """日志写入程序同目录 kaoping.log，便于生成失败等问题的排查。"""
    try:
        # 用显式 FileHandler(encoding=...) 而非 basicConfig(encoding=...)：
        # 后者是 Python 3.9 才支持的参数，Win7 构建(Python 3.8) 会抛 TypeError 导致无日志。
        handler = logging.FileHandler(LOG_PATH, mode="a", encoding="utf-8")
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(message)s",
            handlers=[handler])
    except Exception:
        pass


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


def _auto_fill_super(items):
    """自评得分已填的考评项：自动补齐上级评分（=自评得分）与评语（=已完成），与之对应。

    - 仅当「上级评分 / 评价描述」为空时才生成，已手工填写的内容不覆盖；
    - 双行项（班组长第 1 项）的主行与第 2 行(sub)分别处理；
    - 未填自评得分的考评项不生成。
    """
    for item in items.values():
        for d in (item, item.get("sub") or {}):
            score = (d.get("score") or "").strip()
            if not score:
                continue
            if not (d.get("super_score") or "").strip():
                d["super_score"] = score
            if not (d.get("eval_desc") or "").strip():
                d["eval_desc"] = "已完成"




class ImageItem:
    """单个考评项的录入控件：自评描述/得分/材料文字/材料拖拽/评价描述/上级评分。

    sub_max 非空时表示双行考评项（如第 1 项「安全绩效」在模板中占两行），
    额外渲染第 2 行的一组填写框与材料区，数据保存在 to_item() 的 "sub" 子字典中。
    """

    MAX_MATERIALS = 15

    def __init__(self, parent, root, index, title, std_score, sub_max=None):
        self.root = root
        self.index = index
        self.title = title
        self.std_score = std_score
        self.sub_max = sub_max          # 双行项第 2 行标准分；None=单行项
        self.material_paths = []
        self.sub_material_paths = []
        self._thumb_refs = []
        self._sub_thumb_refs = []

        self.frame = ttk.LabelFrame(parent,
                                    text=f"  {index}、{title}（标准分 {std_score} 分） ",
                                    padding=6)
        self.frame.pack(fill=tk.X, pady=4)

        # 主填写行（第 1 行）
        self._build_fields(self.frame, "")
        self._build_material_zone(self.frame, "")

        # 双行项：第 2 行填写区
        if sub_max is not None:
            ttk.Separator(self.frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(6, 2))
            ttk.Label(self.frame, text=f"第 2 行（标准分 {sub_max} 分）",
                      foreground="#1E5AA8",
                      font=("Microsoft YaHei", 9, "bold")).pack(anchor="w", padx=2, pady=(4, 0))
            self._build_fields(self.frame, "sub_", max_score=sub_max)
            self._build_material_zone(self.frame, "sub_")

    # ---------- 字段构建 ----------
    def _build_fields(self, parent, prefix, max_score=None):
        """构建两行输入框（自评描述/自评得分/材料说明 + 评价描述/上级评分）。
        prefix="" 为主行（属性不带前缀），prefix="sub_" 为第 2 行。"""
        grid = ttk.Frame(parent)
        grid.pack(fill=tk.X)
        if max_score is None:
            max_score = self.std_score

        # 第一行：自评描述 / 自评得分 / 材料说明（两行同列对齐，得分旁标注有效范围）
        ttk.Label(grid, text="自评描述:").grid(row=0, column=0, sticky="e", padx=(2, 6), pady=2)
        desc_var = tk.StringVar()
        desc_entry = ttk.Entry(grid, textvariable=desc_var, width=42)
        desc_entry.grid(row=0, column=1, sticky="we", padx=2, pady=2)
        ToolTip(desc_entry, "自评描述：本项工作完成情况的自我说明，将写入考评表「自评描述」栏。")

        ttk.Label(grid, text="自评得分:").grid(row=0, column=2, sticky="e", padx=(10, 4), pady=2)
        score_var = tk.StringVar(value=str(max_score))
        score_entry = ttk.Entry(grid, textvariable=score_var, width=5)
        score_entry.grid(row=0, column=3, padx=2, pady=2)
        ttk.Label(grid, text=f"0~{max_score} 分", foreground="#8A94A6"
                  ).grid(row=0, column=4, sticky="w", padx=(0, 10))
        ToolTip(score_entry, f"自评得分：0~{max_score} 分（默认已填标准分，可修改或留空）。")

        ttk.Label(grid, text="材料说明:").grid(row=0, column=5, sticky="e", padx=(4, 6), pady=2)
        mat_var = tk.StringVar()
        mat_entry = ttk.Entry(grid, textvariable=mat_var, width=16)
        mat_entry.grid(row=0, column=6, sticky="we", padx=2, pady=2)
        ToolTip(mat_entry, "材料说明：本项支撑材料的文字说明（如文件名称或内容摘要），将写入考评表「材料」栏。")

        # 第二行：评价描述 / 上级评分（与第一行保持同列对齐）
        ttk.Label(grid, text="评价描述:").grid(row=1, column=0, sticky="e", padx=(2, 6), pady=2)
        eval_var = tk.StringVar()
        eval_entry = ttk.Entry(grid, textvariable=eval_var, width=42)
        eval_entry.grid(row=1, column=1, sticky="we", padx=2, pady=2)
        ToolTip(eval_entry, "评价描述：上级 / 安环部对本项工作的评价意见，将写入考评表「评价描述」栏。")

        ttk.Label(grid, text="上级评分:").grid(row=1, column=2, sticky="e", padx=(10, 4), pady=2)
        super_var = tk.StringVar()
        super_entry = ttk.Entry(grid, textvariable=super_var, width=5)
        super_entry.grid(row=1, column=3, padx=2, pady=2)
        ttk.Label(grid, text=f"0~{max_score} 分", foreground="#8A94A6"
                  ).grid(row=1, column=4, sticky="w", padx=(0, 10))
        ToolTip(super_entry, f"上级评分：上级 / 安环部评分，0~{max_score} 分，可留空。")

        grid.columnconfigure(1, weight=3)
        grid.columnconfigure(6, weight=2)

        setattr(self, prefix + "desc_var", desc_var)
        setattr(self, prefix + "score_var", score_var)
        setattr(self, prefix + "mat_var", mat_var)
        setattr(self, prefix + "eval_var", eval_var)
        setattr(self, prefix + "super_var", super_var)

    def _build_material_zone(self, parent, prefix):
        """构建支撑材料拖拽区 + 材料预览区（prefix="" 主行、"sub_" 第 2 行）。"""
        img_row = ttk.Frame(parent)
        img_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(img_row, text="支撑材料:").pack(side=tk.LEFT, padx=(2, 6))
        drop_label = tk.Label(
            img_row,
            text=f"点击选择或拖入材料（最多 {self.MAX_MATERIALS} 个）\n"
                 "支持图片 / Word / Excel / PDF；双击预览，右键删除，悬停显示完整文件名",
            bg=C_INPUT_BG, fg=C_HINT_FG, relief=tk.GROOVE,
            height=2, cursor="hand2")
        drop_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        drop_label.bind("<Button-1>", lambda e, p=prefix: self._choose_materials(p))
        drop_label.drop_target_register(DND_FILES)
        drop_label.dnd_bind("<<Drop>>", lambda e, p=prefix: self._on_drop(e, p))
        ToolTip(drop_label,
                f"为「{self.index}、{self.title}」添加支撑材料：\n"
                f"· 点击选择或拖入文件，最多 {self.MAX_MATERIALS} 个\n"
                "· 双击材料预览并打开原文件\n"
                "· 右键删除单个材料\n"
                "· 鼠标悬停显示完整文件名")
        btn_add = ttk.Button(img_row, text="选择材料",
                             command=lambda p=prefix: self._choose_materials(p))
        btn_add.pack(side=tk.LEFT, padx=2)
        btn_clear = ttk.Button(img_row, text="清空材料",
                               command=lambda p=prefix: self._clear_materials(p))
        btn_clear.pack(side=tk.LEFT, padx=2)

        # 材料预览区（图片缩略图 / 文件名称）
        thumb_frame = ttk.Frame(parent)
        thumb_frame.pack(fill=tk.X, pady=(4, 0))

        setattr(self, prefix + "drop_label", drop_label)
        setattr(self, prefix + "btn_add", btn_add)
        setattr(self, prefix + "btn_clear", btn_clear)
        setattr(self, prefix + "thumb_frame", thumb_frame)

    # ---------- 材料增删 / 预览 ----------
    def _mat_paths(self, prefix):
        return self.material_paths if prefix != "sub_" else self.sub_material_paths

    def _thumb_refs_list(self, prefix):
        return self._thumb_refs if prefix != "sub_" else self._sub_thumb_refs

    def _on_drop(self, event, prefix=""):
        files = self.root.tk.splitlist(event.data)
        self.add_materials(files, prefix)

    def _choose_materials(self, prefix=""):
        files = filedialog.askopenfilenames(
            title=f"选择第{self.index}项支撑材料",
            filetypes=[("所有文件", "*.*"),
                       ("图片文件", "*.jpg;*.jpeg;*.png;*.bmp;*.gif"),
                       ("Word 文档", "*.doc;*.docx"),
                       ("Excel 表格", "*.xls;*.xlsx"),
                       ("PDF", "*.pdf")])
        if files:
            self.add_materials(list(files), prefix)

    def add_materials(self, files, prefix=""):
        paths = self._mat_paths(prefix)
        for p in files:
            if len(paths) >= self.MAX_MATERIALS:
                messagebox.showwarning("提示", f"每个考评项最多 {self.MAX_MATERIALS} 个材料")
                break
            if not os.path.exists(p):
                continue
            if p not in paths:
                paths.append(p)
        self._refresh_previews(prefix)

    def _open_material(self, path):
        """双击预览打开材料源文件"""
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showwarning(
                    "无法打开材料",
                    "无法打开该材料文件：\n%s\n\n%s" % (path, e))

    def _refresh_previews(self, prefix=""):
        thumb_frame = getattr(self, prefix + "thumb_frame")
        refs = self._thumb_refs_list(prefix)
        paths = self._mat_paths(prefix)
        for w in thumb_frame.winfo_children():
            w.destroy()
        refs.clear()
        if not paths:
            return
        cols = 6
        for i, p in enumerate(paths):
            r, c = i // cols, i % cols
            short = os.path.basename(p)
            if len(short) > 12:
                short = short[:11] + "…"
            if kc.is_image(p):
                try:
                    with Image.open(p) as im:
                        im.thumbnail(THUMB_SIZE)
                        photo = ImageTk.PhotoImage(im)
                    refs.append(photo)
                    lab = tk.Label(thumb_frame, image=photo, text=short,
                                   compound=tk.TOP, relief=tk.RIDGE,
                                   cursor="hand2", fg="#333")
                except Exception:
                    lab = tk.Label(thumb_frame, text="🖼️\n" + short,
                                   relief=tk.RIDGE, bg="#f0f4ff", fg="#333",
                                   cursor="hand2", justify=tk.CENTER)
            else:
                lab = tk.Label(thumb_frame, text="📄\n" + short,
                               relief=tk.RIDGE, bg="#FFF8E1", fg="#8a5a00",
                               cursor="hand2", justify=tk.CENTER)
            lab.grid(row=r, column=c, padx=3, pady=2, sticky="n")
            full = os.path.basename(p)
            lab.bind("<Double-Button-1>", lambda e, path=p: self._open_material(path))
            lab.bind("<Button-3>", lambda e, path=p, pfx=prefix: self._remove_material(path, pfx))
            lab.bind("<Enter>", lambda e, ww=lab, full=full: ww.config(text=full))
            lab.bind("<Leave>", lambda e, ww=lab, s=short: ww.config(text=s))

    def to_item(self):
        d = {
            "desc": self.desc_var.get().strip(),
            "score": self.score_var.get().strip(),
            "material_text": self.mat_var.get().strip(),
            "materials": list(self.material_paths),
            "eval_desc": self.eval_var.get().strip(),
            "super_score": self.super_var.get().strip(),
        }
        if self.sub_max is not None:
            d["sub"] = {
                "desc": self.sub_desc_var.get().strip(),
                "score": self.sub_score_var.get().strip(),
                "material_text": self.sub_mat_var.get().strip(),
                "materials": list(self.sub_material_paths),
                "eval_desc": self.sub_eval_var.get().strip(),
                "super_score": self.sub_super_var.get().strip(),
            }
        return d

    def apply_sub(self, data):
        """把读取到的第 2 行数据回填到界面（双行项）。"""
        data = data or {}
        if self.sub_max is None:
            return
        if data.get("desc"):
            self.sub_desc_var.set(data["desc"])
        if data.get("score"):
            self.sub_score_var.set(data["score"])
        if data.get("material_text"):
            self.sub_mat_var.set(data["material_text"])
        if data.get("eval_desc"):
            self.sub_eval_var.set(data["eval_desc"])
        if data.get("super_score"):
            self.sub_super_var.set(data["super_score"])
        mats = [p for p in (data.get("materials") or []) if p and os.path.exists(p)]
        if mats:
            self.add_materials(mats, "sub_")

    def clear(self):
        self.desc_var.set("")
        self.score_var.set(str(self.std_score))
        self.mat_var.set("")
        self.eval_var.set("")
        self.super_var.set("")
        self._clear_materials()
        if self.sub_max is not None:
            self.sub_desc_var.set("")
            self.sub_score_var.set(str(self.sub_max))
            self.sub_mat_var.set("")
            self.sub_eval_var.set("")
            self.sub_super_var.set("")
            self._clear_materials("sub_")

    def _clear_materials(self, prefix=""):
        self._mat_paths(prefix).clear()
        self._refresh_previews(prefix)

    def _remove_material(self, path, prefix=""):
        paths = self._mat_paths(prefix)
        if path in paths:
            paths.remove(path)
            self._refresh_previews(prefix)
class App:
    def __init__(self, root):
        self.root = root
        self.cfg = _load_config()
        self.items_widgets = []
        root.title(f"主要负责人安全生产责任制履职清单考评表自动生成工具 v{VERSION}")
        root.geometry("1180x860")
        root.minsize(1000, 700)
        self._build_ui()
        self._load_template_path()
        # 首次运行自动生成本地文件名权重配置（material_rules.json）
        kc.ensure_material_rules_file(BASE_DIR)

    # ============ 界面构建 ============
    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill=tk.BOTH, expand=True)

        # ================= ① 基本信息与输出 =================
        top = ttk.LabelFrame(outer, text=" ① 基本信息与输出 ", padding=8)
        top.pack(fill=tk.X, pady=(0, 6))

        r1 = ttk.Frame(top)
        r1.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(r1, text="年份:").pack(side=tk.LEFT)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Spinbox(r1, from_=2020, to=2040, textvariable=self.year_var, width=6).pack(side=tk.LEFT, padx=(2, 10))
        ttk.Label(r1, text="月份:").pack(side=tk.LEFT)
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        ttk.Spinbox(r1, from_=1, to=12, textvariable=self.month_var, width=5).pack(side=tk.LEFT, padx=(2, 10))
        ttk.Label(r1, text="姓名:").pack(side=tk.LEFT)
        self.name_var = tk.StringVar()
        self.name_combo = ttk.Combobox(r1, textvariable=self.name_var, width=10)
        self.name_combo.pack(side=tk.LEFT, padx=(2, 10))
        ttk.Label(r1, text="命名模板:").pack(side=tk.LEFT)
        self.pattern_var = tk.StringVar(value=self.cfg.get("pattern", kc.DEFAULT_NAME_PATTERN))
        pattern_entry = ttk.Entry(r1, textvariable=self.pattern_var, width=38)
        pattern_entry.pack(side=tk.LEFT, padx=(2, 4))
        ToolTip(pattern_entry,
                "输出文件名模板：\n"
                "示例：2026年8月张三主要负责人履职考评表\n\n"
                "占位符说明：\n"
                "{Y} 或 {年}     = 年份\n"
                "{X} 或 {月份}   = 月份\n"
                "{XXX} 或 {姓名} = 姓名\n\n"
                "可自由组合，修改后会自动保存记忆。")
        ttk.Label(r1, text="（{Y}=年 {X}=月 {XXX}=姓名）",
                  foreground="#8A94A6").pack(side=tk.LEFT, padx=(4, 0))

        r2 = ttk.Frame(top)
        r2.pack(fill=tk.X)
        ttk.Label(r2, text="考评表模板:").pack(side=tk.LEFT)
        self.tpl_var = tk.StringVar()
        self.tpl_label = tk.Label(r2, text="自动定位模板…", bg=C_INPUT_BG, fg=C_HINT_FG,
                                  relief=tk.GROOVE, cursor="hand2")
        self.tpl_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 2))
        self.tpl_label.bind("<Button-1>", lambda e: self._sel_template())
        self.tpl_label.drop_target_register(DND_FILES)
        self.tpl_label.dnd_bind("<<Drop>>", self._drop_template)
        ToolTip(self.tpl_label,
                "考评表模板（.doc）：点击选择或直接拖入模板文件。\n"
                "程序默认已自动定位到程序目录下的模板，一般无需修改。")
        ttk.Button(r2, text="选择", width=7, command=self._sel_template).pack(side=tk.LEFT, padx=2)
        ttk.Label(r2, text="输出目录:").pack(side=tk.LEFT, padx=(12, 2))
        self.out_var = tk.StringVar(value=self.cfg.get("out_dir", BASE_DIR))
        self.out_label = tk.Label(r2, text=self.out_var.get(), bg=C_INPUT_BG, fg=C_HINT_FG,
                                  relief=tk.GROOVE, cursor="hand2")
        self.out_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 2))
        self.out_label.bind("<Button-1>", lambda e: self._sel_outdir())
        ToolTip(self.out_label,
                "输出目录：生成后的 .doc 考评表保存位置，点击可重新选择。\n"
                "读取已生成考评表时，提取出的支撑材料会保存到「输出目录\\提取材料\\姓名月份」子目录。")
        ttk.Button(r2, text="选择", width=7, command=self._sel_outdir).pack(side=tk.LEFT, padx=2)

        # ================= ② 材料导入（可选） =================
        mat = ttk.LabelFrame(outer, text=" ② 材料导入（可选） ", padding=8)
        mat.pack(fill=tk.X, pady=(0, 6))

        r3 = ttk.Frame(mat)
        r3.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(r3, text="支撑材料文件夹:").pack(side=tk.LEFT)
        self.mat_folder_label = tk.Label(
            r3, text=f"选择当月文件夹或年度根目录，按文件名权重自动匹配 {len(ITEM_LABELS)} 个考评项",
            bg=C_INPUT_BG, fg=C_HINT_FG, relief=tk.GROOVE, cursor="hand2")
        self.mat_folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.mat_folder_label.bind("<Button-1>", lambda e: self._load_material_folder())
        self.mat_folder_label.drop_target_register(DND_FILES)
        self.mat_folder_label.dnd_bind("<<Drop>>", self._drop_material_folder)
        ToolTip(self.mat_folder_label,
                f"按文件名权重把当月支撑材料自动匹配到 {len(ITEM_LABELS)} 个考评项：\n"
                "· 选择当月文件夹；或选年度根目录，程序按界面年份/月份自动定位当月（如 2026.08）\n"
                "· 未匹配的文件会弹出提示，可手动拖入对应考评项\n"
                "· 匹配规则保存在 material_rules.json，可用「编辑权重」查看或修改")
        ttk.Button(r3, text="选择", width=7, command=self._load_material_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(r3, text="编辑权重", width=9, command=self._open_material_rules).pack(side=tk.LEFT, padx=2)

        r4 = ttk.Frame(mat)
        r4.pack(fill=tk.X)
        ttk.Label(r4, text="读取考评表:").pack(side=tk.LEFT)
        self.import_label = tk.Label(
            r4, text=f"选择已生成的考评表(.doc)，回填 {len(ITEM_LABELS)} 项评分/评价与支撑材料，便于替换后生成新月份",
            bg=C_INPUT_BG, fg=C_HINT_FG, relief=tk.GROOVE, cursor="hand2")
        self.import_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.import_label.bind("<Button-1>", lambda e: self._read_kaoping_doc())
        self.import_label.drop_target_register(DND_FILES)
        self.import_label.dnd_bind("<<Drop>>", self._drop_kaoping_doc)
        ToolTip(self.import_label,
                "复用已生成的考评表（.doc）：\n"
                f"自动读取姓名/月份、{len(ITEM_LABELS)} 项自评与上级评分/评价、内嵌的支撑材料\n"
                "（图片、Excel、Word 等），回填界面后可替换材料、改月份再生成新表。\n"
                "提取出的材料保存到「输出目录\\提取材料\\姓名月份」子目录。")
        ttk.Button(r4, text="读取考评表", width=9, command=self._read_kaoping_doc).pack(side=tk.LEFT, padx=2)

        r5 = ttk.Frame(mat)
        r5.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(r5, text="履职履责表:").pack(side=tk.LEFT)
        self.eval_label = tk.Label(
            r5, text=f"选择月度履职履责表(.xlsx)，自动填入 {len(ITEM_LABELS)} 项自评得分（excel 只提供姓名和分数）",
            bg=C_INPUT_BG, fg=C_HINT_FG, relief=tk.GROOVE, cursor="hand2")
        self.eval_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        self.eval_label.bind("<Button-1>", lambda e: self._read_eval_scores())
        self.eval_label.drop_target_register(DND_FILES)
        self.eval_label.dnd_bind("<<Drop>>", self._drop_eval_scores)
        ToolTip(self.eval_label,
                "读取月度履职履责表（.xlsx）：\n"
                f"按「主要负责人月度履职评价表」sheet 读取所选主要负责人的 {len(ITEM_LABELS)} 项评分（excel 只提供姓名和分数），\n"
                "自动填入各考评项的「自评得分」。\n"
                "文件中包含多个姓名时，会弹出对话框让您选择本次生成的主要负责人。")
        ttk.Button(r5, text="读取评分", width=9, command=self._read_eval_scores).pack(side=tk.LEFT, padx=2)

        # ---- 中部：13 项滚动列表 ----
        mid_lf = ttk.LabelFrame(outer,
                                text=f" ③ 考评项填写（{len(ITEM_LABELS)} 项；每项可拖入图片 / Word / Excel / PDF 等支撑材料） ",
                                padding=4)
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
            w = ImageItem(self.list_frame, self.root, idx, title, std,
                          sub_max=SUB_ROW_MAX.get(idx))
            self.items_widgets.append(w)

        # 鼠标滚轮：悬停在任意子控件上也能滚动中部列表
        self._bind_mousewheel(mid_lf)

        # ================= 底部工具栏 =================
        bottom = ttk.Frame(outer)
        bottom.pack(fill=tk.X, pady=(8, 0))
        # “一键生成”最先 pack 并靠右，窗口缩小也保留空间不被遮挡
        self.btn_generate = ttk.Button(bottom, text="一键生成", command=self._generate)
        self.btn_generate.pack(side=tk.RIGHT, padx=2)
        ToolTip(self.btn_generate,
                "校验通过后，在后台用 Word 按「命名模板」生成考评表到输出目录，\n"
                "完成后可选择打开文件或所在文件夹。")
        ttk.Button(bottom, text="清空全部", command=self._clear_all).pack(side=tk.LEFT, padx=2)
        self.prog = ttk.Progressbar(bottom, length=260, mode="determinate")
        self.prog.pack(side=tk.LEFT, padx=10)
        self.status_var = tk.StringVar(value="就绪：填写信息、导入材料后，点击右下角「一键生成」")
        ttk.Label(bottom, textvariable=self.status_var, width=70,
                  foreground="#5A6472").pack(side=tk.LEFT, padx=8)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _bind_mousewheel(self, widget):
        """递归绑定滚轮事件，使鼠标悬停在任意子控件上也能滚动中部列表。"""
        widget.bind("<MouseWheel>", self._on_mousewheel)
        for child in widget.winfo_children():
            self._bind_mousewheel(child)

    # ============ 模板 / 输出目录 ============
    def _load_template_path(self):
        try:
            tpl = kc.find_template()
            self.tpl_var.set(tpl)
            self.tpl_label.config(text=tpl, fg=C_OK_FG)
        except FileNotFoundError as e:
            self.tpl_var.set("")
            self.tpl_label.config(text=str(e), fg=C_ERR_FG)

    def _sel_template(self):
        p = filedialog.askopenfilename(title="选择考评表模板(.doc)",
                                       filetypes=[("Word 文档", "*.doc;*.docx"), ("所有文件", "*.*")])
        if p:
            self.tpl_var.set(p)
            self.tpl_label.config(text=p, fg=C_OK_FG)

    def _drop_template(self, event):
        files = self.root.tk.splitlist(event.data)
        if files:
            self.tpl_var.set(files[0])
            self.tpl_label.config(text=files[0], fg=C_OK_FG)

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
                title="选择支撑材料文件夹或年度根目录（自动定位当月子文件夹）")
        if not path:
            return
        # 年度根目录自动定位当月子文件夹（如 2026.08 / 2026.8 / 8月）
        try:
            year = int(self.year_var.get().strip())
            month = int(self.month_var.get().strip())
        except ValueError:
            year, month = None, None
        if year and month and os.path.isdir(path):
            sub = kc.find_month_subfolder(path, year, month)
            if sub != os.path.abspath(path):
                path = sub
        self.status_var.set(f"正在扫描：{os.path.basename(path)} ...")
        self.mat_folder_label.config(text=f"扫描中… {os.path.basename(path)}", fg=C_HINT_FG)
        rules = kc.load_material_rules(os.path.join(BASE_DIR, kc.MATERIAL_RULES_FILE))

        def worker():
            try:
                result, unmatched = kc.scan_materials_folder(path, rules=rules)
            except Exception as e:
                self.root.after(0, lambda err=str(e): self._on_scan_failed(err))
                return
            self.root.after(0, lambda r=result, u=unmatched, p=path:
                            self._on_scan_done(r, u, p))

        threading.Thread(target=worker, daemon=True).start()

    def _on_scan_done(self, result, unmatched, path):
        filled = 0
        for w in self.items_widgets:
            paths = result.get(w.index, [])
            if paths:
                before = len(w.material_paths)
                w.add_materials(paths)
                filled += len(w.material_paths) - before
        self.mat_folder_label.config(text=os.path.basename(path), fg=C_OK_FG)
        parts = [f"{i}项{len(result[i])}个" for i in range(1, len(ITEM_LABELS) + 1) if result[i]]
        detail = "、".join(parts)
        if len(detail) > 80:
            detail = detail[:77] + "…"
        suffix = (f"；{len(unmatched)} 个文件未匹配（可手动拖入）" if unmatched
                  else "；全部匹配")
        self.status_var.set(f"已自动填入 {filled} 个支撑材料{suffix}（{detail}）")
        # 未匹配文件过多时仅状态栏提示，避免弹窗刷屏
        if 0 < len(unmatched) <= 200:
            names = "\n".join(os.path.basename(u) for u in unmatched[:40])
            if len(unmatched) > 40:
                names += f"\n……共 {len(unmatched)} 个未匹配"
            messagebox.showinfo("未匹配文件",
                                "以下文件未匹配到任何考评项，请手动拖入对应项：\n\n" + names)

    def _on_scan_failed(self, err):
        logging.error("扫描支撑材料失败: %s", err)
        self.mat_folder_label.config(text="扫描失败", fg=C_ERR_FG)
        self.status_var.set("扫描失败")
        messagebox.showerror("扫描失败", err)

    # ============ 读取已生成考评表(.doc) ============
    def _read_kaoping_doc(self):
        """选择已生成的考评表(.doc)，把评分/评价/支撑材料回填到界面。"""
        path = filedialog.askopenfilename(title="选择已生成的考评表(.doc)",
                                          filetypes=[("Word 文档", "*.doc"),
                                                     ("所有文件", "*.*")])
        if path:
            self._apply_kaoping_doc(path)

    def _drop_kaoping_doc(self, event):
        for p in self.root.tk.splitlist(event.data):
            if os.path.isfile(p):
                self._apply_kaoping_doc(p)
                return

    def _apply_kaoping_doc(self, path):
        """读取 .doc：姓名/月份/各项自评与上级评价/支撑材料回填到界面，材料提取到输出目录\\提取材料。"""
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "文件不存在：" + str(path))
            return
        out_dir = os.path.join(self.out_var.get().strip() or BASE_DIR, "提取材料")
        try:
            self.status_var.set("正在读取考评表…")
            data = kc.extract_kaoping_doc(path, out_dir=out_dir)
        except Exception as e:
            logging.error("读取考评表失败: %s", e)
            messagebox.showerror("读取失败", str(e))
            return
        name = data.get("name") or ""
        month = data.get("month") or ""
        if name:
            self.name_var.set(name)
        if month:
            self.month_var.set(month)
        mat_count = 0
        for idx, it in data["items"].items():
            w = self.items_widgets[idx - 1]
            if it.get("desc"):
                w.desc_var.set(it["desc"])
            if it.get("score"):
                w.score_var.set(it["score"])
            if it.get("material_text"):
                w.mat_var.set(it["material_text"])
            if it.get("eval_desc"):
                w.eval_var.set(it["eval_desc"])
            if it.get("super_score"):
                w.super_var.set(it["super_score"])
            mats = [p for p in (it.get("materials") or []) if p and os.path.exists(p)]
            if mats:
                w.add_materials(mats)
                mat_count += len(mats)
            if it.get("sub"):
                w.apply_sub(it["sub"])
        self.import_label.config(text=os.path.basename(path), fg=C_OK_FG)
        who = name + ("%s月" % month if month else "")
        mat_dir = data.get("out_dir") or out_dir
        msg = "已读取 %s 考评表：%d 项内容 + %d 个支撑材料%s" % (
            who or os.path.basename(path), len(ITEM_LABELS), mat_count,
            "；材料已提取到 " + mat_dir if mat_count else "")
        self.status_var.set(msg)
        if data.get("warnings"):
            messagebox.showwarning("读取提示", "\n".join(data["warnings"]))

    # ============ 月度履职履责表(xlsx) 评分读取 ============
    def _read_eval_scores(self, path=None):
        """读取履职履责表 .xlsx：按姓名把各考评项评分填入「自评得分」（excel 只提供姓名和分数）。"""
        if not path:
            path = filedialog.askopenfilename(
                title="选择月度履职履责表",
                filetypes=[("Excel 工作簿", "*.xlsx;*.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            names = kc.list_eval_names(path)
        except Exception as e:
            logging.error("读取履职履责表失败: %s", e)
            messagebox.showerror("读取失败", str(e))
            return
        if not names:
            messagebox.showwarning("提示", "履职履责表中未找到姓名行")
            return
        self.name_combo["values"] = names
        cur = self.name_var.get().strip()
        if cur not in names:
            if len(names) == 1:
                cur = names[0]
            else:
                cur = self._pick_eval_name(names)
                if not cur:
                    return
            self.name_var.set(cur)
        try:
            data = kc.read_eval_scores(path, cur)
        except Exception as e:
            logging.error("读取履职履责表评分失败: %s", e)
            messagebox.showwarning("读取失败", str(e))
            return
        filled = 0
        for idx, it in (data.get("items") or {}).items():
            w = self.items_widgets[idx - 1]
            if it.get("score"):
                w.score_var.set(it["score"])
                filled += 1
        self.eval_label.config(text=os.path.basename(path), fg=C_OK_FG)
        msg = "已读取 %s 的履职履责表：%d 项评分已填入" % (cur, filled)
        total = data.get("total") or ""
        if total:
            msg += "，总分 %s" % total
        self.status_var.set(msg)

    def _drop_eval_scores(self, event):
        for p in self.root.tk.splitlist(event.data):
            if os.path.isfile(p) and p.lower().endswith((".xlsx", ".xlsm", ".xls")):
                self._read_eval_scores(p)
                return

    def _pick_eval_name(self, names):
        """履职履责表含多个姓名时弹出选择框，返回选中姓名或 None。"""
        win = tk.Toplevel(self.root)
        win.title("选择姓名")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        ttk.Label(win, text="履职履责表中包含多个主要负责人，请选择本次生成的姓名：",
                  padding=(14, 10)).pack()
        var = tk.StringVar(value=names[0])
        combo = ttk.Combobox(win, textvariable=var, values=list(names),
                             state="readonly", width=14)
        combo.pack(padx=14, pady=(0, 10))
        result = {}

        def _ok():
            result["v"] = var.get()
            win.destroy()

        ttk.Button(win, text="确定", command=_ok).pack(pady=(0, 12))
        win.bind("<Return>", lambda e: _ok())
        win.after(100, combo.focus_set)
        self.root.wait_window(win)
        return result.get("v")

    # ============ 文件名权重配置 ============
    def _open_material_rules(self):
        """打开本地文件名权重配置文件（material_rules.json），便于按业务调整关键词。"""
        path = kc.ensure_material_rules_file(BASE_DIR)
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("打开失败",
                                 f"无法打开权重配置文件：\n{path}\n\n{e}")

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
        # 自评得分已填的考评项：上级评分自动等于自评得分、评语自动填"已完成"（已填内容不覆盖）
        _auto_fill_super(items)
        if not self._validate_items(items):
            return

        out_name = kc.build_filename(self.pattern_var.get().strip(), year, month, name)
        out_path = os.path.join(out_dir, out_name)
        if os.path.exists(out_path):
            if not messagebox.askyesno("文件已存在",
                                       f"文件已存在：\n{out_path}\n\n是否覆盖？"):
                return

        self.prog["value"] = 10
        self._gen_started = time.time()
        self.status_var.set("正在生成（后台进行，可继续操作界面）...")
        self.btn_generate.config(state=tk.DISABLED)

        def worker():
            pythoncom.CoInitialize()
            try:
                def cb(i):
                    # 进度 10 → 88（逐项推进），完成时置 100
                    self.root.after(0, lambda v=10 + int(i * (88 - 10) / len(ITEM_LABELS)): self.prog.config(value=v))

                kc.generate_doc(tpl, out_path, name, month, items,
                                year=year, progress_cb=cb)
            except Exception as e:
                logging.exception("生成失败: %s", out_path)
                self.root.after(0, lambda: self._on_generate_failed(str(e)))
            else:
                self.root.after(0, lambda: self._on_generate_done(out_path, out_dir))
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=worker, daemon=True).start()

    def _validate_items(self, items):
        """校验自评得分/上级评分：必须为 0~标准分 之间的数字（双行项第 2 行按 SUB_ROW_MAX 校验）。"""
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
                # 双行项第 1 行的有效范围 = 标准分 - 第 2 行标准分
                sub_max = SUB_ROW_MAX.get(idx)
                if sub_max is not None:
                    if num < 0 or num > std[idx] - sub_max:
                        messagebox.showwarning(
                            "提示",
                            f"第 {idx} 项第 1 行「{cn}」= {num} 超出 0~{std[idx] - sub_max} 分范围")
                        return False
            # 双行项第 2 行（sub）
            sub_max = SUB_ROW_MAX.get(idx)
            if sub_max is not None:
                sub = it.get("sub") or {}
                for key, cn in (("score", "自评得分"), ("super_score", "上级评分")):
                    v = (sub.get(key) or "").strip()
                    if not v:
                        continue
                    try:
                        num = float(v.rstrip("分"))
                    except ValueError:
                        messagebox.showwarning("提示",
                                               f"第 {idx} 项第 2 行「{cn}」=「{v}」不是有效数字")
                        return False
                    if num < 0 or num > sub_max:
                        messagebox.showwarning("提示",
                                               f"第 {idx} 项第 2 行「{cn}」= {num} 超出 0~{sub_max} 分范围")
                        return False
        return True

    def _on_generate_done(self, out_path, out_dir):
        self.btn_generate.config(state=tk.NORMAL)
        self.prog["value"] = 100
        elapsed = time.time() - getattr(self, "_gen_started", time.time())
        logging.info("生成成功: %s (耗时 %.1f 秒)", out_path, elapsed)
        self.status_var.set(f"生成完成（{elapsed:.1f} 秒）：" + out_path)
        self.cfg["out_dir"] = out_dir
        self.cfg["pattern"] = self.pattern_var.get().strip()
        if not _save_config(self.cfg):
            self.status_var.set("生成完成（配置记忆保存失败）：" + out_path)
        try:
            size_mb = os.path.getsize(out_path) / 1048576.0
            if size_mb > 100:
                messagebox.showwarning(
                    "文件体积过大",
                    f"生成的文档约 {size_mb:.0f}MB。\n\n"
                    "非图片材料（Excel/Word/PDF 等）会以 OLE 方式整文件嵌入文档，\n"
                    "体积由源文件决定；图片类材料已在嵌入前压缩。\n"
                    "如需控制体积，请精简源文件后重新生成。")
        except Exception:
            pass
        if messagebox.askyesno("生成完成",
                               f"已生成：\n{out_path}\n\n是否打开文件？"):
            self._open_file(out_path)
        elif messagebox.askyesno("生成完成",
                                 f"已生成：\n{out_path}\n\n是否打开所在文件夹？"):
            self._open_file(out_dir)   # startfile 打开目录

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


def _self_check():
    """启动时校验岗位参数自洽：core 行映射 + GUI 考评项/标准分/权重首词/双行拆分。

    与 kc.self_check() 配合使用；发现不一致直接抛 AssertionError，
    把「模板被换错/参数被改坏」这类事故在启动时拦下。
    """
    kc.self_check()
    problems = []
    if len(ITEM_LABELS) != len(kc.ITEM_ROWS):
        problems.append("GUI 考评项数(%d) 与 core ITEM_ROWS(%d) 不一致"
                        % (len(ITEM_LABELS), len(kc.ITEM_ROWS)))
    if sorted(i for i, _n, _s in ITEM_LABELS) != sorted(kc.ITEM_ROWS):
        problems.append("GUI ITEM_LABELS 序号应与 core ITEM_ROWS 一致")
    total_score = sum(s for _i, _n, s in ITEM_LABELS)
    if total_score != 100:
        problems.append("考评项标准分合计应为 100，实际 %d" % total_score)
    score_of = dict((i, s) for i, _n, s in ITEM_LABELS)
    for idx, label, score in ITEM_LABELS:
        kws = kc.ITEM_MATCH_RULES.get(idx) or []
        if kws and kws[0] not in label:
            problems.append("第 %d 项「%s」未含权重首词「%s」" % (idx, label, kws[0]))
    sub = globals().get("SUB_ROW_MAX") or {}
    for idx, split_score in sub.items():
        score = score_of.get(idx)
        if score is None:
            problems.append("SUB_ROW_MAX 序号 %d 不存在于考评项" % idx)
        elif not (0 < split_score < score):
            problems.append("SUB_ROW_MAX[%d]=%d 应在 (0, %d) 内" % (idx, split_score, score))
    if problems:
        raise AssertionError("GUI 岗位参数自检失败：\n- " + "\n- ".join(problems))


def main():
    # 高 DPI 显示器适配（Win8.1+）：让 tkinter 按物理分辨率缩放，避免界面偏小
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    _setup_logging()
    try:
        _self_check()
    except AssertionError as e:
        messagebox.showerror("岗位参数自检失败", str(e))
        raise
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()




