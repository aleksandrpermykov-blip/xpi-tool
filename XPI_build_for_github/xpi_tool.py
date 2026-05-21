#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Raychem XPI — Подбор гильз и муфт
Десктоп-утилита. Запуск: python xpi_tool.py
Требования: Python 3.8+  (tkinter входит в стандартную поставку)
"""

import tkinter as tk
from tkinter import ttk, font as tkfont
import json, math

# ════════════════════════════════════════════════════════════════════════════
#  ВСТРОЕННЫЕ ДАННЫЕ  (извлечены из Excel XPI_Подбор_гильз_Raychem_v8.xlsx)
# ════════════════════════════════════════════════════════════════════════════
_CABLES_RAW = [
["XPI-7",2.5,7],["XPI-10",1.8,10],["XPI-11.7",1.5,11.7],["XPI-15",1.2,15],
["XPI-17.8",1.0,17.8],["XPI-25",1.0,25],["XPI-31.5",1.6,31.5],["XPI-50",1.0,50],
["XPI-65",0.8,65],["XPI-80",1.2,80],["XPI-100",1.5,100],["XPI-150",1.0,150],
["XPI-200",0.8,200],["XPI-320",0.9,320],["XPI-380",0.8,380],["XPI-480",0.6,480],
["XPI-600",0.5,600],["XPI-700",0.4,700],["XPI-810",0.6,810],["XPI-1000",0.5,1000],
["XPI-1440",0.3,1440],["XPI-1750",0.3,1750],["XPI-2000",0.5,2000],["XPI-3000",0.3,3000],
["XPI-4000",0.3,4000],["XPI-4400",0.3,4400],["XPI-5160",0.3,5160],["XPI-5600",0.2,5600],
["XPI-7000",0.2,7000],["XPI-8000",0.1,8000],
["XPI-0.8",22.4,0.8],["XPI-1.1",16.1,1.1],["XPI-1.8",10.2,1.8],
["XPI-2.9",5.9,2.9],["XPI-4.4",4.0,4.4],
["XPI-S-7",2.5,7],["XPI-S-10",1.8,10],["XPI-S-11.7",1.5,11.7],["XPI-S-15",1.2,15],
["XPI-S-17.8",1.0,17.8],["XPI-S-25",1.0,25],["XPI-S-31.5",1.6,31.5],["XPI-S-50",1.0,50],
["XPI-S-65",0.8,65],["XPI-S-80",1.2,80],["XPI-S-100",1.5,100],["XPI-S-150",1.0,150],
["XPI-S-200",0.8,200],["XPI-S-320",0.9,320],["XPI-S-380",0.8,380],["XPI-S-480",0.6,480],
["XPI-S-600",0.5,600],["XPI-S-700",0.4,700],["XPI-S-810",0.6,810],["XPI-S-1000",0.5,1000],
["XPI-S-1440",0.3,1440],["XPI-S-1750",0.3,1750],["XPI-S-2000",0.5,2000],["XPI-S-3000",0.3,3000],
["XPI-S-4000",0.3,4000],["XPI-S-4400",0.3,4400],["XPI-S-5160",0.3,5160],["XPI-S-5600",0.2,5600],
["XPI-S-7000",0.2,7000],["XPI-S-8000",0.1,8000],
["XPI-S-0.8",22.4,0.8],["XPI-S-1.1",16.1,1.1],["XPI-S-1.8",10.2,1.8],
["XPI-S-2.9",5.9,2.9],["XPI-S-4.4",4.0,4.4],
]

# sleeve data: key "C1|C2" → [sleeve, catalog, matrix, tool, kit]
_SLEEVES = {}

# braid data: key "KIT|FAMILY" → [sleeve, catalog, matrix, tool]
_BRAID = {
    "CS-150-2.5-PI|XPI":  ["BR-CRP-2.5","1244-016304","CD-PI-02","PI-TOOL-01"],
    "CS-150-2.5-PI|XPI-S":["BR-CRP-2.5","1244-016304","CD-PI-02","PI-TOOL-01"],
    "CS-150-6-PI|XPI":    ["BR-CRP-6",  "1244-016305","CD-PI-03","PI-TOOL-02"],
    "CS-150-6-PI|XPI-S":  ["BR-CRP-6",  "1244-016305","CD-PI-03","PI-TOOL-02"],
    "CS-150-25-PI|XPI":   ["BR-CRP-25", "1244-016306","CD-PI-04","PI-TOOL-02"],
    "CS-150-25-PI|XPI-S": ["BR-CRP-25", "1244-016306","CD-PI-04","PI-TOOL-02"],
}

# ── compatibility + sleeve calculation (mirrors Excel logic) ─────────────────
_R = {0.8:22.4,1.1:16.1,1.8:10.2,2.9:5.9,4.4:4.0,7:2.5,10:1.8,
      11.7:1.5,15:1.2,17.8:1.0,25:1.0,31.5:1.6,50:1.0,65:0.8,
      80:1.2,100:1.5,150:1.0,200:0.8,320:0.9,380:0.8,480:0.6,
      600:0.5,700:0.4,810:0.6,1000:0.5,1440:0.3,1750:0.3,
      2000:0.5,3000:0.3,4000:0.3,4400:0.3,5160:0.3,5600:0.2,
      7000:0.2,8000:0.1}

def _family(name):
    if name.startswith("XPI-S-"): return "XPI-S"
    return "XPI"

def _val(name):
    return float(name.replace("XPI-S-","").replace("XPI-",""))

def _cond(name):
    return _R.get(_val(name))

def _sleeve(c1, c2):
    v1,v2 = _val(c1),_val(c2)
    f1,f2 = _family(c1),_family(c2)
    cd1,cd2 = _cond(c1),_cond(c2)
    if not cd1 or not cd2: return None
    s1,s2 = cd1<=2.5, cd2<=2.5

    if s1 and s2:
        def tiny(v,f): return v>=65 or v in [380,480,600,700,810,1000,1440,1750,2000,3000,4000,4400,5160,5600,7000,8000]
        def med(v,f):  return v in [11.7,15,17.8,25,50,80,150,320] or (v==100 and f in ["XPI","XPI-S"])
        def ls(v,f):   return v in [7,10,11.7,31.5] or (v==100 and f in ["XPI","XPI-S"])
        t1,t2 = tiny(v1,f1),tiny(v2,f2)
        m1,m2 = med(v1,f1),med(v2,f2)
        l1,l2 = ls(v1,f1),ls(v2,f2)
        if t1 and t2:                           return ("PI-CRP-01N","1244-016256","CD-PI-02","PI-TOOL-01","CS-150-2.5-PI")
        if (v1==11.7 and t2) or (t1 and v2==11.7): return ("PI-CRP-02N","1244-016257","CD-PI-02","PI-TOOL-01","CS-150-2.5-PI")
        if m1 and m2:                           return ("PI-CRP-03N","1244-016258","CD-PI-02","PI-TOOL-01","CS-150-2.5-PI")
        if (l1 and t2) or (t1 and l2):         return ("PI-CRP-04", "1244-016259","CD-PI-02","PI-TOOL-01","CS-150-2.5-PI")
        if (l1 and m2) or (m1 and l2):         return ("PI-CRP-05", "1244-016260","CD-PI-02","PI-TOOL-01","CS-150-2.5-PI")
        if l1 and l2:                           return ("PI-CRP-06", "1244-016261","CD-PI-02","PI-TOOL-01","CS-150-2.5-PI")
        return None

    for ca,cb in [(c1,c2),(c2,c1)]:
        va,vb,fb = _val(ca),_val(cb),_family(cb)
        if va==4.4:
            if vb in [11.7,15]: return ("PI-CRP-07","1244-016262","CD-PI-03","PI-TOOL-02","CS-150-6-PI")
            if vb in [7,10]:    return ("PI-CRP-08","1244-016263","CD-PI-03","PI-TOOL-02","CS-150-6-PI")
            if vb==4.4:         return ("PI-CRP-09","1244-016264","CD-PI-03","PI-TOOL-02","CS-150-6-PI")
        if va==2.9:
            if vb in [11.7,31.5] or (vb==100 and fb in ["XPI","XPI-S"]): return ("PI-CRP-10","1244-016265","CD-PI-04","PI-TOOL-02","CS-150-6-PI")
            if vb in [7,10]:    return ("PI-CRP-11","1244-016266","CD-PI-04","PI-TOOL-02","CS-150-6-PI")
            if vb==4.4:         return ("PI-CRP-12","1244-016267","CD-PI-04","PI-TOOL-02","CS-150-6-PI")
            if vb==2.9:         return ("PI-CRP-13","1244-016268","CD-PI-04","PI-TOOL-02","CS-150-6-PI")
        if va==1.8:
            if vb==7:           return ("PI-CRP-14","1244-016269","CD-PI-04","PI-TOOL-02","CS-150-25-PI")
            if vb in [7,4.4]:   return ("PI-CRP-15","1244-016270","CD-PI-04","PI-TOOL-02","CS-150-25-PI")
            if vb==2.9:         return ("PI-CRP-16","1244-016271","CD-PI-04","PI-TOOL-02","CS-150-25-PI")
            if vb==1.8:         return ("PI-CRP-17","1244-016272","CD-PI-04","PI-TOOL-02","CS-150-25-PI")
        if va==1.1:
            if vb==4.4:         return ("PI-CRP-18","1244-016273","CD-PI-05","PI-TOOL-02","CS-150-25-PI")
            if vb==2.9:         return ("PI-CRP-19","1244-016274","CD-PI-05","PI-TOOL-02","CS-150-25-PI")
            if vb==1.8:         return ("PI-CRP-20","1244-016275","CD-PI-05","PI-TOOL-02","CS-150-25-PI")
            if vb==1.1:         return ("PI-CRP-21","1244-016276","CD-PI-05","PI-TOOL-02","CS-150-25-PI")
        if va==0.8:
            if vb==2.9:         return ("PI-CRP-22","1244-016277","CD-PI-06","PI-TOOL-02","CS-150-25-PI")
            if vb==1.8:         return ("PI-CRP-23","1244-016278","CD-PI-06","PI-TOOL-02","CS-150-25-PI")
            if vb==1.1:         return ("PI-CRP-24","1244-016279","CD-PI-06","PI-TOOL-02","CS-150-25-PI")
    return None

def _compat_for(cable1):
    all_names = [c[0] for c in _CABLES_RAW]
    return [c2 for c2 in all_names if _sleeve(cable1, c2)]

# ════════════════════════════════════════════════════════════════════════════
#  ПАЛИТРА И КОНСТАНТЫ
# ════════════════════════════════════════════════════════════════════════════
C = {
    "bg":       "#1C2B3A",
    "panel":    "#F8F5F0",
    "teal":     "#1A6B72",
    "teal_lt":  "#E8F6F7",
    "orange":   "#D9600A",
    "amber":    "#8B5E00",
    "amber_lt": "#FEF8ED",
    "blue":     "#2F5597",
    "blue_lt":  "#E8EEF7",
    "green":    "#1A6B40",
    "green_lt": "#E8F5EE",
    "warn":     "#FFFAE8",
    "warn_txt": "#7A4800",
    "grey":     "#5A7A8A",
    "white":    "#FFFFFF",
    "text":     "#1C2B3A",
    "border":   "#D0D8E0",
}

# ════════════════════════════════════════════════════════════════════════════
#  ГЛАВНОЕ ОКНО
# ════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Raychem XPI — Подбор гильз и муфт")
        self.configure(bg=C["bg"])
        self.resizable(True, True)
        self.minsize(720, 700)

        # ── fonts ────────────────────────────────────────────────────────────
        self.f_title  = tkfont.Font(family="Segoe UI", size=16, weight="bold")
        self.f_sub    = tkfont.Font(family="Segoe UI", size=8)
        self.f_sec    = tkfont.Font(family="Segoe UI", size=9,  weight="bold")
        self.f_label  = tkfont.Font(family="Segoe UI", size=9)
        self.f_result = tkfont.Font(family="Segoe UI", size=14, weight="bold")
        self.f_med    = tkfont.Font(family="Segoe UI", size=11)
        self.f_small  = tkfont.Font(family="Segoe UI", size=8)

        # ── state ────────────────────────────────────────────────────────────
        self.cable1_var = tk.StringVar()
        self.cable2_var = tk.StringVar()
        self.volt_var   = tk.StringVar(value="230")
        self.len1_var   = tk.StringVar(value="1000")
        self.len2_var   = tk.StringVar(value="1000")

        self.cable1_var.trace_add("write", self._on_cable1_change)
        self.cable2_var.trace_add("write", self._on_any_change)
        self.volt_var.trace_add("write",   self._on_any_change)
        self.len1_var.trace_add("write",   self._on_any_change)
        self.len2_var.trace_add("write",   self._on_any_change)

        self._build_ui()
        self._on_cable1_change()

    # ── UI builder ───────────────────────────────────────────────────────────
    def _build_ui(self):
        # Dark top header
        hdr = tk.Frame(self, bg=C["bg"], pady=12)
        hdr.pack(fill="x", padx=0)
        tk.Label(hdr, text="  ПОДБОР ГИЛЬЗ И МУФТ  ·  КАБЕЛИ XPI",
                 font=self.f_title, bg=C["bg"], fg=C["white"]).pack(side="left", padx=16)
        tk.Label(hdr, text="Raychem  ·  PI-TOOL-SET-xx  ·  Chemelex",
                 font=self.f_sub,  bg=C["bg"], fg=C["grey"]).pack(side="right", padx=16)

        # Teal accent line
        tk.Frame(self, bg=C["teal"], height=3).pack(fill="x")

        # Scrollable main area
        container = tk.Frame(self, bg=C["panel"])
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=C["panel"], highlightthickness=0)
        scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self.inner = tk.Frame(canvas, bg=C["panel"])
        win_id = canvas.create_window((0,0), window=self.inner, anchor="nw")

        def _resize(e):
            canvas.itemconfig(win_id, width=e.width)
        canvas.bind("<Configure>", _resize)
        self.inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        self._build_section_01()
        self._build_section_02()
        self._build_section_03()
        self._build_section_04()

        # Bottom strip
        tk.Frame(self, bg=C["teal"], height=4).pack(fill="x", side="bottom")

    def _section_header(self, parent, num, text, color):
        f = tk.Frame(parent, bg=color, pady=6)
        f.pack(fill="x", padx=0, pady=(12,0))
        tk.Label(f, text=f"  {num}  /  {text}",
                 font=self.f_sec, bg=color, fg=C["white"]).pack(side="left", padx=8)
        return f

    def _row(self, parent, label, widget_fn, bg=None):
        bg = bg or C["panel"]
        f = tk.Frame(parent, bg=bg, pady=3)
        f.pack(fill="x", padx=16, pady=1)
        tk.Label(f, text=label, font=self.f_label,
                 bg=bg, fg=C["grey"], width=20, anchor="w").pack(side="left")
        w = widget_fn(f)
        w.pack(side="left", fill="x", expand=True, padx=(4,0))
        return f, w

    # ── SECTION 01: cable selection ──────────────────────────────────────────
    def _build_section_01(self):
        self._section_header(self.inner, "01", "ВЫБОР КАБЕЛЕЙ", C["teal"])

        all_names = [c[0] for c in _CABLES_RAW]

        outer = tk.Frame(self.inner, bg=C["panel"])
        outer.pack(fill="x", padx=16, pady=8)

        # Cable 1
        c1f = tk.LabelFrame(outer, text="▼  КАБЕЛЬ 1  (ОТ)",
                             font=self.f_sec, bg=C["teal_lt"],
                             fg=C["teal"], bd=1, relief="solid", pady=6, padx=8)
        c1f.pack(side="left", fill="both", expand=True, padx=(0,6))

        self.cb1 = ttk.Combobox(c1f, textvariable=self.cable1_var,
                                 values=all_names, state="readonly", width=22)
        self.cb1.pack(fill="x", pady=2)

        self.lbl_c1_info = tk.Label(c1f, text="", font=self.f_small,
                                    bg=C["teal_lt"], fg=C["grey"])
        self.lbl_c1_info.pack(anchor="w")

        # Arrow
        tk.Label(outer, text="→", font=tkfont.Font(size=18),
                 bg=C["panel"], fg=C["grey"]).pack(side="left", padx=8)

        # Cable 2
        c2f = tk.LabelFrame(outer, text="▼  КАБЕЛЬ 2  (ДО)",
                             font=self.f_sec, bg=C["teal_lt"],
                             fg=C["teal"], bd=1, relief="solid", pady=6, padx=8)
        c2f.pack(side="left", fill="both", expand=True, padx=(6,0))

        self.cb2 = ttk.Combobox(c2f, textvariable=self.cable2_var,
                                 values=all_names, state="readonly", width=22)
        self.cb2.pack(fill="x", pady=2)

        self.lbl_c2_info = tk.Label(c2f, text="", font=self.f_small,
                                    bg=C["teal_lt"], fg=C["grey"])
        self.lbl_c2_info.pack(anchor="w")

    # ── SECTION 02: sleeve result ────────────────────────────────────────────
    def _build_section_02(self):
        self._section_header(self.inner, "02", "РЕЗУЛЬТАТ  —  ОСНОВНАЯ ГИЛЬЗА И НАБОР", C["teal"])

        f = tk.Frame(self.inner, bg=C["green_lt"], bd=1, relief="solid")
        f.pack(fill="x", padx=16, pady=6)

        self.lbl_sleeve = tk.Label(f, text="—", font=self.f_result,
                                   bg=C["green_lt"], fg=C["green"], pady=8)
        self.lbl_sleeve.pack()

        details = tk.Frame(self.inner, bg=C["panel"])
        details.pack(fill="x", padx=16)

        rows = [("Номер по каталогу", "lbl_cat"),
                ("Обжимная матрица",  "lbl_mat"),
                ("Инструмент",        "lbl_tool")]
        for lbl, attr in rows:
            rf = tk.Frame(details, bg=C["teal_lt"], pady=4)
            rf.pack(fill="x", pady=1)
            tk.Label(rf, text=lbl, font=self.f_label, bg=C["teal_lt"],
                     fg=C["grey"], width=22, anchor="w").pack(side="left", padx=8)
            w = tk.Label(rf, text="—", font=self.f_med,
                         bg=C["white"], fg=C["text"], anchor="center", width=30)
            w.pack(side="left", fill="x", expand=True, padx=(0,8), pady=2)
            setattr(self, attr, w)

        kit_f = tk.Frame(self.inner, bg=C["panel"], pady=2)
        kit_f.pack(fill="x", padx=16)
        tk.Label(kit_f, text="НАБОР / МУФТА", font=self.f_sec,
                 bg=C["teal_lt"], fg=C["teal"], width=22, anchor="w",
                 pady=6).pack(side="left", padx=8)
        self.lbl_kit = tk.Label(kit_f, text="—", font=self.f_result,
                                bg="#FEF0E6", fg=C["orange"], anchor="center", width=30)
        self.lbl_kit.pack(side="left", fill="x", expand=True, padx=(0,8), pady=2)

    # ── SECTION 03: braid sleeve ─────────────────────────────────────────────
    def _build_section_03(self):
        self._section_header(self.inner, "03", "ГИЛЬЗА ДЛЯ ОПЛЁТКИ  (входит в набор CS-150-xx-PI)", C["amber"])

        f = tk.Frame(self.inner, bg=C["amber_lt"], bd=1, relief="solid")
        f.pack(fill="x", padx=16, pady=6)

        self.lbl_braid = tk.Label(f, text="—", font=self.f_result,
                                  bg=C["amber_lt"], fg=C["amber"], pady=8)
        self.lbl_braid.pack()

        details = tk.Frame(self.inner, bg=C["panel"])
        details.pack(fill="x", padx=16)

        rows = [("Номер по каталогу", "lbl_br_cat"),
                ("Матрица",           "lbl_br_mat"),
                ("Инструмент",        "lbl_br_tool")]
        for lbl, attr in rows:
            rf = tk.Frame(details, bg=C["amber_lt"], pady=4)
            rf.pack(fill="x", pady=1)
            tk.Label(rf, text=lbl, font=self.f_label, bg=C["amber_lt"],
                     fg=C["grey"], width=22, anchor="w").pack(side="left", padx=8)
            w = tk.Label(rf, text="—", font=self.f_med,
                         bg=C["white"], fg=C["amber"], anchor="center", width=30)
            w.pack(side="left", fill="x", expand=True, padx=(0,8), pady=2)
            setattr(self, attr, w)

    # ── SECTION 04: power ────────────────────────────────────────────────────
    def _build_section_04(self):
        self._section_header(self.inner, "04", "МОЩНОСТЬ НАГРЕВА  —  ВТ/М", C["blue"])

        inp_frame = tk.Frame(self.inner, bg=C["blue_lt"], pady=6)
        inp_frame.pack(fill="x", padx=16, pady=(4,2))

        def inp_row(parent, label, var, unit):
            f = tk.Frame(parent, bg=C["blue_lt"])
            f.pack(fill="x", padx=8, pady=2)
            tk.Label(f, text=label, font=self.f_label, bg=C["blue_lt"],
                     fg=C["grey"], width=20, anchor="w").pack(side="left")
            e = tk.Entry(f, textvariable=var, font=self.f_med, width=10,
                         bg="#FFFFF0", fg=C["blue"], relief="solid",
                         bd=1, justify="center")
            e.pack(side="left", padx=4)
            tk.Label(f, text=unit, font=self.f_label, bg=C["blue_lt"],
                     fg=C["grey"]).pack(side="left")

        inp_row(inp_frame, "Напряжение",      self.volt_var, "В")
        inp_row(inp_frame, "Длина кабеля 1",  self.len1_var, "м")
        inp_row(inp_frame, "Длина кабеля 2",  self.len2_var, "м")

        res_frame = tk.Frame(self.inner, bg=C["panel"])
        res_frame.pack(fill="x", padx=16, pady=(2,12))

        rows = [("Вт/м  (кабель 1)", "lbl_wm1", C["blue"], "#D0DCF0", True),
                ("Вт/м  (кабель 2)", "lbl_wm2", C["text"], C["white"], False),
                ("Мощность цепи",    "lbl_pwr",  C["text"], C["white"], False)]
        for lbl, attr, fg, bg, bold in rows:
            rf = tk.Frame(res_frame, bg=C["blue_lt"], pady=4)
            rf.pack(fill="x", pady=1)
            tk.Label(rf, text=lbl, font=self.f_label, bg=C["blue_lt"],
                     fg=C["grey"], width=22, anchor="w").pack(side="left", padx=8)
            fn = self.f_result if bold else self.f_med
            w = tk.Label(rf, text="—", font=fn, bg=bg, fg=fg,
                         anchor="center", width=30)
            w.pack(side="left", fill="x", expand=True, padx=(0,8), pady=2)
            setattr(self, attr, w)

    # ── LOGIC ────────────────────────────────────────────────────────────────
    def _on_cable1_change(self, *_):
        c1 = self.cable1_var.get()
        if c1:
            opts = _compat_for(c1)
            self.cb2.configure(values=opts)
            if self.cable2_var.get() not in opts:
                self.cable2_var.set(opts[0] if opts else "")
        self._refresh()

    def _on_any_change(self, *_):
        self._refresh()

    def _refresh(self):
        c1 = self.cable1_var.get()
        c2 = self.cable2_var.get()

        # Cable info labels
        def cable_info(name):
            for row in _CABLES_RAW:
                if row[0] == name:
                    return f"Сечение: {row[1]} мм²   Сопр.: {row[2]} Ом/км"
            return ""
        self.lbl_c1_info.config(text=cable_info(c1))
        self.lbl_c2_info.config(text=cable_info(c2))

        if not c1 or not c2:
            self._clear_results()
            return

        # Sleeve
        sl = _sleeve(c1, c2)
        if sl:
            sleeve, cat, mat, tool, kit = sl
            self.lbl_sleeve.config(text=sleeve, fg=C["green"])
            self.lbl_cat.config(text=cat)
            self.lbl_mat.config(text=mat)
            self.lbl_tool.config(text=tool)
            self.lbl_kit.config(text=kit, fg=C["orange"])

            # Braid
            fam = _family(c1)
            bk = f"{kit}|{fam}"
            br = _BRAID.get(bk)
            if br:
                self.lbl_braid.config(text=br[0], fg=C["amber"])
                self.lbl_br_cat.config(text=br[1])
                self.lbl_br_mat.config(text=br[2])
                self.lbl_br_tool.config(text=br[3])
            else:
                self._clear_braid()
        else:
            self.lbl_sleeve.config(text="⚠  Нет совместимой гильзы", fg="#B00000")
            for w in [self.lbl_cat, self.lbl_mat, self.lbl_tool]:
                w.config(text="—")
            self.lbl_kit.config(text="—", fg=C["grey"])
            self._clear_braid()

        # Power
        try:
            U  = float(self.volt_var.get())
            L1 = float(self.len1_var.get())
            L2 = float(self.len2_var.get())
            r1 = next(c[2] for c in _CABLES_RAW if c[0]==c1)
            r2 = next(c[2] for c in _CABLES_RAW if c[0]==c2)
            wm1 = round(U**2 * 1000 / (r1 * L1**2), 1) if L1 else 0
            wm2 = round(U**2 * 1000 / (r2 * L2**2), 1) if L2 else 0
            pwr = round(U**2 / (r1*L1 + r2*L2), 2) if (r1*L1+r2*L2) else 0
            self.lbl_wm1.config(text=f"{wm1} Вт/м")
            self.lbl_wm2.config(text=f"{wm2} Вт/м")
            self.lbl_pwr.config(text=f"{pwr} кВт")
        except Exception:
            self.lbl_wm1.config(text="—")
            self.lbl_wm2.config(text="—")
            self.lbl_pwr.config(text="—")

    def _clear_results(self):
        for w in [self.lbl_sleeve, self.lbl_cat, self.lbl_mat,
                  self.lbl_tool,  self.lbl_kit,
                  self.lbl_wm1,   self.lbl_wm2, self.lbl_pwr]:
            w.config(text="—")
        self._clear_braid()

    def _clear_braid(self):
        for w in [self.lbl_braid, self.lbl_br_cat,
                  self.lbl_br_mat, self.lbl_br_tool]:
            w.config(text="—")


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    # Center window
    app.update_idletasks()
    w, h = 820, 780
    x = (app.winfo_screenwidth()  - w) // 2
    y = (app.winfo_screenheight() - h) // 2
    app.geometry(f"{w}x{h}+{x}+{y}")
    app.mainloop()
