"""
DocFlow Pro — 模型设置窗口 v2
双模型配置：VLM（多模态，文档解析）+ LLM（纯文本，数据标注）
"""
import tkinter as tk
from tkinter import ttk, messagebox
import sys

from gui.theme import COLORS, FONTS, PADDING
from gui.widgets import StyledButton, Divider
from config.settings import get_config, update_config

SILICONFLOW_BASE_URL = "https://api.siliconflow.cn/v1"

SILICONFLOW_VLM_MODELS = [
    "Qwen/Qwen2.5-VL-72B-Instruct",
    "Qwen/Qwen2-VL-7B-Instruct",
    "Pro/Qwen/Qwen2-VL-7B-Instruct",
    "Pro/Qwen/Qwen2.5-VL-3B-Instruct",
]

SILICONFLOW_LLM_MODELS = [
    "Pro/deepseek-ai/DeepSeek-V3",
    "deepseek-ai/DeepSeek-V3",
    "Qwen/Qwen2.5-72B-Instruct",
    "Qwen/Qwen2.5-32B-Instruct",
    "Pro/Qwen/Qwen2.5-7B-Instruct",
    "deepseek-ai/DeepSeek-R1",
    "Pro/deepseek-ai/DeepSeek-R1",
]


class SettingsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("模型设置")
        self.geometry("660x980")
        self.resizable(True, False)
        self.configure(bg=COLORS["bg_main"])
        self.grab_set()
        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width() // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        self.geometry(f"+{pw - 320}+{ph - 350}")
        self._cfg = get_config()
        self._build()
        self._load()

    def _build(self):
        hdr = tk.Frame(self, bg=COLORS["bg_card"], height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="模型设置",
                 bg=COLORS["bg_card"], fg=COLORS["text_primary"],
                 font=FONTS["h1"]).pack(side="left", padx=PADDING["xl"])
        Divider(self).pack(fill="x")

        # 底部按钮栏先 pack（side=bottom），body 后 pack（fill remaining）
        Divider(self).pack(side="bottom", fill="x")
        btn_bar_outer = tk.Frame(self, bg=COLORS["bg_main"])
        btn_bar_outer.pack(side="bottom", fill="x",
                           padx=PADDING["xl"], pady=PADDING["md"])
        StyledButton(btn_bar_outer, text="测试连接", style="ghost",
                     command=self._test).pack(side="left")
        StyledButton(btn_bar_outer, text="取消", style="secondary",
                     command=self.destroy).pack(side="right",
                                                padx=(PADDING["sm"], 0))
        StyledButton(btn_bar_outer, text="保存", style="primary",
                     command=self._save).pack(side="right")

        body = tk.Frame(self, bg=COLORS["bg_main"])
        body.pack(fill="both", expand=True, padx=PADDING["xl"], pady=PADDING["lg"])

        self._build_model_block(body, "vlm",
            title="视觉模型  (VLM)",
            subtitle="用于文档解析 — 识别版面、表格、公式",
            color=COLORS["tool1_color"],
            models=SILICONFLOW_VLM_MODELS)

        self._build_fallback_block(body, "vlm",
            color=COLORS["tool1_color"],
            models=SILICONFLOW_VLM_MODELS)

        tk.Frame(body, bg=COLORS["border"], height=1).pack(fill="x", pady=PADDING["lg"])

        self._build_model_block(body, "llm",
            title="文本模型  (LLM)",
            subtitle="用于数据标注 — 打标签、表格转QA",
            color=COLORS["tool3_color"],
            models=SILICONFLOW_LLM_MODELS)

        self._build_fallback_block(body, "llm",
            color=COLORS["tool3_color"],
            models=SILICONFLOW_LLM_MODELS)

        tk.Frame(body, bg=COLORS["border"], height=1).pack(fill="x", pady=PADDING["lg"])
        self._build_parse_options(body)



    def _build_model_block(self, parent, tag, title, subtitle, color, models):
        # 标题
        hdr = tk.Frame(parent, bg=COLORS["bg_main"])
        hdr.pack(fill="x", pady=(0, PADDING["sm"]))
        tk.Frame(hdr, bg=color, width=4).pack(side="left", fill="y")
        col = tk.Frame(hdr, bg=COLORS["bg_main"])
        col.pack(side="left", padx=(PADDING["sm"], 0))
        tk.Label(col, text=title, bg=COLORS["bg_main"], fg=color, font=FONTS["h2"]).pack(anchor="w")
        tk.Label(col, text=subtitle, bg=COLORS["bg_main"], fg=COLORS["text_muted"], font=FONTS["xs"]).pack(anchor="w")

        # 表单
        form = tk.Frame(parent, bg=COLORS["bg_card"])
        form.pack(fill="x")
        inner = tk.Frame(form, bg=COLORS["bg_card"])
        inner.pack(fill="x", padx=PADDING["lg"], pady=PADDING["md"])
        inner.columnconfigure(1, weight=1)

        def lbl(t, r):
            tk.Label(inner, text=t, bg=COLORS["bg_card"], fg=COLORS["text_secondary"],
                     font=FONTS["sm"], anchor="w", width=9).grid(row=r, column=0, sticky="w", pady=5)

        def entry(r, default=""):
            e = tk.Entry(inner, bg=COLORS["bg_input"], fg=COLORS["text_primary"],
                         insertbackground=COLORS["text_primary"],
                         relief="flat", bd=0, highlightthickness=1,
                         highlightbackground=COLORS["border"],
                         highlightcolor=color, font=FONTS["mono_sm"])
            e.grid(row=r, column=1, sticky="ew", padx=(PADDING["sm"],0), pady=5, ipady=6)
            if default:
                e.insert(0, default)
            return e

        lbl("Base URL", 0)
        url_e = entry(0, SILICONFLOW_BASE_URL)

        lbl("API Key", 1)
        key_row = tk.Frame(inner, bg=COLORS["bg_card"])
        key_row.grid(row=1, column=1, sticky="ew", padx=(PADDING["sm"],0), pady=5)
        key_e = tk.Entry(key_row, bg=COLORS["bg_input"], fg=COLORS["text_primary"],
                          insertbackground=COLORS["text_primary"],
                          relief="flat", bd=0, highlightthickness=1,
                          highlightbackground=COLORS["border"],
                          highlightcolor=color, font=FONTS["mono_sm"], show="•")
        key_e.pack(side="left", fill="x", expand=True, ipady=6)
        sv = tk.BooleanVar(value=False)
        tk.Checkbutton(key_row, text="显示", variable=sv,
                       bg=COLORS["bg_card"], fg=COLORS["text_muted"],
                       activebackground=COLORS["bg_card"],
                       selectcolor=COLORS["bg_input"], font=FONTS["xs"],
                       command=lambda e=key_e, v=sv: e.config(show="" if v.get() else "•")
                       ).pack(side="right", padx=(6,0))

        lbl("模型", 2)
        mv = tk.StringVar(value=models[0] if models else "")
        cb = ttk.Combobox(inner, textvariable=mv, values=models,
                           font=FONTS["sm"], style="Dark.TCombobox", width=42)
        cb.grid(row=2, column=1, sticky="ew", padx=(PADDING["sm"],0), pady=5)

        setattr(self, f"_{tag}_url",   url_e)
        setattr(self, f"_{tag}_key",   key_e)
        setattr(self, f"_{tag}_model", mv)

    def _build_fallback_block(self, parent, tag, color, models):
        """备选服务商折叠区块（主服务商全部重试失败后自动切换）"""
        fb_enabled = tk.BooleanVar(value=False)

        # 折叠头部
        header = tk.Frame(parent, bg=COLORS["bg_main"])
        header.pack(fill="x", pady=(PADDING["sm"], 0))

        arrow_var = tk.StringVar(value="▶")
        content_frame = tk.Frame(parent, bg=COLORS["bg_card"])

        def _toggle():
            if content_frame.winfo_viewable():
                content_frame.pack_forget()
                arrow_var.set("▶")
            else:
                content_frame.pack(fill="x")
                arrow_var.set("▼")

        tk.Label(header, textvariable=arrow_var,
                 bg=COLORS["bg_main"], fg=COLORS["text_muted"],
                 font=FONTS["xs"], cursor="hand2").pack(side="left")
        toggle_lbl = tk.Label(header,
                 text=f"备选服务商（{tag.upper()} 主服务商失败时自动切换）",
                 bg=COLORS["bg_main"], fg=COLORS["text_muted"],
                 font=FONTS["xs"], cursor="hand2")
        toggle_lbl.pack(side="left", padx=(4, 0))
        toggle_lbl.bind("<Button-1>", lambda e: _toggle())
        header.bind("<Button-1>", lambda e: _toggle())

        # 内容区（默认折叠）
        inner = tk.Frame(content_frame, bg=COLORS["bg_card"])
        inner.pack(fill="x", padx=PADDING["lg"], pady=PADDING["md"])
        inner.columnconfigure(1, weight=1)

        def lbl(t, r):
            tk.Label(inner, text=t, bg=COLORS["bg_card"], fg=COLORS["text_muted"],
                     font=FONTS["sm"], anchor="w", width=9).grid(
                         row=r, column=0, sticky="w", pady=4)

        def entry(r, show=""):
            e = tk.Entry(inner, bg=COLORS["bg_input"], fg=COLORS["text_primary"],
                         insertbackground=COLORS["text_primary"],
                         relief="flat", bd=0, highlightthickness=1,
                         highlightbackground=COLORS["border"],
                         highlightcolor=color, font=FONTS["mono_sm"], show=show)
            e.grid(row=r, column=1, sticky="ew",
                   padx=(PADDING["sm"], 0), pady=4, ipady=5)
            return e

        # 启用开关
        tk.Checkbutton(inner, text="启用备选服务商",
                       variable=fb_enabled,
                       bg=COLORS["bg_card"], fg=COLORS["text_primary"],
                       activebackground=COLORS["bg_card"],
                       selectcolor=COLORS["bg_input"],
                       font=FONTS["sm"]).grid(
                           row=0, column=0, columnspan=2,
                           sticky="w", pady=(0, 6))

        lbl("Base URL", 1)
        fb_url_e = entry(1)

        lbl("API Key", 2)
        key_row = tk.Frame(inner, bg=COLORS["bg_card"])
        key_row.grid(row=2, column=1, sticky="ew",
                     padx=(PADDING["sm"], 0), pady=4)
        fb_key_e = tk.Entry(key_row, bg=COLORS["bg_input"], fg=COLORS["text_primary"],
                             insertbackground=COLORS["text_primary"],
                             relief="flat", bd=0, highlightthickness=1,
                             highlightbackground=COLORS["border"],
                             highlightcolor=color, font=FONTS["mono_sm"], show="•")
        fb_key_e.pack(side="left", fill="x", expand=True, ipady=5)
        sv = tk.BooleanVar(value=False)
        tk.Checkbutton(key_row, text="显示", variable=sv,
                       bg=COLORS["bg_card"], fg=COLORS["text_muted"],
                       activebackground=COLORS["bg_card"],
                       selectcolor=COLORS["bg_input"], font=FONTS["xs"],
                       command=lambda e=fb_key_e, v=sv:
                           e.config(show="" if v.get() else "•")
                       ).pack(side="right", padx=(6, 0))

        lbl("模型", 3)
        fb_mv = tk.StringVar()
        fb_cb = ttk.Combobox(inner, textvariable=fb_mv, values=models,
                              font=FONTS["sm"], style="Dark.TCombobox", width=42)
        fb_cb.grid(row=3, column=1, sticky="ew",
                   padx=(PADDING["sm"], 0), pady=4)

        setattr(self, f"_{tag}_fb_enabled", fb_enabled)
        setattr(self, f"_{tag}_fb_url",     fb_url_e)
        setattr(self, f"_{tag}_fb_key",     fb_key_e)
        setattr(self, f"_{tag}_fb_model",   fb_mv)

    def _build_parse_options(self, parent):
        tk.Label(parent, text="解析选项", bg=COLORS["bg_main"],
                 fg=COLORS["text_secondary"], font=FONTS["h3"]).pack(anchor="w", pady=(0, PADDING["sm"]))
        form = tk.Frame(parent, bg=COLORS["bg_card"])
        form.pack(fill="x")
        inner = tk.Frame(form, bg=COLORS["bg_card"])
        inner.pack(fill="x", padx=PADDING["lg"], pady=PADDING["md"])
        opts = self._cfg.get("parse_options", {})
        self._opt_tables   = tk.BooleanVar(value=opts.get("extract_tables_as_md", True))
        self._opt_formulas = tk.BooleanVar(value=opts.get("extract_formulas", True))
        self._opt_bbox     = tk.BooleanVar(value=opts.get("add_bbox_comments", True))
        for var, text in [
            (self._opt_tables,   "表格识别为 Markdown 格式"),
            (self._opt_formulas, "公式识别为 LaTeX"),
            (self._opt_bbox,     "在 Markdown 中添加位置注释"),
        ]:
            tk.Checkbutton(inner, text=text, variable=var,
                           bg=COLORS["bg_card"], fg=COLORS["text_primary"],
                           activebackground=COLORS["bg_card"],
                           selectcolor=COLORS["bg_input"],
                           font=FONTS["sm"]).pack(anchor="w", pady=3)

    def _load(self):
        cfg = self._cfg
        vlm = cfg.get("vlm", {})
        self._set(self._vlm_url, vlm.get("base_url", SILICONFLOW_BASE_URL))
        self._set(self._vlm_key, vlm.get("api_key", ""))
        if vlm.get("model"): self._vlm_model.set(vlm["model"])

        llm = cfg.get("llm", {})
        self._set(self._llm_url, llm.get("base_url", SILICONFLOW_BASE_URL))
        self._set(self._llm_key, llm.get("api_key", ""))
        if llm.get("model"): self._llm_model.set(llm["model"])

        vfb = cfg.get("vlm_fallback", {})
        self._vlm_fb_enabled.set(vfb.get("enabled", False))
        self._set(self._vlm_fb_url, vfb.get("base_url", ""))
        self._set(self._vlm_fb_key, vfb.get("api_key", ""))
        if vfb.get("model"): self._vlm_fb_model.set(vfb["model"])

        lfb = cfg.get("llm_fallback", {})
        self._llm_fb_enabled.set(lfb.get("enabled", False))
        self._set(self._llm_fb_url, lfb.get("base_url", ""))
        self._set(self._llm_fb_key, lfb.get("api_key", ""))
        if lfb.get("model"): self._llm_fb_model.set(lfb["model"])

    def _set(self, entry, val):
        entry.delete(0, "end")
        entry.config(fg=COLORS["text_primary"])
        if val: entry.insert(0, val)

    def _save(self):
        updates = {
            "vlm": {
                "base_url": self._vlm_url.get().strip() or SILICONFLOW_BASE_URL,
                "api_key":  self._vlm_key.get().strip(),
                "model":    self._vlm_model.get().strip(),
            },
            "vlm_fallback": {
                "enabled":  self._vlm_fb_enabled.get(),
                "base_url": self._vlm_fb_url.get().strip(),
                "api_key":  self._vlm_fb_key.get().strip(),
                "model":    self._vlm_fb_model.get().strip(),
            },
            "llm": {
                "base_url": self._llm_url.get().strip() or SILICONFLOW_BASE_URL,
                "api_key":  self._llm_key.get().strip(),
                "model":    self._llm_model.get().strip(),
            },
            "llm_fallback": {
                "enabled":  self._llm_fb_enabled.get(),
                "base_url": self._llm_fb_url.get().strip(),
                "api_key":  self._llm_fb_key.get().strip(),
                "model":    self._llm_fb_model.get().strip(),
            },
            "model_mode": "online",
            "online": {
                "provider": "siliconflow",
                "base_url": self._vlm_url.get().strip() or SILICONFLOW_BASE_URL,
                "api_key":  self._vlm_key.get().strip(),
                "model":    self._vlm_model.get().strip(),
            },
            "parse_options": {
                "extract_tables_as_md": self._opt_tables.get(),
                "extract_formulas":     self._opt_formulas.get(),
                "add_bbox_comments":    self._opt_bbox.get(),
            },
        }
        update_config(updates)
        messagebox.showinfo("已保存", "模型设置已保存", parent=self)
        self.destroy()

    def _test(self):
        self._save()
        try:
            from core.vlm_client import test_connection
            ok, msg = test_connection()
            if ok:
                messagebox.showinfo("连接测试", f"✅ {msg}", parent=self)
            else:
                messagebox.showerror("连接测试", f"❌ {msg}", parent=self)
        except Exception as e:
            messagebox.showerror("连接测试", str(e), parent=self)
