"""
DocFlow Pro — 共享 Widgets 库
"""
import tkinter as tk
from tkinter import ttk
from gui.theme import COLORS, FONTS, PADDING


def apply_global_styles(root):
    """应用全局 ttk 样式"""
    s = ttk.Style(root)
    s.theme_use("clam")

    # 进度条
    s.configure("Gold.Horizontal.TProgressbar",
                 troughcolor=COLORS["bg_input"],
                 background=COLORS["accent"],
                 bordercolor=COLORS["border"],
                 lightcolor=COLORS["accent"],
                 darkcolor=COLORS["accent_dim"])

    s.configure("Blue.Horizontal.TProgressbar",
                 troughcolor=COLORS["bg_input"],
                 background=COLORS["tool1_color"],
                 bordercolor=COLORS["border"],
                 lightcolor=COLORS["tool1_color"],
                 darkcolor=COLORS["tool1_color"])

    s.configure("Green.Horizontal.TProgressbar",
                 troughcolor=COLORS["bg_input"],
                 background=COLORS["tool2_color"],
                 bordercolor=COLORS["border"],
                 lightcolor=COLORS["tool2_color"],
                 darkcolor=COLORS["tool2_color"])

    s.configure("Orange.Horizontal.TProgressbar",
                 troughcolor=COLORS["bg_input"],
                 background=COLORS["tool3_color"],
                 bordercolor=COLORS["border"],
                 lightcolor=COLORS["tool3_color"],
                 darkcolor=COLORS["tool3_color"])

    # Scrollbar
    s.configure("Dark.Vertical.TScrollbar",
                 background=COLORS["bg_hover"],
                 troughcolor=COLORS["bg_main"],
                 arrowcolor=COLORS["text_muted"],
                 bordercolor=COLORS["border"],
                 relief="flat")

    # Combobox
    s.configure("Dark.TCombobox",
                 fieldbackground=COLORS["bg_input"],
                 background=COLORS["bg_hover"],
                 foreground=COLORS["text_primary"],
                 arrowcolor=COLORS["text_secondary"],
                 bordercolor=COLORS["border"],
                 selectbackground=COLORS["bg_active"],
                 selectforeground=COLORS["text_primary"])


class Divider(tk.Frame):
    def __init__(self, parent, color=None, **kw):
        c = color or COLORS["border"]
        super().__init__(parent, bg=c, height=1, **kw)


class IconButton(tk.Label):
    """图标按钮（用 Label 实现，无边框）"""
    def __init__(self, parent, text, command=None,
                 fg=None, hover_fg=None, font=None, **kw):
        self._fg = fg or COLORS["text_secondary"]
        self._hover_fg = hover_fg or COLORS["accent"]
        self._cmd = command
        bg = kw.pop("bg", parent.cget("bg"))
        super().__init__(parent, text=text, bg=bg,
                         fg=self._fg,
                         font=font or FONTS["md"],
                         cursor="hand2", **kw)
        self.bind("<Enter>", lambda e: self.config(fg=self._hover_fg))
        self.bind("<Leave>", lambda e: self.config(fg=self._fg))
        self.bind("<Button-1>", lambda e: command() if command else None)


class StyledButton(tk.Frame):
    """统一样式按钮"""
    STYLES = {
        "primary": {
            "bg": COLORS["accent"], "fg": "#0D1117",
            "hover": COLORS["accent_bright"],
        },
        "secondary": {
            "bg": COLORS["bg_hover"], "fg": COLORS["text_primary"],
            "hover": COLORS["bg_active"],
        },
        "danger": {
            "bg": COLORS["error"], "fg": "#FFFFFF",
            "hover": "#C03040",
        },
        "ghost": {
            "bg": COLORS["bg_card"], "fg": COLORS["accent"],
            "hover": COLORS["bg_hover"],
        },
        "blue": {
            "bg": COLORS["tool1_color"], "fg": "#0D1117",
            "hover": "#79BCFF",
        },
        "green": {
            "bg": COLORS["tool2_color"], "fg": "#0D1117",
            "hover": "#5FD975",
        },
        "orange": {
            "bg": COLORS["tool3_color"], "fg": "#0D1117",
            "hover": "#FF9F5E",
        },
    }

    def __init__(self, parent, text="", style="primary",
                 command=None, width=None, padx=None, pady=None, **kw):
        s = self.STYLES.get(style, self.STYLES["primary"])
        bg = kw.pop("bg", parent.cget("bg"))
        super().__init__(parent, bg=bg, **kw)

        btn_kw = dict(
            text=text, bg=s["bg"], fg=s["fg"],
            activebackground=s["hover"],
            activeforeground=s["fg"],
            relief="flat", bd=0, cursor="hand2",
            font=FONTS["md"],
            padx=padx if padx is not None else PADDING["lg"],
            pady=pady if pady is not None else PADDING["sm"],
            command=command,
        )
        if width:
            btn_kw["width"] = width

        self._btn = tk.Button(self, **btn_kw)
        self._btn.pack(fill="both", expand=True)
        self._s = s
        self._btn.bind("<Enter>", lambda e: self._btn.config(bg=s["hover"]))
        self._btn.bind("<Leave>", lambda e: self._btn.config(bg=s["bg"]))

    def config(self, **kw):
        if "state" in kw:
            self._btn.config(state=kw.pop("state"))
        if "text" in kw:
            self._btn.config(text=kw.pop("text"))
        if kw:
            super().config(**kw)

    def set_text(self, t):
        self._btn.config(text=t)

    def set_enabled(self, v):
        if v:
            self._btn.config(state="normal", bg=self._s["bg"],
                             cursor="hand2", fg=self._s["fg"])
        else:
            self._btn.config(state="disabled",
                             bg=COLORS["bg_hover"],
                             cursor="", fg=COLORS["text_muted"])


class LabeledEntry(tk.Frame):
    """带标签的输入框行"""
    def __init__(self, parent, label, width=40, show=None,
                 default="", read_only=False, **kw):
        bg = kw.pop("bg", COLORS["bg_card"])
        super().__init__(parent, bg=bg, **kw)

        tk.Label(self, text=label, bg=bg,
                 fg=COLORS["text_secondary"],
                 font=FONTS["sm"], anchor="w",
                 width=12).pack(side="left")

        entry_kw = dict(
            bg=COLORS["bg_input"], fg=COLORS["text_primary"],
            insertbackground=COLORS["text_primary"],
            relief="flat", bd=0,
            font=FONTS["mono_md"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["border_focus"],
            width=width,
        )
        if show:
            entry_kw["show"] = show

        self._var = tk.StringVar(value=default)
        self._entry = tk.Entry(self, textvariable=self._var, **entry_kw)
        self._entry.pack(side="left", padx=(6, 0), ipady=4, fill="x", expand=True)

        if read_only:
            self._entry.config(state="readonly",
                                readonlybackground=COLORS["bg_input"])

    def get(self):
        return self._var.get()

    def set(self, val):
        self._var.set(val)

    def set_show(self, show):
        self._entry.config(show=show)


class LogView(tk.Frame):
    """日志文本框"""
    def __init__(self, parent, height=10, **kw):
        bg = kw.pop("bg", COLORS["bg_deep"])
        super().__init__(parent, bg=bg, **kw)

        sb = ttk.Scrollbar(self, orient="vertical",
                            style="Dark.Vertical.TScrollbar")
        sb.pack(side="right", fill="y")

        self._text = tk.Text(
            self, height=height,
            bg=COLORS["bg_deep"], fg="#A8B8C8",
            font=FONTS["mono_sm"],
            relief="flat", bd=0,
            wrap="word",
            insertbackground=COLORS["text_primary"],
            selectbackground=COLORS["bg_active"],
            state="disabled",
            yscrollcommand=sb.set,
            padx=8, pady=6,
        )
        self._text.pack(side="left", fill="both", expand=True)
        sb.config(command=self._text.yview)

        # 颜色标签
        self._text.tag_config("INFO",    foreground="#79C0FF")
        self._text.tag_config("SUCCESS", foreground="#3FB950")
        self._text.tag_config("WARNING", foreground="#D29922")
        self._text.tag_config("ERROR",   foreground="#F85149")
        self._text.tag_config("DIM",     foreground="#484F58")

        # 滚轮支持
        import sys as _sys
        def _scroll_log(event):
            if _sys.platform == "win32":
                self._text.yview_scroll(int(-1*(event.delta/120)), "units")
            elif _sys.platform == "darwin":
                self._text.yview_scroll(int(-1*event.delta), "units")
            else:
                self._text.yview_scroll(-1 if event.num==4 else 1, "units")
        if _sys.platform == "linux":
            self._text.bind("<Button-4>", _scroll_log)
            self._text.bind("<Button-5>", _scroll_log)
        else:
            self._text.bind("<MouseWheel>", _scroll_log)
        self._text.tag_config("GOLD",    foreground="#C8A96E")

    def append(self, msg, level="INFO"):
        from datetime import datetime
        ts = datetime.now().strftime("%H:%M:%S")

        def _do():
            self._text.config(state="normal")
            self._text.insert("end", f"[{ts}] ", "DIM")
            self._text.insert("end", msg + "\n", level)
            self._text.see("end")
            self._text.config(state="disabled")

        try:
            self._text.after(0, _do)
        except Exception:
            pass

    def clear(self):
        self._text.config(state="normal")
        self._text.delete("1.0", "end")
        self._text.config(state="disabled")


class ProgressRow(tk.Frame):
    """进度条行（标签 + 进度条 + 百分比）"""
    def __init__(self, parent, bar_style="Gold.Horizontal.TProgressbar", **kw):
        bg = kw.pop("bg", COLORS["bg_card"])
        super().__init__(parent, bg=bg, **kw)

        self._label = tk.Label(self, text="", bg=bg,
                                fg=COLORS["text_secondary"],
                                font=FONTS["sm"], anchor="w")
        self._label.pack(fill="x", pady=(0, 2))

        bar_row = tk.Frame(self, bg=bg)
        bar_row.pack(fill="x")

        self._var = tk.DoubleVar(value=0)
        self._bar = ttk.Progressbar(
            bar_row, variable=self._var, maximum=100,
            style=bar_style, mode="determinate",
        )
        self._bar.pack(side="left", fill="x", expand=True)

        self._pct = tk.Label(bar_row, text="0%", bg=bg,
                              fg=COLORS["text_secondary"],
                              font=FONTS["xs"], width=5)
        self._pct.pack(side="right", padx=(6, 0))

    def update(self, value, label=None):
        def _do():
            self._var.set(value)
            self._pct.config(text=f"{int(value)}%")
            if label is not None:
                self._label.config(text=label)
        try:
            self._bar.after(0, _do)
        except Exception:
            pass

    def reset(self, label=""):
        self.update(0, label)


class SectionHeader(tk.Frame):
    """带色条的节标题"""
    def __init__(self, parent, title, color=None, icon="", **kw):
        bg = kw.pop("bg", COLORS["bg_card"])
        super().__init__(parent, bg=bg, **kw)

        c = color or COLORS["accent"]
        # 左侧色条
        tk.Frame(self, bg=c, width=3).pack(side="left", fill="y")

        inner = tk.Frame(self, bg=bg)
        inner.pack(side="left", padx=(8, 0), pady=4)

        if icon:
            tk.Label(inner, text=icon, bg=bg,
                     fg=c, font=FONTS["lg"]).pack(side="left", padx=(0, 6))

        tk.Label(inner, text=title, bg=bg,
                 fg=COLORS["text_primary"],
                 font=FONTS["h3"]).pack(side="left")


class FilePickRow(tk.Frame):
    """文件路径选择行"""
    def __init__(self, parent, label, browse_cmd,
                 placeholder="（未选择）", **kw):
        bg = kw.pop("bg", COLORS["bg_card"])
        super().__init__(parent, bg=bg, **kw)

        self._ph = placeholder

        tk.Label(self, text=label, bg=bg,
                 fg=COLORS["text_secondary"],
                 font=FONTS["sm"], width=9,
                 anchor="w").pack(side="left")

        self._var = tk.StringVar(value=placeholder)
        e = tk.Entry(self, textvariable=self._var,
                     bg=COLORS["bg_input"], fg=COLORS["text_muted"],
                     font=FONTS["mono_sm"], relief="flat", bd=0,
                     highlightthickness=1,
                     highlightbackground=COLORS["border"],
                     highlightcolor=COLORS["border_focus"],
                     readonlybackground=COLORS["bg_input"],
                     state="readonly", width=48)
        e.pack(side="left", padx=(6, 8), ipady=4, fill="x", expand=True)
        self._entry = e

        tk.Button(self, text="浏览…",
                  bg=COLORS["bg_hover"], fg=COLORS["text_secondary"],
                  activebackground=COLORS["accent"],
                  activeforeground="#0D1117",
                  relief="flat", bd=0, cursor="hand2",
                  font=FONTS["sm"], padx=10, pady=4,
                  command=browse_cmd).pack(side="left")

    def set(self, path):
        self._var.set(path)
        self._entry.config(fg=COLORS["text_primary"]
                           if path != self._ph else COLORS["text_muted"])

    def get(self):
        v = self._var.get()
        return "" if v == self._ph else v


class StatusBar(tk.Frame):
    def __init__(self, parent, **kw):
        bg = kw.pop("bg", COLORS["bg_sidebar"])
        super().__init__(parent, bg=bg, height=26, **kw)
        self.pack_propagate(False)

        self._dot = tk.Label(self, text="●", bg=bg,
                              fg=COLORS["success"],
                              font=FONTS["xs"])
        self._dot.pack(side="left", padx=(PADDING["md"], 4))

        self._label = tk.Label(self, text="就绪",
                                bg=bg, fg=COLORS["text_secondary"],
                                font=FONTS["xs"], anchor="w")
        self._label.pack(side="left", fill="x", expand=True)

        tk.Label(self, text="HHPDI v1.0",
                 bg=bg, fg=COLORS["text_muted"],
                 font=FONTS["xs"]).pack(side="right", padx=PADDING["md"])

    def set(self, msg, state="normal"):
        dot_c = {
            "normal": COLORS["text_muted"],
            "running": COLORS["running"],
            "success": COLORS["success"],
            "error": COLORS["error"],
            "warning": COLORS["warning"],
        }.get(state, COLORS["text_muted"])
        self._dot.config(fg=dot_c)
        self._label.config(text=msg)
