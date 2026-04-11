"""
DocFlow Pro — 主窗口
侧边栏导航 + 多工具面板切换
"""
import sys
import os
import tkinter as tk
from tkinter import messagebox

from gui.theme import COLORS, FONTS, PADDING, NAV_ITEMS
from gui.widgets import Divider, StatusBar, apply_global_styles


class DocFlowProApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HHPDI文档数据处理智能助手")
        self.geometry("1200x760")
        self.minsize(900, 600)
        self.configure(bg=COLORS["bg_main"])

        apply_global_styles(self)

        # 跨工具共享状态
        self._shared = {
            "last_md_path": "",
            "last_images_dir": "",
        }

        self._panels = {}
        self._active_nav = None
        self._nav_btns   = {}

        self._build()
        self._navigate("home")
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = 1200, 760
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ── UI 构建 ───────────────────────────────────────────────

    def _build(self):
        # 主布局：侧边栏 | 内容区
        main = tk.Frame(self, bg=COLORS["bg_main"])
        main.pack(fill="both", expand=True)

        self._build_sidebar(main)
        self._build_content(main)

        # 底部状态栏
        self._status_bar = StatusBar(self, bg=COLORS["bg_sidebar"])
        self._status_bar.pack(side="bottom", fill="x")

        # 延迟初始化各面板（先显示主窗口）
        self.after(100, self._init_panels)

    def _build_sidebar(self, parent):
        sidebar = tk.Frame(parent, bg=COLORS["bg_sidebar"], width=210)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # 顶部 Logo
        logo = tk.Frame(sidebar, bg=COLORS["bg_sidebar"], height=64)
        logo.pack(fill="x")
        logo.pack_propagate(False)
        tk.Label(logo, text="◈  HHPDI",
                 bg=COLORS["bg_sidebar"], fg=COLORS["accent"],
                 font=(FONTS["family"], 13, "bold")).pack(
            expand=True, pady=PADDING["lg"])

        Divider(sidebar, color=COLORS["border_light"]).pack(fill="x")

        # 导航按钮区
        nav_area = tk.Frame(sidebar, bg=COLORS["bg_sidebar"])
        nav_area.pack(fill="x", pady=PADDING["sm"])

        for item in NAV_ITEMS:
            btn = self._make_nav_btn(nav_area, item)
            btn.pack(fill="x", padx=PADDING["sm"], pady=1)
            self._nav_btns[item["id"]] = btn

        Divider(sidebar, color=COLORS["border_light"]).pack(fill="x",
                                                              pady=PADDING["sm"])

        # 设置按钮
        settings_btn = self._make_action_btn(sidebar, "⚙  模型设置",
                                              self._open_settings)
        settings_btn.pack(fill="x", padx=PADDING["sm"])

        help_btn = self._make_action_btn(sidebar, "?  使用说明",
                                          self._show_help)
        help_btn.pack(fill="x", padx=PADDING["sm"], pady=(2, 0))

    def _make_nav_btn(self, parent, item):
        """创建侧边栏导航按钮"""
        btn_frame = tk.Frame(parent, bg=COLORS["bg_sidebar"],
                              cursor="hand2")
        btn_frame.pack_propagate(False)

        color   = item["color"]
        nav_id  = item["id"]
        icon    = item["icon"]
        label   = item["label"]
        sub     = item.get("subtitle", "")

        # 左侧激活指示条（初始隐藏）
        indicator = tk.Frame(btn_frame, bg=color, width=3)
        indicator.pack(side="left", fill="y")
        indicator.pack_forget()   # 初始隐藏

        inner = tk.Frame(btn_frame, bg=COLORS["bg_sidebar"])
        inner.pack(side="left", fill="both", expand=True,
                   padx=(PADDING["sm"], 0), pady=PADDING["sm"])

        row = tk.Frame(inner, bg=COLORS["bg_sidebar"])
        row.pack(fill="x")

        icon_lbl = tk.Label(row, text=icon, bg=COLORS["bg_sidebar"],
                             fg=color, font=FONTS["md"], width=2)
        icon_lbl.pack(side="left")

        title_lbl = tk.Label(row, text=label, bg=COLORS["bg_sidebar"],
                              fg=COLORS["text_secondary"],
                              font=FONTS["md"], anchor="w")
        title_lbl.pack(side="left", padx=(4, 0))

        sub_lbl = None
        if sub:
            sub_lbl = tk.Label(inner, text=sub,
                                bg=COLORS["bg_sidebar"],
                                fg=COLORS["text_muted"],
                                font=FONTS["xs"], anchor="w",
                                padx=(PADDING["md"] + 6))
            sub_lbl.pack(fill="x")

        # 存储引用
        btn_frame._indicator = indicator
        btn_frame._inner     = inner
        btn_frame._icon      = icon_lbl
        btn_frame._title     = title_lbl
        btn_frame._sub       = sub_lbl
        btn_frame._color     = color
        btn_frame._nav_id    = nav_id

        def _hover_in(e):
            if self._active_nav != nav_id:
                self._set_nav_hover(btn_frame, True)

        def _hover_out(e):
            if self._active_nav != nav_id:
                self._set_nav_hover(btn_frame, False)

        def _click(e):
            self._navigate(nav_id)

        for w in [btn_frame, inner, row, icon_lbl, title_lbl]:
            w.bind("<Enter>", _hover_in)
            w.bind("<Leave>", _hover_out)
            w.bind("<Button-1>", _click)
        if sub_lbl:
            sub_lbl.bind("<Enter>", _hover_in)
            sub_lbl.bind("<Leave>", _hover_out)
            sub_lbl.bind("<Button-1>", _click)

        return btn_frame

    def _set_nav_hover(self, btn_frame, active):
        bg = COLORS["bg_hover"] if active else COLORS["bg_sidebar"]
        btn_frame.config(bg=bg)
        btn_frame._inner.config(bg=bg)
        for w in btn_frame._inner.winfo_children():
            w.config(bg=bg)
            for ww in w.winfo_children():
                ww.config(bg=bg)

    def _set_nav_active(self, btn_frame, active):
        color = btn_frame._color
        if active:
            btn_frame._indicator.pack(side="left", fill="y")
            btn_frame.config(bg=COLORS["bg_active"])
            btn_frame._inner.config(bg=COLORS["bg_active"])
            btn_frame._icon.config(fg=color,
                                    bg=COLORS["bg_active"])
            btn_frame._title.config(fg=COLORS["text_primary"],
                                     bg=COLORS["bg_active"],
                                     font=FONTS["h3"])
            if btn_frame._sub:
                btn_frame._sub.config(bg=COLORS["bg_active"])
            for w in btn_frame._inner.winfo_children():
                w.config(bg=COLORS["bg_active"])
                for ww in w.winfo_children():
                    try:
                        ww.config(bg=COLORS["bg_active"])
                    except Exception:
                        pass
        else:
            btn_frame._indicator.pack_forget()
            btn_frame.config(bg=COLORS["bg_sidebar"])
            btn_frame._inner.config(bg=COLORS["bg_sidebar"])
            btn_frame._icon.config(fg=color,
                                    bg=COLORS["bg_sidebar"])
            btn_frame._title.config(fg=COLORS["text_secondary"],
                                     bg=COLORS["bg_sidebar"],
                                     font=FONTS["md"])
            if btn_frame._sub:
                btn_frame._sub.config(bg=COLORS["bg_sidebar"])
            for w in btn_frame._inner.winfo_children():
                try:
                    w.config(bg=COLORS["bg_sidebar"])
                except Exception:
                    pass
                for ww in w.winfo_children():
                    try:
                        ww.config(bg=COLORS["bg_sidebar"])
                    except Exception:
                        pass

    def _make_action_btn(self, parent, text, cmd):
        btn = tk.Button(parent, text=text,
                        bg=COLORS["bg_sidebar"],
                        fg=COLORS["text_muted"],
                        activebackground=COLORS["bg_hover"],
                        activeforeground=COLORS["text_secondary"],
                        relief="flat", bd=0, cursor="hand2",
                        font=FONTS["sm"],
                        padx=PADDING["md"], pady=PADDING["sm"],
                        anchor="w", command=cmd)
        btn.bind("<Enter>",
                 lambda e: btn.config(fg=COLORS["text_secondary"]))
        btn.bind("<Leave>",
                 lambda e: btn.config(fg=COLORS["text_muted"]))
        return btn

    def _build_content(self, parent):
        self._content = tk.Frame(parent, bg=COLORS["bg_main"])
        self._content.pack(side="left", fill="both", expand=True)

    # ── 面板初始化（延迟）───────────────────────────────────

    def _init_panels(self):
        """延迟初始化所有工具面板"""
        from gui.home_panel import HomePanel
        self._panels["home"] = HomePanel(
            self._content,
            navigate_cb=self._navigate,
            bg=COLORS["bg_main"],
        )

        from tools.tool1_parser import Tool1Panel
        self._panels["tool1"] = Tool1Panel(
            self._content,
            shared_state=self._shared,
            status_bar=self._status_bar,
            navigate_cb=self._navigate,
            bg=COLORS["bg_main"],
        )

        from tools.tool2_converter import Tool2Panel
        self._panels["tool2"] = Tool2Panel(
            self._content,
            shared_state=self._shared,
            status_bar=self._status_bar,
            navigate_cb=self._navigate,
            bg=COLORS["bg_main"],
        )

        from tools.tool3_annotator import Tool3Panel
        self._panels["tool3"] = Tool3Panel(
            self._content,
            shared_state=self._shared,
            status_bar=self._status_bar,
            navigate_cb=self._navigate,
            bg=COLORS["bg_main"],
        )

        from tools.pipeline_panel import PipelinePanel
        self._panels["pipeline"] = PipelinePanel(
            self._content,
            shared_state=self._shared,
            status_bar=self._status_bar,
            switch_to_tool_cb=self._navigate,
            bg=COLORS["bg_main"],
        )

        # 重新显示当前页（home）
        self._navigate("home")

    # ── 导航 ─────────────────────────────────────────────────

    def _navigate(self, nav_id: str):
        if nav_id not in self._panels:
            return

        # 更新侧边栏
        if self._active_nav and self._active_nav in self._nav_btns:
            self._set_nav_active(self._nav_btns[self._active_nav], False)
        if nav_id in self._nav_btns:
            self._set_nav_active(self._nav_btns[nav_id], True)

        # 切换面板
        for pid, panel in self._panels.items():
            panel.pack_forget()

        self._panels[nav_id].pack(fill="both", expand=True)
        self._active_nav = nav_id

        # 更新状态栏
        nav_labels = {item["id"]: item["label"] for item in NAV_ITEMS}
        lbl = nav_labels.get(nav_id, nav_id)
        self._status_bar.set(f"当前：{lbl}", "normal")

    # ── 菜单动作 ─────────────────────────────────────────────

    def _open_settings(self):
        from gui.settings_window import SettingsWindow
        SettingsWindow(self)

    def _show_help(self):
        help_text = (
            "HHPDI文档数据处理智能助手 — 使用说明\n\n"
            "① 文档解析\n"
            "   选择 PDF 或 Word 文档，点击「开始解析」\n"
            "   需要先在「模型设置」中配置 VLM API Key\n\n"
            "② MD → Word\n"
            "   选择 Markdown 文件，点击「开始转换」\n"
            "   可直接使用步骤①的输出结果\n\n"
            "③ 数据标注\n"
            "   选择 MD 文件，配置 LLM API Key，点击「开始处理」\n"
            "   输出带 tags 的标注文件\n\n"
            "⚡ 流水线\n"
            "   选择文档，勾选步骤，一键自动执行全流程\n\n"
            "💡 提示\n"
            "   文档解析完成后可点击「→ 传递给 MD→Word / 数据标注」\n"
            "   在其他工具中直接使用，无需重新选择文件"
        )
        messagebox.showinfo("使用说明", help_text, parent=self)
