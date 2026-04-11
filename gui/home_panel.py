"""
主页面板 — HHPDI 文档数据处理智能助手
设计系统来源：ui-ux-pro-max skill
  · 风格：Glassmorphism Dark + Bento Grid
  · 配色：深蓝黑底 + 工具专属彩色 + 绿色 CTA
  · 动效：Canvas 圆角卡片 + hover 光晕 + 颜色过渡
"""
import tkinter as tk
from gui.theme import COLORS, FONTS, PADDING

# ── 设计系统 (ui-ux-pro-max 推荐，针对数据处理工具) ──────────
# 配色方案：专业数据工具 · 深蓝黑 + 彩色强调 + 绿色 CTA
_DS = {
    # 背景层次
    "bg_page":      "#0A0E1A",   # 最深页面背景（深蓝黑）
    "card_normal":  "#111827",   # 卡片默认（深海军蓝）
    "card_hover":   "#162035",   # 卡片悬停（略亮）
    "card_border":  "#1E2D45",   # 卡片默认边框
    "glass_top":    "#1C2A42",   # 玻璃顶部高光条
    "glass_shine":  "#243354",   # 高光条亮边
    # 文字
    "text_title":   "#EDF2FF",   # 主标题
    "text_body":    "#7B90B8",   # 正文
    "text_dim":     "#3D4F6A",   # 辅助说明
    # 状态颜色
    "cta_green":    "#22C55E",   # CTA 主色（设计系统推荐）
    "cta_bg":       "#0B1F12",   # CTA 底色
    "tag_bg":       "#0D1828",   # 小标签背景
}

# 功能卡片定义
_CARDS = [
    {
        "id":     "tool1",
        "icon_c": "M",          # 用文字模拟图标符号
        "symbol": "⬢",
        "title":  "文档解析",
        "lines":  ["PDF / Word", "→ Markdown"],
        "tags":   ["批量处理", "扫描件 OCR"],
        "color":  "#4FACFF",    # 亮蓝
        "glow":   "#0A1E35",    # 低饱和蓝底（外发光用）
    },
    {
        "id":     "tool2",
        "symbol": "⬢",
        "title":  "MD → Word",
        "lines":  ["Markdown", "→ Word 文档"],
        "tags":   ["保留结构", "一键导出"],
        "color":  "#4ADE80",    # 亮绿
        "glow":   "#082018",
    },
    {
        "id":     "tool3",
        "symbol": "⬢",
        "title":  "数据标注",
        "lines":  ["文本打标签", "表格转 QA"],
        "tags":   ["LLM 驱动", "智能标注"],
        "color":  "#FF8C42",    # 亮橙
        "glow":   "#201008",
    },
    {
        "id":     "pipeline",
        "symbol": "◈",
        "title":  "流水线",
        "lines":  ["一键串联", "三个工具"],
        "tags":   ["全自动", "多文件并行"],
        "color":  "#F0C060",    # 金色
        "glow":   "#201A08",
    },
]

_CARD_W = 196
_CARD_H = 228


def _rr(canvas, x1, y1, x2, y2, r, **kw):
    """在 Canvas 上绘制圆角矩形（smooth polygon 实现）"""
    return canvas.create_polygon(
        x1 + r, y1,   x2 - r, y1,
        x2,     y1,   x2,     y1 + r,
        x2,     y2 - r, x2,   y2,
        x2 - r, y2,   x1 + r, y2,
        x1,     y2,   x1,     y2 - r,
        x1,     y1 + r, x1,   y1,
        smooth=True, **kw,
    )


class HomePanel(tk.Frame):
    """主页：Glassmorphism 风格，Bento Grid 布局"""

    def __init__(self, parent, navigate_cb=None, **kw):
        bg = kw.pop("bg", COLORS["bg_main"])
        super().__init__(parent, bg=bg, **kw)
        self._navigate  = navigate_cb
        self._card_refs = {}   # canvas → card_info，供悬停动画使用
        self._build()

    # ── 构建 ──────────────────────────────────────────────────

    def _build(self):
        # 整体居中容器（place 实现真正垂直居中）
        outer = tk.Frame(self, bg=COLORS["bg_main"])
        outer.place(relx=0.5, rely=0.5, anchor="center")

        self._build_hero(outer)
        self._build_cards(outer)
        self._build_footer(outer)

    # ── Hero ──────────────────────────────────────────────────

    def _build_hero(self, parent):
        hero = tk.Frame(parent, bg=COLORS["bg_main"])
        hero.pack(pady=(0, 32))

        # 装饰性顶部光晕条（模拟 glow）
        glow_canvas = tk.Canvas(hero, width=320, height=4,
                                 bg=COLORS["bg_main"], highlightthickness=0)
        glow_canvas.pack()
        glow_canvas.create_rectangle(60, 1, 260, 3,
                                      fill=COLORS["accent"], outline="")
        glow_canvas.create_rectangle(30, 2, 290, 3,
                                      fill="#6A4010", outline="")

        # 主 Logo 图标（Canvas 绘制光晕圈）
        logo_c = tk.Canvas(hero, width=90, height=90,
                            bg=COLORS["bg_main"], highlightthickness=0)
        logo_c.pack(pady=(12, 0))
        # 外光晕圈
        logo_c.create_oval(10, 10, 80, 80, fill="#1A1500", outline="#3A2800", width=1)
        logo_c.create_oval(16, 16, 74, 74, fill="#201A00", outline="#5A3A00", width=1)
        # 图标文字
        logo_c.create_text(45, 45, text="◈",
                            font=(FONTS["family"], 32),
                            fill=COLORS["accent"])

        # 标题
        tk.Label(hero, text="HHPDI 文档数据处理智能助手",
                 bg=COLORS["bg_main"], fg=_DS["text_title"],
                 font=(FONTS["family"], 22, "bold")).pack(pady=(14, 0))

        # 副标题（彩色标签行）
        tag_row = tk.Frame(hero, bg=COLORS["bg_main"])
        tag_row.pack(pady=(8, 0))
        for txt, clr in [("智能文档解析", "#4FACFF"),
                          ("·", COLORS["text_muted"]),
                          ("数据标注",     "#4ADE80"),
                          ("·", COLORS["text_muted"]),
                          ("格式转换",     "#FF8C42"),
                          ("·", COLORS["text_muted"]),
                          ("批量并行",     "#F0C060")]:
            tk.Label(tag_row, text=txt,
                     bg=COLORS["bg_main"], fg=clr,
                     font=FONTS["sm"]).pack(side="left", padx=3)

    # ── 功能卡片区 ────────────────────────────────────────────

    def _build_cards(self, parent):
        row = tk.Frame(parent, bg=COLORS["bg_main"])
        row.pack()
        for i, card_info in enumerate(_CARDS):
            w = self._make_glass_card(row, card_info)
            w.grid(row=0, column=i, padx=10)

    def _make_glass_card(self, parent, info: dict) -> tk.Canvas:
        """
        Glassmorphism 玻璃卡片：
          · Canvas 绘制圆角矩形主体
          · 顶部高光条 → 玻璃折射感
          · 左侧彩色竖条 → 工具标识
          · Hover：外发光 + 彩色边框 + 文字亮化
        """
        W, H, R = _CARD_W, _CARD_H, 14
        color = info["color"]
        glow  = info["glow"]
        bg    = COLORS["bg_main"]

        cv = tk.Canvas(parent, width=W, height=H,
                       bg=bg, highlightthickness=0, cursor="hand2")

        def _redraw(hover: bool):
            cv.delete("all")
            card_fill   = _DS["card_hover"]   if hover else _DS["card_normal"]
            border_c    = color               if hover else _DS["card_border"]
            bw          = 2                   if hover else 1
            txt_title   = _DS["text_title"]   if hover else "#C5D0E8"
            txt_line    = color               if hover else "#5A6E8A"
            txt_tag     = color               if hover else _DS["text_dim"]

            # ── 外发光（hover 时，更大的圆角矩形） ──
            if hover:
                _rr(cv, 3, 3, W-3, H-3, R+3, fill=glow, outline="")
                _rr(cv, 5, 5, W-5, H-5, R+1, fill=glow, outline="")

            # ── 卡片主体 ──
            _rr(cv, 6, 6, W-6, H-6, R, fill=card_fill,
                outline=border_c, width=bw)

            # ── 顶部玻璃高光条 ──
            _rr(cv, 7, 7, W-7, 46, R-2,
                fill=_DS["glass_top"], outline="")
            # 高光最亮上沿
            cv.create_rectangle(7+R, 7, W-7-R, 10,
                                  fill=_DS["glass_shine"], outline="")

            # ── 左侧彩色竖线（工具标识） ──
            cv.create_rectangle(6, 6+R, 9, H-6-R,
                                  fill=color, outline="")

            # ── 图标（顶部高光区内居中） ──
            icon_clr = color if hover else "#506080"
            cv.create_text(W // 2, 28,
                           text=info["symbol"],
                           font=(FONTS["family"], 18),
                           fill=icon_clr)

            # ── 标题 ──
            cv.create_text(W // 2, 68,
                           text=info["title"],
                           font=(FONTS["family"], 14, "bold"),
                           fill=txt_title)

            # ── 功能描述（2行） ──
            for li, line in enumerate(info["lines"]):
                cv.create_text(W // 2, 94 + li * 20,
                               text=line,
                               font=FONTS["xs"],
                               fill=txt_line)

            # ── 分隔线 ──
            sep_y = 148
            cv.create_line(20, sep_y, W-20, sep_y,
                           fill=_DS["card_border"], width=1)

            # ── 标签（居中文字，各占半宽，避免溢出） ──
            tag_y = sep_y + 18
            tag_centers = [W // 4, W * 3 // 4]   # 左半中心 / 右半中心
            for ti, tag in enumerate(info["tags"][:2]):
                cx = tag_centers[ti]
                # 小圆点装饰
                cv.create_oval(cx - 26, tag_y - 3,
                               cx - 20, tag_y + 3,
                               fill=txt_tag, outline="")
                cv.create_text(cx - 6, tag_y,
                               text=tag,
                               font=(FONTS["family"], 9),
                               fill=txt_tag,
                               anchor="w")

            # ── 底部右侧箭头（hover 时出现） ──
            if hover:
                cv.create_text(W - 18, H - 18,
                               text="→",
                               font=FONTS["sm"],
                               fill=color)

        # 初始绘制
        _redraw(False)

        # 悬停 + 点击绑定
        cv.bind("<Enter>",        lambda e: _redraw(True))
        cv.bind("<Leave>",        lambda e: _redraw(False))
        cv.bind("<ButtonPress-1>",
                lambda e: cv.configure(bg=_DS["card_hover"]))
        cv.bind("<ButtonRelease-1>",
                lambda e: (cv.configure(bg=COLORS["bg_main"]),
                           self._navigate and self._navigate(info["id"])))

        return cv

    # ── Footer ────────────────────────────────────────────────

    def _build_footer(self, parent):
        footer = tk.Frame(parent, bg=COLORS["bg_main"])
        footer.pack(pady=(28, 0))

        # 版本信息行
        info_row = tk.Frame(footer, bg=COLORS["bg_main"])
        info_row.pack()
        for txt, clr in [
            ("v 1.0", _DS["text_dim"]),
            ("  ·  ", _DS["text_dim"]),
            ("AI 驱动", _DS["cta_green"]),
            ("  ·  ", _DS["text_dim"]),
            ("批量并行处理", _DS["text_body"]),
            ("  ·  ", _DS["text_dim"]),
            ("VLM + LLM", _DS["text_body"]),
        ]:
            tk.Label(info_row, text=txt,
                     bg=COLORS["bg_main"], fg=clr,
                     font=(FONTS["family"], 10)).pack(side="left")
