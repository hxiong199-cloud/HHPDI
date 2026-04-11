"""
DocFlow Pro — 统一视觉主题 v2
高对比深色 + 明亮强调色，大字号，去除冗余注释文字
"""
import sys

# ── 颜色系统（高对比版）────────────────────────────────────

COLORS = {
    # 背景层次：拉大色差
    "bg_deep":    "#080B0F",   # 最深底色（几乎纯黑）
    "bg_main":    "#0F1318",   # 主背景
    "bg_card":    "#181E28",   # 卡片（与主背景对比明显）
    "bg_sidebar": "#0A0D12",   # 侧边栏（最深）
    "bg_input":   "#1C2333",   # 输入框（略蓝调）
    "bg_hover":   "#232C3D",   # 悬停
    "bg_active":  "#152542",   # 激活（蓝调）
    "bg_stripe":  "#141920",

    # 主色调：高饱和金色
    "accent":        "#F0C060",   # 亮金（更饱和、更亮）
    "accent_bright": "#FFD980",   # 高亮金
    "accent_dim":    "#9A7830",
    "accent_bg":     "#1A1500",

    # 功能色（全部提高亮度/饱和度）
    "success": "#4ADE80",   # 更亮的绿
    "warning": "#FBBF24",   # 更亮的黄
    "error":   "#FF5555",
    "info":    "#60B4FF",
    "running": "#93C5FD",

    # 文字层次（对比度大幅提升）
    "text_primary":   "#F0F4FF",   # 几乎纯白
    "text_secondary": "#A8B8D0",   # 浅蓝灰（原来太暗）
    "text_muted":     "#5A6880",   # 辅助文字
    "text_accent":    "#F0C060",

    # 边框（更明显）
    "border":       "#2A3550",
    "border_light": "#1E2840",
    "border_focus": "#F0C060",

    # 工具专属色（高饱和）
    "tool1_color":    "#4FACFF",   # 文档解析：亮蓝
    "tool2_color":    "#4ADE80",   # MD→Word：亮绿
    "tool3_color":    "#FF8C42",   # 数据标注：亮橙
    "pipeline_color": "#F0C060",   # 流水线：金
}

# ── 字体系统（字号全面放大）────────────────────────────────
if sys.platform == "win32":
    _FONT_FAMILY = "Microsoft YaHei UI"
elif sys.platform == "darwin":
    _FONT_FAMILY = "PingFang SC"
else:
    _FONT_FAMILY = "Noto Sans CJK SC"

_MONO = "Consolas" if sys.platform == "win32" else "Menlo"

FONTS = {
    "family": _FONT_FAMILY,
    "mono":   _MONO,
    # 正文字号从9/10/11 → 11/12/13
    "xs":    (_FONT_FAMILY, 11),
    "sm":    (_FONT_FAMILY, 12),
    "md":    (_FONT_FAMILY, 13),
    "lg":    (_FONT_FAMILY, 15),
    "xl":    (_FONT_FAMILY, 17),
    "title": (_FONT_FAMILY, 20, "bold"),
    "h1":    (_FONT_FAMILY, 18, "bold"),
    "h2":    (_FONT_FAMILY, 15, "bold"),
    "h3":    (_FONT_FAMILY, 13, "bold"),
    "mono_sm": (_MONO, 11),
    "mono_md": (_MONO, 12),
}

PADDING = {
    "xs": 4, "sm": 8, "md": 14, "lg": 18, "xl": 26, "xxl": 36,
}

# ── 侧边栏导航配置 ───────────────────────────────────────────
NAV_ITEMS = [
    {
        "id": "home",
        "icon": "◈",
        "label": "主页",
        "color": COLORS["accent"],
    },
    {
        "id": "tool1",
        "icon": "◆",
        "label": "文档解析",
        "color": COLORS["tool1_color"],
    },
    {
        "id": "tool2",
        "icon": "◆",
        "label": "MD → Word",
        "color": COLORS["tool2_color"],
    },
    {
        "id": "tool3",
        "icon": "◆",
        "label": "数据标注",
        "color": COLORS["tool3_color"],
    },
    {
        "id": "pipeline",
        "icon": "◈",
        "label": "流水线",
        "color": COLORS["pipeline_color"],
    },
]
