"""
PDF 加载器
- 文字型 PDF: PyMuPDF 提取文本，基于全页字号分布做相对标题判断
- 扫描型 PDF: 渲染为图片，走 VLM 分析
- 每页同时渲染为 PNG 供 VLM 做语义层级校正
"""

import fitz  # PyMuPDF
from pathlib import Path
from PIL import Image
import io
import statistics
from config.settings import get_config


def render_page_to_image(page: fitz.Page, dpi: int = 150) -> Image.Image:
    """将 PDF 页面渲染为 PIL Image"""
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img_data = pix.tobytes("png")
    return Image.open(io.BytesIO(img_data))


def is_scanned_page(page: fitz.Page, min_text_len: int = 50) -> bool:
    """判断页面是否为扫描图（文字内容极少）"""
    text = page.get_text("text").strip()
    return len(text) < min_text_len


def has_visual_elements(page: fitz.Page, drawing_threshold: int = 6) -> bool:
    """
    判断文字型页面是否含有需要 VLM 识别的视觉元素（图片/表格/公式）。
    True  → 需要调用 VLM 检测特殊区域
    False → 纯文字页，可跳过 VLM

    检测逻辑：
    1. 页面含嵌入式图片 → 有视觉元素
    2. 矢量绘图路径数量超过阈值 → 可能存在表格线条或图形
       （一张简单水平线 ≈ 1~2条，3列×5行表格 ≈ 10条）
    """
    if page.get_images(full=True):
        return True
    if len(page.get_drawings()) >= drawing_threshold:
        return True
    return False


def _collect_all_font_sizes(page: fitz.Page) -> list[float]:
    """收集页面内所有 span 的字号，用于统计基准"""
    sizes = []
    raw = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
    for block in raw.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                t = span.get("text", "").strip()
                size = span.get("size", 0)
                if t and size > 0:
                    sizes.append(size)
    return sizes


def _infer_heading_level(font_size: float, is_bold: bool,
                          body_size: float, size_tiers: list[float]) -> tuple[str, int]:
    """
    根据字号与正文基准的相对关系推断块类型和标题级别。
    size_tiers: 从大到小排列的标题字号档位列表
    返回 (block_type, level)，level=0 表示正文
    """
    # 字号明显大于正文，或字号相近但加粗 → 视为标题
    ratio = font_size / body_size if body_size > 0 else 1.0

    if ratio < 1.05 and not is_bold:
        return "text", 0

    # 根据在 size_tiers 中的位置确定 level
    for i, tier_size in enumerate(size_tiers):
        if font_size >= tier_size - 0.5:
            return "heading", i + 1

    # 加粗但字号接近正文 → h3 级别
    if is_bold:
        return "heading", 3

    return "text", 0


def extract_page_text_blocks(page: fitz.Page) -> list[dict]:
    """
    从文字型 PDF 页面提取文字块。
    使用全页字号分布做相对判断，而非固定阈值。
    返回 list of {type, level, text, bbox, font_size, is_bold, page_width, page_height}
    """
    page_h = page.rect.height
    page_w = page.rect.width

    # ── Step 1: 收集全页字号分布，计算正文基准字号 ──────────
    all_sizes = _collect_all_font_sizes(page)
    if not all_sizes:
        return []

    # 正文基准：出现频率最高的字号（众数近似）
    # 用中位数作为正文基准，更稳健
    body_size = statistics.median(all_sizes)

    # 标题档位：比正文大的字号，去重排序（从大到小）
    larger_sizes = sorted(set(
        s for s in all_sizes if s > body_size * 1.05
    ), reverse=True)

    # 最多保留 4 个档位（h1~h4），避免噪音
    size_tiers = larger_sizes[:4]

    # ── Step 2: 逐块提取 ─────────────────────────────────────
    raw = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
    blocks_raw = []

    for block in raw.get("blocks", []):
        if block.get("type") != 0:
            continue

        lines_text = []
        max_size = 0
        is_bold = False

        for line in block.get("lines", []):
            for span in line.get("spans", []):
                t = span.get("text", "").strip()
                if not t:
                    continue
                lines_text.append(t)
                size = span.get("size", 12)
                max_size = max(max_size, size)
                # flags bit 4 = bold in PyMuPDF
                if span.get("flags", 0) & 16:
                    is_bold = True

        full_text = " ".join(lines_text).strip()
        if not full_text:
            continue

        # 过滤页眉页脚：位于页面顶部 8% 或底部 8% 且内容短小
        bbox = block["bbox"]
        rel_y_top = bbox[1] / page_h
        rel_y_bot = bbox[3] / page_h
        text_len = len(full_text)
        if (rel_y_top < 0.08 or rel_y_bot > 0.92) and text_len < 60:
            continue

        block_type, level = _infer_heading_level(
            max_size, is_bold, body_size, size_tiers
        )

        blocks_raw.append({
            "type": block_type,
            "level": level,
            "text": full_text,
            "bbox": [bbox[0], bbox[1], bbox[2], bbox[3]],
            "font_size": max_size,
            "is_bold": is_bold,
            "page_width": page_w,
            "page_height": page_h,
        })

    # ── Step 3: 标题级别归一化（确保层级连续，避免跳级） ─────
    blocks_raw = _normalize_heading_levels(blocks_raw)

    return blocks_raw


def _normalize_heading_levels(blocks: list[dict]) -> list[dict]:
    """
    对整页标题级别做归一化：
    1. 收集所有出现的标题级别，重新映射为连续的 1,2,3...
    2. 避免 h1 下面直接跳到 h4 这类情况
    """
    heading_sizes = sorted(set(
        b["font_size"] for b in blocks if b["type"] == "heading"
    ), reverse=True)

    # 建立字号 → 连续 level 的映射
    size_to_level = {size: i + 1 for i, size in enumerate(heading_sizes)}

    for b in blocks:
        if b["type"] == "heading":
            b["level"] = size_to_level.get(b["font_size"], 1)

    return blocks


def extract_page_images(page: fitz.Page, out_dir: Path,
                        page_no: int, doc: fitz.Document) -> list[dict]:
    """
    提取页面内嵌图片，保存到 out_dir/images/
    返回 list of {path, bbox, page_no}
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    results = []
    img_list = page.get_images(full=True)
    for idx, img_info in enumerate(img_list):
        xref = img_info[0]
        try:
            base_img = doc.extract_image(xref)
            img_bytes = base_img["image"]
            ext = base_img.get("ext", "png")
            fname = f"fig_p{page_no + 1}_{idx + 1}.{ext}"
            fpath = out_dir / fname
            with open(fpath, "wb") as f:
                f.write(img_bytes)

            rects = page.get_image_rects(xref)
            bbox = list(rects[0]) if rects else [0, 0, 100, 100]

            results.append({
                "type": "figure",
                "path": str(fpath),
                "filename": fname,
                "bbox": bbox,
                "page_no": page_no,
            })
        except Exception:
            continue
    return results


def load_pdf(file_path: str, progress_cb=None) -> dict:
    """
    加载 PDF，返回结构化页面数据
    progress_cb(current, total, message) 用于 GUI 进度更新
    """
    cfg = get_config()
    dpi = cfg["parse_options"]["render_dpi"]
    doc = fitz.open(file_path)
    total = len(doc)
    pages_data = []

    for page_no in range(total):
        if progress_cb:
            progress_cb(page_no, total, f"加载 PDF 第 {page_no + 1}/{total} 页...")

        page = doc[page_no]
        scanned = is_scanned_page(page)
        img = render_page_to_image(page, dpi)

        pages_data.append({
            "page_no": page_no,
            "is_scanned": scanned,
            "has_graphics": scanned or has_visual_elements(page),
            "pil_image": img,
            "text_blocks": [] if scanned else extract_page_text_blocks(page),
            "embedded_images": [],
            "width": page.rect.width,
            "height": page.rect.height,
        })

    doc.close()
    return {
        "source": file_path,
        "doc_type": "pdf",
        "total_pages": total,
        "pages": pages_data,
    }
