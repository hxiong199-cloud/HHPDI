"""
区域裁切器
根据 VLM 分析结果，从页面图片中裁切图片/表格/公式区域
"""

from pathlib import Path
from PIL import Image


def crop_region(page_image: Image.Image, bbox: list,
                out_path: Path, padding: int = 4) -> bool:
    """
    从页面图片中裁切指定区域并保存
    bbox: [x1, y1, x2, y2] 像素坐标（相对于渲染图）
    padding: 裁切边距
    """
    w, h = page_image.size
    x1 = max(0, int(bbox[0]) - padding)
    y1 = max(0, int(bbox[1]) - padding)
    x2 = min(w, int(bbox[2]) + padding)
    y2 = min(h, int(bbox[3]) + padding)

    if x2 <= x1 or y2 <= y1:
        return False

    cropped = page_image.crop((x1, y1, x2, y2))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    cropped.save(str(out_path), "PNG")
    return True


def scale_bbox_to_page(bbox: list, page_w: float, page_h: float,
                       img_w: int, img_h: int) -> list:
    """
    将 VLM 返回的像素坐标（基于渲染图）转换为 PDF 页面坐标
    """
    sx = page_w / img_w
    sy = page_h / img_h
    return [bbox[0] * sx, bbox[1] * sy, bbox[2] * sx, bbox[3] * sy]


def save_page_image(page_image: Image.Image, out_path: Path) -> str:
    """保存整页渲染图（用于调试/备用）"""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    page_image.save(str(out_path), "PNG")
    return str(out_path)
