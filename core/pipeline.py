"""
解析总调度器 Pipeline
协调 pdf_loader / word_loader / vlm_client / region_extractor / md_builder
"""

import os
import threading
from pathlib import Path
from datetime import datetime

from config.settings import get_config
from core.pdf_loader import load_pdf
from core.word_loader import load_word
from core.region_extractor import crop_region, scale_bbox_to_page
from core.vlm_client import analyze_page_layout, table_image_to_markdown, formula_image_to_latex
from core.md_builder import blocks_to_markdown, write_markdown, create_zip
from core.word_exporter import export_word
from core.md_cleaner import clean_markdown


class ParseResult:
    def __init__(self):
        self.success = False
        self.md_path = ""
        self.docx_path = ""
        self.zip_path = ""
        self.images_dir = ""
        self.md_content = ""
        self.error = ""
        self.stats = {}


def _make_output_dir(source_path: str) -> Path:
    # 输出目录与输入文件在同一目录，子文件夹以文件名+时间戳命名
    stem = Path(source_path).stem
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = Path(source_path).parent
    out = base / f"{stem}_{ts}"
    out.mkdir(parents=True, exist_ok=True)
    (out / "images").mkdir(exist_ok=True)
    return out


# ── PDF Pipeline ─────────────────────────────────────────────

def _process_pdf(file_path: str, out_dir: Path,
                 progress_cb, cancel_event) -> list[dict]:
    """处理 PDF，返回所有内容块"""
    cfg = get_config()
    extract_formulas = cfg["parse_options"]["extract_formulas"]
    extract_tables_md = cfg["parse_options"]["extract_tables_as_md"]
    img_dir = out_dir / "images"
    img_dir.mkdir(exist_ok=True)

    pdf_data = load_pdf(file_path, progress_cb=progress_cb)
    total_pages = pdf_data["total_pages"]
    all_blocks = []

    fig_counter = {"figure": 0, "table": 0, "formula": 0}

    for page_info in pdf_data["pages"]:
        if cancel_event and cancel_event.is_set():
            break

        page_no = page_info["page_no"]
        pil_img = page_info["pil_image"]
        img_w, img_h = pil_img.size
        page_w = page_info["width"]
        page_h = page_info["height"]

        if progress_cb:
            progress_cb(page_no, total_pages,
                        f"VLM 分析第 {page_no + 1}/{total_pages} 页版面...")

        # 保存页面图片临时文件
        tmp_page_path = out_dir / f"_tmp_page_{page_no}.png"
        pil_img.save(str(tmp_page_path), "PNG")

        if page_info["is_scanned"]:
            # 扫描页：全页交给 VLM 分析
            try:
                layout = analyze_page_layout(str(tmp_page_path))
                raw_blocks = layout.get("blocks", [])
            except Exception as e:
                all_blocks.append({"type": "text", "level": 0,
                                   "text": f"[第{page_no+1}页解析失败: {e}]",
                                   "page_no": page_no})
                tmp_page_path.unlink(missing_ok=True)
                continue
        else:
            # 文字页：用 PyMuPDF 提取文字/标题块
            # 仅当页面含图形元素时才调用 VLM 检测表格/公式（纯文字页跳过）
            text_blocks = page_info["text_blocks"]

            # VLM 全页分析，只取 table/formula/figure 区域
            vlm_special_blocks = []
            if page_info.get("has_graphics", True):
                try:
                    if progress_cb:
                        progress_cb(page_no, total_pages,
                                    f"VLM 检测第 {page_no+1}/{total_pages} 页表格/公式...")
                    layout = analyze_page_layout(str(tmp_page_path))
                    for b in layout.get("blocks", []):
                        if b.get("type") in ("table", "formula", "figure"):
                            vlm_special_blocks.append(b)
                except Exception:
                    pass  # VLM 失败时只用文字块
            else:
                if progress_cb:
                    progress_cb(page_no, total_pages,
                                f"第 {page_no+1}/{total_pages} 页（纯文字，跳过VLM）")

            # 将 VLM 检测到的特殊区域 bbox 转换为 PDF 坐标系
            # VLM 返回的是像素坐标（基于渲染图），需要转回 PDF 点坐标
            def _vlm_to_pdf_bbox(vbbox):
                """将 VLM 像素坐标 bbox 转换回 PDF 坐标"""
                if not vbbox or len(vbbox) < 4:
                    return vbbox
                x0 = vbbox[0] / img_w * page_w
                y0 = vbbox[1] / img_h * page_h
                x1 = vbbox[2] / img_w * page_w
                y1 = vbbox[3] / img_h * page_h
                return [x0, y0, x1, y1]

            # 过滤掉与 VLM 特殊区域重叠的文字块（避免把表格内文字当正文）
            def _overlaps(tb_bbox, special_bboxes_pdf, threshold=0.5):
                """判断文字块是否与任一特殊区域重叠超过阈值"""
                if not tb_bbox or len(tb_bbox) < 4:
                    return False
                tx0, ty0, tx1, ty1 = tb_bbox
                ta = max(0, tx1 - tx0) * max(0, ty1 - ty0)
                if ta == 0:
                    return False
                for sx0, sy0, sx1, sy1 in special_bboxes_pdf:
                    ix0 = max(tx0, sx0)
                    iy0 = max(ty0, sy0)
                    ix1 = min(tx1, sx1)
                    iy1 = min(ty1, sy1)
                    inter = max(0, ix1 - ix0) * max(0, iy1 - iy0)
                    if inter / ta >= threshold:
                        return True
                return False

            special_pdf_bboxes = [_vlm_to_pdf_bbox(b.get("bbox", [])) for b in vlm_special_blocks]
            filtered_text_blocks = [
                b for b in text_blocks
                if not _overlaps(b.get("bbox", []), special_pdf_bboxes)
            ]

            # 合并：过滤后的文字块 + VLM 检测的特殊块（bbox 仍用像素坐标，后续统一转换）
            raw_blocks = filtered_text_blocks + vlm_special_blocks

            # 附加嵌入图片信息（从 PDF 中提取）
            try:
                import fitz
                doc = fitz.open(file_path)
                page = doc[page_no]
                from core.pdf_loader import extract_page_images
                emb_imgs = extract_page_images(page, img_dir, page_no, doc)
                doc.close()
                for ei in emb_imgs:
                    raw_blocks.append({
                        "type": "figure",
                        "level": 0,
                        "text": "",
                        "bbox": ei["bbox"],
                        "path": ei["path"],
                        "filename": ei["filename"],
                        "page_no": page_no,
                        "needs_vlm": False,
                    })
            except Exception:
                pass

        # 处理各区域块
        for block in raw_blocks:
            btype = block.get("type", "text")
            bbox = block.get("bbox", [])
            page_bbox = (scale_bbox_to_page(bbox, page_w, page_h, img_w, img_h)
                         if bbox else [])

            if btype in ("figure", "table", "formula"):
                fig_counter[btype] += 1
                fname = f"{btype}_p{page_no+1}_{fig_counter[btype]}.png"
                fpath = img_dir / fname

                if not block.get("path"):  # 需要从页面图裁切
                    success = crop_region(pil_img, bbox, fpath)
                    if not success:
                        continue
                else:
                    # 已有路径（嵌入图），转换为 PNG 后复制
                    # 注意：PyMuPDF 可能提取出 .jpeg/.jpg/.jpx 等格式
                    # 必须用 PIL 重新编码为 PNG，不能直接 shutil.copy2
                    # （原样复制会导致 JPEG 内容 + .png 扩展名，python-docx 报错）
                    src = Path(block["path"])
                    if src.exists() and src != fpath:
                        try:
                            from PIL import Image as _PILImg
                            with _PILImg.open(str(src)) as _im:
                                _im.convert("RGBA" if _im.mode in ("RGBA", "LA", "PA") else "RGB").save(str(fpath), "PNG")
                        except Exception:
                            import shutil
                            shutil.copy2(src, fpath)  # 转换失败时兜底

                out_block = {
                    "type": btype,
                    "level": 0,
                    "text": "",
                    "filename": fname,
                    "bbox": page_bbox,
                    "page_no": page_no,
                    "md_table": "",
                    "latex": "",
                }

                # 表格额外请求 VLM 转 MD
                if btype == "table" and extract_tables_md:
                    try:
                        if progress_cb:
                            progress_cb(page_no, total_pages,
                                        f"识别第{page_no+1}页表格...")
                        out_block["md_table"] = table_image_to_markdown(str(fpath))
                    except Exception:
                        pass

                # 公式请求 VLM 转 LaTeX
                if btype == "formula" and extract_formulas:
                    try:
                        if progress_cb:
                            progress_cb(page_no, total_pages,
                                        f"识别第{page_no+1}页公式...")
                        out_block["latex"] = formula_image_to_latex(str(fpath))
                    except Exception:
                        pass

                all_blocks.append(out_block)

            else:
                all_blocks.append({
                    "type": btype,
                    "level": block.get("level", 0),
                    "text": block.get("text", ""),
                    "filename": None,
                    "bbox": page_bbox,
                    "page_no": page_no,
                    "md_table": "",
                    "latex": "",
                })

        tmp_page_path.unlink(missing_ok=True)

    return all_blocks


# ── Word Pipeline ────────────────────────────────────────────

def _process_word(file_path: str, out_dir: Path,
                  progress_cb, cancel_event) -> list[dict]:
    cfg = get_config()
    extract_tables_md = cfg["parse_options"]["extract_tables_as_md"]
    extract_formulas = cfg["parse_options"]["extract_formulas"]

    word_data = load_word(file_path, out_dir, progress_cb=progress_cb)
    raw_blocks = word_data["blocks"]
    all_blocks = []

    total = len(raw_blocks)
    for idx, block in enumerate(raw_blocks):
        if cancel_event and cancel_event.is_set():
            break

        btype = block.get("type", "text")

        if btype == "table" and extract_tables_md and not block.get("md_table"):
            # 理论上 word_loader 已经处理了，这里是兜底
            pass

        all_blocks.append(block)

        if progress_cb:
            progress_cb(idx, total, f"处理内容块 {idx+1}/{total}...")

    return all_blocks


# ── 主入口 ───────────────────────────────────────────────────

def run_pipeline(file_path: str,
                 progress_cb=None,
                 done_cb=None,
                 cancel_event=None) -> ParseResult:
    """
    同步执行完整解析流程
    progress_cb(current, total, message)
    done_cb(result: ParseResult)
    """
    result = ParseResult()
    try:
        out_dir = _make_output_dir(file_path)
        suffix = Path(file_path).suffix.lower()

        if suffix == ".pdf":
            blocks = _process_pdf(file_path, out_dir, progress_cb, cancel_event)
        elif suffix in (".docx", ".doc"):
            blocks = _process_word(file_path, out_dir, progress_cb, cancel_event)
        else:
            raise ValueError(f"不支持的文件格式：{suffix}")

        if progress_cb:
            progress_cb(1, 1, "组装 Markdown...")

        md_content = blocks_to_markdown(blocks, images_rel_dir="images",
                                        tables_rel_dir="tables", out_dir=out_dir)
        md_content = clean_markdown(md_content)
        stem = Path(file_path).stem
        md_path = out_dir / f"{stem}.md"
        write_markdown(md_content, md_path)

        # 同步导出 Word
        docx_path = out_dir / f"{stem}.docx"
        try:
            if progress_cb:
                progress_cb(1, 1, "导出 Word 文档...")
            export_word(blocks, out_dir / "images", docx_path)
        except Exception as e:
            import traceback
            docx_path = None
            # Word 导出失败不影响 MD 结果，记录错误但继续
            result.error = f"Word 导出失败（MD 仍可用）: {e}"

        zip_path = out_dir / f"{stem}.zip"
        create_zip(md_path, out_dir / "images", zip_path,
                   docx_path=docx_path)

        result.success = True
        result.md_path = str(md_path)
        result.docx_path = str(docx_path) if docx_path else ""
        result.zip_path = str(zip_path)
        result.images_dir = str(out_dir / "images")
        result.md_content = md_content
        result.stats = {
            "blocks": len(blocks),
            "figures": sum(1 for b in blocks if b.get("type") == "figure"),
            "tables": sum(1 for b in blocks if b.get("type") == "table"),
            "formulas": sum(1 for b in blocks if b.get("type") == "formula"),
        }

    except Exception as e:
        import traceback
        result.error = f"{e}\n{traceback.format_exc()}"

    if done_cb:
        done_cb(result)
    return result


def run_pipeline_async(file_path: str,
                       progress_cb=None,
                       done_cb=None) -> threading.Event:
    """异步执行，返回 cancel_event（调用 .set() 可取消）"""
    cancel_event = threading.Event()

    def _run():
        run_pipeline(file_path, progress_cb, done_cb, cancel_event)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return cancel_event


def run_batch_async(file_paths: list,
                    file_progress_cb=None,
                    file_done_cb=None,
                    all_done_cb=None,
                    max_workers: int = 3) -> threading.Event:
    """
    并行处理多个文件，返回 cancel_event（调用 .set() 可取消所有文件）

    file_progress_cb(idx, cur, total, msg)  — 某文件进度更新
    file_done_cb(idx, result)               — 某文件完成
    all_done_cb(results)                    — 全部完成
    """
    from concurrent.futures import ThreadPoolExecutor

    cancel_event = threading.Event()
    results = [None] * len(file_paths)

    def _run_one(idx):
        def _progress(cur, total, msg):
            if file_progress_cb:
                file_progress_cb(idx, cur, total, msg)

        result = run_pipeline(file_paths[idx], _progress, None, cancel_event)
        results[idx] = result
        if file_done_cb:
            file_done_cb(idx, result)

    def _run_all():
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(_run_one, i)
                       for i in range(len(file_paths))]
            for future in futures:
                try:
                    future.result()
                except Exception:
                    pass  # 错误已在 run_pipeline 内捕获并存入 result.error
        if all_done_cb:
            all_done_cb(results)

    t = threading.Thread(target=_run_all, daemon=True)
    t.start()
    return cancel_event
