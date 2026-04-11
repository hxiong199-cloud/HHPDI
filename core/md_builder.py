"""
Markdown 组装器
将解析后的结构化块列表组装为 Markdown 文本
- 图片 → images/fig_N.png（已有）
- 表格 → tables/table_N.html（另存 HTML，MD里用占位符）
"""

from pathlib import Path
from config.settings import get_config


def blocks_to_markdown(blocks: list[dict], images_rel_dir: str = "images",
                       add_bbox: bool = True,
                       tables_rel_dir: str = "tables",
                       out_dir: Path = None) -> str:
    """
    将内容块列表组装为 Markdown 字符串。
    表格另存为 tables/table_N.html，MD里用占位符：
        <!-- TABLE:tables/table_N.html -->
    """
    cfg = get_config()
    add_bbox = add_bbox and cfg["parse_options"]["add_bbox_comments"]

    lines = []
    prev_type = None
    table_counter = 0

    for block in blocks:
        btype = block.get("type", "text")

        if prev_type is not None:
            lines.append("")

        if btype == "heading":
            level = max(1, min(6, block.get("level", 1)))
            prefix = "#" * level
            lines.append(f"{prefix} {block.get('text', '').strip()}")

        elif btype == "text":
            text = block.get("text", "").strip()
            if text:
                lines.append(text)

        elif btype == "figure":
            fname = block.get("filename", "image.png")
            img_path = f"{images_rel_dir}/{fname}"
            lines.append(f"![图片]({img_path})")
            if add_bbox and block.get("bbox"):
                bbox = block["bbox"]
                page = block.get("page_no", 0) + 1
                bbox_str = ",".join(str(int(v)) for v in bbox)
                lines.append(f"<!-- type:figure | page:{page} | bbox:{bbox_str} -->")

        elif btype == "table":
            fname = block.get("filename")   # 已是 table_001.png
            path  = block.get("path")
            md_table = block.get("md_table", "")
            grid = block.get("grid", [])

            if fname and path and fname.endswith(".png"):
                # 表格已渲染为 PNG，用图片引用（供 Word 插图）
                png_rel = f"{tables_rel_dir}/{fname}"
                lines.append(f"<!-- TABLE:{png_rel} -->")
                lines.append(f"![表格]({png_rel})")

                # 同时把 grid 存为 JSON，供 Step02 标注时提取行数据
                if out_dir is not None and grid:
                    import json as _json
                    tables_dir = out_dir / tables_rel_dir
                    tables_dir.mkdir(parents=True, exist_ok=True)
                    json_fname = fname.replace(".png", ".json")
                    (tables_dir / json_fname).write_text(
                        _json.dumps(grid, ensure_ascii=False, indent=2),
                        encoding="utf-8"
                    )

            elif md_table:
                # 回退：没有PNG时用HTML占位符
                table_counter += 1
                html_fname = f"table_{table_counter:03d}.html"
                placeholder = f"<!-- TABLE:{tables_rel_dir}/{html_fname} -->"
                if out_dir is not None:
                    tables_dir = out_dir / tables_rel_dir
                    tables_dir.mkdir(parents=True, exist_ok=True)
                    (tables_dir / html_fname).write_text(md_table, encoding="utf-8")
                lines.append(placeholder)
            elif fname:
                img_path = f"{images_rel_dir}/{fname}"
                lines.append(f"![表格原图]({img_path})")

            if add_bbox and block.get("bbox"):
                bbox = block["bbox"]
                page = block.get("page_no", 0) + 1
                bbox_str = ",".join(str(int(v)) for v in bbox)
                lines.append(f"<!-- type:table | page:{page} | bbox:{bbox_str} -->")

        elif btype == "formula":
            fname = block.get("filename")
            latex = block.get("latex", "")
            if latex:
                lines.append(latex)
            if fname:
                img_path = f"{images_rel_dir}/{fname}"
                lines.append(f"![公式]({img_path})")
            if add_bbox and block.get("bbox"):
                bbox = block["bbox"]
                page = block.get("page_no", 0) + 1
                bbox_str = ",".join(str(int(v)) for v in bbox)
                lines.append(f"<!-- type:formula | page:{page} | bbox:{bbox_str} -->")

        prev_type = btype

    return "\n".join(lines)


def write_markdown(content: str, out_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(content)


def create_zip(md_path: Path, images_dir: Path, zip_path: Path,
               docx_path: Path = None):
    """打包 MD + images/ + tables/ + 可选 docx"""
    import zipfile
    base = md_path.parent
    tables_dir = base / "tables"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(md_path, md_path.name)
        if images_dir.exists():
            for f in images_dir.iterdir():
                zf.write(f, f"images/{f.name}")
        if tables_dir.exists():
            for f in tables_dir.iterdir():
                zf.write(f, f"tables/{f.name}")
        if docx_path and docx_path.exists():
            zf.write(docx_path, docx_path.name)
