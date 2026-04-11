"""
Word 导出器
将内容块列表生成 .docx 文件
依赖 Node.js + docx npm 包（npm install -g docx）
"""

import json
import os
import subprocess
import tempfile
from pathlib import Path


# ── JS 模板 ────────────────────────────────────────────────────

_JS_TEMPLATE = r"""
const fs = require('fs');
const path = require('path');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    ImageRun, HeadingLevel, AlignmentType, BorderStyle, WidthType,
    ShadingType, VerticalAlign, TableOfContents
} = require('docx');

const blocks = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const imagesDir = process.argv[3];
const outPath = process.argv[4];

// ── 工具函数 ────────────────────────────────────────────────

const HEADING_MAP = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6,
};

const border = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };

// 解析 HTML 表格为 docx Table
function parseHtmlTable(html) {
    // 提取所有行
    const rowMatches = [...html.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)];
    if (!rowMatches.length) return null;

    // 计算列数
    let maxCols = 0;
    const rowData = rowMatches.map(rowM => {
        const cells = [...rowM[1].matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi)];
        if (cells.length > maxCols) maxCols = cells.length;
        return cells.map(c => {
            const isHeader = c[0].toLowerCase().startsWith('<th');
            const colspanM = c[0].match(/colspan=["']?(\d+)/i);
            const rowspanM = c[0].match(/rowspan=["']?(\d+)/i);
            // 去除内层 HTML 标签，保留纯文本
            const text = c[1].replace(/<[^>]+>/g, '').trim();
            return {
                text,
                isHeader: isHeader,
                colspan: colspanM ? parseInt(colspanM[1]) : 1,
                rowspan: rowspanM ? parseInt(rowspanM[1]) : 1,
            };
        });
    });

    if (maxCols === 0) return null;

    // 计算列宽（A4，内容宽约 9026 DXA，留边距后用 8800）
    const tableWidth = 8800;
    const colWidth = Math.floor(tableWidth / maxCols);
    const columnWidths = Array(maxCols).fill(colWidth);

    const rows = rowData.map((cells, ri) => {
        const tableCells = cells.map(cell => {
            const shading = cell.isHeader
                ? { fill: "D5E8F0", type: ShadingType.CLEAR }
                : { fill: "FFFFFF", type: ShadingType.CLEAR };

            const cellOpts = {
                borders,
                shading,
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                    new Paragraph({
                        children: [new TextRun({
                            text: cell.text,
                            bold: cell.isHeader,
                            size: 20,
                            font: "Arial",
                        })]
                    })
                ],
            };
            if (cell.colspan > 1) cellOpts.columnSpan = cell.colspan;
            if (cell.rowspan > 1) cellOpts.rowSpan = cell.rowspan;
            // 单元格宽度 = colspan * colWidth
            cellOpts.width = { size: colWidth * cell.colspan, type: WidthType.DXA };

            return new TableCell(cellOpts);
        });

        return new TableRow({ children: tableCells });
    });

    return new Table({
        width: { size: tableWidth, type: WidthType.DXA },
        columnWidths,
        rows,
    });
}

// ── 构建 children ────────────────────────────────────────────

const children = [];

for (const block of blocks) {
    const btype = block.type || 'text';

    if (btype === 'heading') {
        const level = Math.max(1, Math.min(6, block.level || 1));
        children.push(new Paragraph({
            heading: HEADING_MAP[level],
            children: [new TextRun({
                text: (block.text || '').trim(),
                font: "Arial",
            })]
        }));

    } else if (btype === 'text') {
        const text = (block.text || '').trim();
        if (text) {
            children.push(new Paragraph({
                children: [new TextRun({ text, font: "Arial", size: 22 })]
            }));
        }

    } else if (btype === 'table') {
        const htmlTable = block.md_table || '';
        if (htmlTable) {
            try {
                const tbl = parseHtmlTable(htmlTable);
                if (tbl) {
                    children.push(tbl);
                    children.push(new Paragraph({ children: [] })); // 表后空行
                }
            } catch (e) {
                children.push(new Paragraph({
                    children: [new TextRun({ text: '[表格解析失败]', color: 'FF0000', font: "Arial" })]
                }));
            }
        } else if (block.filename) {
            // 没有 HTML，插入图片
            const imgPath = path.join(imagesDir, block.filename);
            if (fs.existsSync(imgPath)) {
                const imgData = fs.readFileSync(imgPath);
                const ext = path.extname(block.filename).slice(1).toLowerCase();
                const typeMap = { jpg: 'jpg', jpeg: 'jpg', png: 'png', gif: 'gif', bmp: 'bmp' };
                children.push(new Paragraph({
                    children: [new ImageRun({
                        data: imgData,
                        transformation: { width: 500, height: 200 },
                        type: typeMap[ext] || 'png',
                    })]
                }));
            }
        }

    } else if (btype === 'figure') {
        if (block.filename) {
            const imgPath = path.join(imagesDir, block.filename);
            if (fs.existsSync(imgPath)) {
                const imgData = fs.readFileSync(imgPath);
                const ext = path.extname(block.filename).slice(1).toLowerCase();
                const typeMap = { jpg: 'jpg', jpeg: 'jpg', png: 'png', gif: 'gif', bmp: 'bmp' };
                // 获取图片尺寸，按比例缩放到最大宽度 500pt
                let w = 500, h = 300;
                try {
                    // 简单读取 PNG/JPEG 尺寸
                    if (ext === 'png') {
                        const view = new DataView(imgData.buffer);
                        const pw = imgData.readUInt32BE(16);
                        const ph = imgData.readUInt32BE(20);
                        if (pw > 0 && ph > 0) {
                            const scale = Math.min(500 / pw, 400 / ph, 1);
                            w = Math.round(pw * scale);
                            h = Math.round(ph * scale);
                        }
                    }
                } catch(e) {}
                children.push(new Paragraph({
                    children: [new ImageRun({
                        data: imgData,
                        transformation: { width: w, height: h },
                        type: typeMap[ext] || 'png',
                    })]
                }));
            }
        }

    } else if (btype === 'formula') {
        const latex = (block.latex || '').trim();
        if (latex) {
            children.push(new Paragraph({
                children: [new TextRun({
                    text: latex,
                    font: "Courier New",
                    size: 20,
                    color: "333333",
                })]
            }));
        } else if (block.filename) {
            const imgPath = path.join(imagesDir, block.filename);
            if (fs.existsSync(imgPath)) {
                const imgData = fs.readFileSync(imgPath);
                children.push(new Paragraph({
                    children: [new ImageRun({
                        data: imgData,
                        transformation: { width: 300, height: 60 },
                        type: 'png',
                    })]
                }));
            }
        }
    }
}

// ── 生成文档 ────────────────────────────────────────────────

const doc = new Document({
    styles: {
        default: {
            document: { run: { font: "Arial", size: 22 } }
        },
        paragraphStyles: [
            { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
              run: { size: 32, bold: true, font: "Arial", color: "1F3864" },
              paragraph: { spacing: { before: 300, after: 150 }, outlineLevel: 0 } },
            { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
              run: { size: 28, bold: true, font: "Arial", color: "2E5496" },
              paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
            { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
              run: { size: 24, bold: true, font: "Arial", color: "375623" },
              paragraph: { spacing: { before: 180, after: 90 }, outlineLevel: 2 } },
        ]
    },
    sections: [{
        properties: {
            page: {
                size: { width: 11906, height: 16838 }, // A4
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
            }
        },
        children,
    }]
});

Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(outPath, buf);
    console.log('OK:' + outPath);
}).catch(err => {
    console.error('ERROR:' + err.message);
    process.exit(1);
});
"""


# ── Python 接口 ────────────────────────────────────────────────

def export_word(blocks: list[dict], images_dir: str | Path, out_path: str | Path) -> str:
    """
    将内容块列表导出为 .docx 文件。
    blocks: pipeline 生成的内容块列表
    images_dir: 图片目录路径
    out_path: 输出 .docx 路径
    返回: 输出文件路径字符串（失败时抛出异常）
    """
    images_dir = str(Path(images_dir).resolve())
    out_path = str(Path(out_path).resolve())

    with tempfile.TemporaryDirectory() as tmpdir:
        # 写入块数据 JSON
        blocks_json = Path(tmpdir) / "blocks.json"
        with open(blocks_json, "w", encoding="utf-8") as f:
            json.dump(blocks, f, ensure_ascii=False)

        # 写入 JS 脚本
        js_file = Path(tmpdir) / "export.js"
        js_file.write_text(_JS_TEMPLATE, encoding="utf-8")

        # 查找 docx 模块路径
        npm_prefix = subprocess.run(
            ["npm", "root", "-g"],
            capture_output=True, text=True
        ).stdout.strip()

        env = os.environ.copy()
        if npm_prefix:
            env["NODE_PATH"] = npm_prefix

        result = subprocess.run(
            ["node", str(js_file), str(blocks_json), images_dir, out_path],
            capture_output=True, text=True, env=env
        )

        if result.returncode != 0 or "ERROR:" in result.stdout:
            err = result.stderr or result.stdout
            raise RuntimeError(f"Word 导出失败: {err}")

        if not Path(out_path).exists():
            raise RuntimeError("Word 导出失败：输出文件未生成")

    return out_path
