"""
VLM API 客户端
兼容 OpenAI 格式的所有供应商（硅基流动、OpenAI、Azure、Google 等）
及本地模型（Ollama）
"""

import base64
import io
import json
import re
from pathlib import Path
from typing import Optional
from PIL import Image
from openai import OpenAI
from config.settings import get_config


def _get_client() -> tuple[OpenAI, str]:
    """根据当前配置返回 (client, model_name)"""
    cfg = get_config()
    mode = cfg["model_mode"]
    if mode == "online":
        ocfg = cfg["online"]
        client = OpenAI(api_key=ocfg["api_key"], base_url=ocfg["base_url"])
        return client, ocfg["model"]
    else:
        lcfg = cfg["local"]
        client = OpenAI(api_key=lcfg["api_key"], base_url=lcfg["base_url"])
        return client, lcfg["model"]


def _encode_image(image_path: str, max_long_side: int = 1500, jpeg_quality: int = 85) -> str:
    """
    读取图片，自动压缩后返回 base64 字符串。
    - 长边超过 max_long_side 时等比缩放
    - 统一转为 JPEG 输出，减少传输体积
    """
    img = Image.open(image_path)
    if img.mode in ("RGBA", "P", "LA"):
        img = img.convert("RGB")
    w, h = img.size
    long_side = max(w, h)
    if long_side > max_long_side:
        scale = max_long_side / long_side
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=jpeg_quality, optimize=True)
    return base64.b64encode(buf.getvalue()).decode("utf-8")


# ── Prompt 模板 ──────────────────────────────────────────────

LAYOUT_PROMPT = """你是一个专业的文档结构分析专家。请仔细分析这张文档页面图片，识别所有内容区域并判断其语义层级。

请严格按照以下 JSON 格式返回，不要输出任何 JSON 以外的内容（不要有 ```json 标记）：
{
  "page_width": <页面像素宽度>,
  "page_height": <页面像素高度>,
  "blocks": [
    {
      "type": "text|heading|figure|table|formula",
      "bbox": [x1, y1, x2, y2],
      "level": 0,
      "text": "识别到的文字内容"
    }
  ]
}

【标题层级判断规则 - 非常重要】
判断 heading 及其 level 时，综合考虑以下特征：
1. 字号大小：相对于页面正文字号，明显偏大的是标题
2. 字体粗细：加粗文字通常是标题
3. 语义内容：章节编号（第一章、1.1、一、(一)等）、明确的标题词汇
4. 位置特征：居中、独占一行、上下有较大间距
5. 层级映射：
   - level 1：文档大标题、章标题（如"第一章"、"Chapter 1"）
   - level 2：节标题（如"1.1"、"第一节"）
   - level 3：小节标题（如"1.1.1"、条款编号）
   - level 4~6：更细的层级
   - level 0：正文段落，不是标题

【其他规则】
- type 只能是：text、heading、figure、table、formula
- heading 必须填写正确的 level（1-6），正文 text 的 level 填 0
- bbox 为像素坐标 [左边, 上边, 右边, 下边]
- blocks 严格按照从上到下、从左到右的阅读顺序排列
- text 和 heading 的 text 字段填写完整识别文字
- figure/table/formula 的 text 字段填空字符串 ""
- 页眉、页脚、页码不需要识别，直接忽略
- 同一段连续正文合并为一个 text block，不要拆散
"""

TABLE_TO_MD_PROMPT = """请将这张表格图片转换为 HTML 表格格式。

要求：
- 输出标准 HTML 表格语法，使用 <table><tr><td>/<th> 标签
- 第一行如果是表头，使用 <th> 标签，其余行使用 <td> 标签
- 保留所有单元格内容，空单元格保留为 <td></td>
- 合并单元格使用 colspan 和 rowspan 属性表示
- 只输出 HTML 表格代码，不要任何解释文字，不要 ```html 代码块标记
"""

FORMULA_DESCRIBE_PROMPT = """请识别这张图片中的数学/化学公式，用 LaTeX 语法输出。

只输出 LaTeX 公式本身，用 $$ 包裹，例如：$$E = mc^2$$
如果无法识别，输出：$$[公式]$$
"""


def analyze_page_layout(image_path: str, progress_cb=None) -> dict:
    """
    发送页面图片给 VLM，返回版面分析结果 dict
    """
    client, model = _get_client()
    b64 = _encode_image(image_path)

    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                    {"type": "text", "text": LAYOUT_PROMPT},
                ],
            }
        ],
        max_tokens=4096,
        temperature=0.1,
    )

    raw = response.choices[0].message.content.strip()
    # 去掉可能的 markdown 代码块包裹
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # 尝试提取 JSON 对象
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            return json.loads(m.group())
        return {"blocks": [], "error": raw}


def table_image_to_markdown(image_path: str) -> str:
    """将表格图片转为 Markdown 表格文本"""
    client, model = _get_client()
    b64 = _encode_image(image_path)

    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                    {"type": "text", "text": TABLE_TO_MD_PROMPT},
                ],
            }
        ],
        max_tokens=2048,
        temperature=0.1,
    )
    return response.choices[0].message.content.strip()


def formula_image_to_latex(image_path: str) -> str:
    """将公式图片转为 LaTeX 字符串"""
    client, model = _get_client()
    b64 = _encode_image(image_path)

    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                    {"type": "text", "text": FORMULA_DESCRIBE_PROMPT},
                ],
            }
        ],
        max_tokens=512,
        temperature=0.1,
    )
    return response.choices[0].message.content.strip()


def test_connection() -> tuple[bool, str]:
    """测试当前配置的 API 连通性，返回 (成功, 消息)"""
    try:
        client, model = _get_client()
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": "Reply OK"}],
            max_tokens=10,
        )
        return True, f"连接成功，模型：{model}"
    except Exception as e:
        return False, f"连接失败：{e}"
