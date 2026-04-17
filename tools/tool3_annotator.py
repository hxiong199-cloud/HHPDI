"""
Tool3 — MD 数据标注 v3
输出格式（恢复旧格式，修正bug）：
  _annotated.md 格式：
    正文段落
    @@@标签1@@@标签2@@@标签3@@@
    ####
    下一段...

  表格处理：
    原始表格保留
    每行展开为一条自然语言描述，附上 tags
    ####

  在 Dify 中设置分段标识符为 #### 即可精确切分，tags 与正文一起被向量化。
"""

import os, re, json, threading, time, random
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from gui.theme import COLORS, FONTS, PADDING
from gui.widgets import StyledButton, FilePickRow, LogView, ProgressRow, Divider
from config.settings import get_config

# ══════════════════════════════════════════════════════════════
#  全局限速器（跨 worker 共享，防止并发请求同时打爆 API）
# ══════════════════════════════════════════════════════════════

class _RateLimiter:
    """令牌桶：限制每秒最多 max_rps 个请求"""
    def __init__(self, min_interval: float = 0.4):
        self._lock = threading.Lock()
        self._last = 0.0
        self._min_interval = min_interval  # 相邻两次请求最小间隔（秒）

    def acquire(self):
        with self._lock:
            now = time.time()
            wait = self._min_interval - (now - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.time()

_rate_limiter = _RateLimiter(min_interval=0.4)  # 全局单例，≤150 RPM

# ══════════════════════════════════════════════════════════════
#  提示词
# ══════════════════════════════════════════════════════════════

PROMPT_TEXT = """你是一个专业的知识库构建助手。请为下面这段文字生成检索标签。

<所属标题>：__HEADING__
<段落内容>：__PARAGRAPH__

要求：
1. 输出6个以内的标签，格式为JSON对象 {"tags":[...]}
2. 标签包含：所属标题、2~3个核心术语、1~2个自然语言问句（如"该河流发源于哪里"）
3. 只输出JSON，不要代码块，不要其他文字

示例：{"tags": ["流域概况","溆水","流域面积","河流长度","溆水的基本情况是什么","流域面积多大"]}"""

PROMPT_TABLE_BATCH = """你是一个专业的知识库构建助手，负责将表格数据转化为高质量的问答对，用于向量检索知识库。

## 输入信息
<表格标题>：__TITLE__
<表格内容>：
__TABLE__

## 你的任务
对表格中每一条**数据行**（跳过表头行、分隔行、空行），生成一个包含 question、answer、tags 的 JSON 对象。

## 字段规范

### question（问题）
- 找出该行最具标识性的字段值作为"主体"（优先级：姓名 > 名称 > 编号）
- 格式固定为：`{主体}的相关信息是？`
- 示例：`刘小松的相关信息是？`、`溆水的相关信息是？`、`圭洞溪的相关信息是？`

### answer（回答）
- 用**完整流畅的自然语言**写一句话，把该行**所有字段的值**都包含进去，一个都不能省略
- 语言要通顺，像人说话，不要生硬地"字段名：值"罗列
- 示例风格：
  - ✅ "刘小松是男性，学历本科，身份证号****，无职称，工日25天，单价300元/天，合计7500元，岗位系统运维。"
  - ✅ "溆水是沅江的1级支流，流域面积3290km²，河流长度143km，坡降0.191‰，河源位于溆浦县架枧田，流经横坡、双江口等地，河口在大江口。"
  - ❌ "编号：1；干流名称：沅江；支流名称及级别-1级：溆水..."（禁止用冒号分号罗列）

### tags（检索标签）
- 3～5个**简短关键词**，每个不超过10个字
- 包含：主体名称、1～2个核心属性值、1个检索问句
- **严禁**把 answer 本身或超过10字的长句放进 tags
- 示例：`["刘小松", "系统运维", "工日25天", "刘小松的岗位是什么"]`

## 输出规范
- 只输出一个 JSON 数组，数组长度 = 数据行数
- 不要任何解释、不要 markdown 代码块、不要多余文字
- 每个对象必须包含且只包含 question、answer、tags 三个字段

## 完整示例

输入表格（人员信息表）：
| 序号 | 姓名 | 性别 | 学历 | 职称 | 工日(天) | 单价(元/天) | 合计(元) | 岗位 |
|------|------|------|------|------|----------|-------------|----------|------|
| 1 | 刘小松 | 男 | 本科 | 无 | 25 | 300 | 7500 | 系统运维 |
| 2 | 王芳 | 女 | 硕士 | 工程师 | 20 | 500 | 10000 | 项目管理 |

输出：
[
  {"question":"刘小松的相关信息是？","answer":"刘小松是男性，学历本科，无职称，工日25天，单价300元/天，合计7500元，岗位为系统运维。","tags":["刘小松","系统运维","工日25天","刘小松的岗位是什么"]},
  {"question":"王芳的相关信息是？","answer":"王芳是女性，学历硕士，职称工程师，工日20天，单价500元/天，合计10000元，岗位为项目管理。","tags":["王芳","项目管理","工程师","王芳的学历是什么"]}
]"""


# ══════════════════════════════════════════════════════════════
#  表格解析
# ══════════════════════════════════════════════════════════════

def _is_sub_header(row: list) -> bool:
    """
    判断表格某行是否为副表头行（如 ['编号','干流名称','1级','2级','3级',...]）。
    副表头特征：所有非空值都是短标签（≤8字），且没有纯数字或地名。
    """
    import re as _re
    vals = [str(v).strip() for v in row if str(v).strip()]
    if not vals:
        return False
    has_number   = any(_re.fullmatch(r'[\d.,]+', v) for v in vals)
    has_location = any(_re.search(r'县|市|省|村|乡|镇|区', v) for v in vals)
    all_short    = all(len(v) <= 8 for v in vals)
    return all_short and not has_number and not has_location


def _merge_header_rows(row0: list, row1: list) -> list:
    """
    合并两行表头，处理合并单元格展开的情况。
    规则：若 row1[i] 非空且与 row0[i] 不同，则用 row0[i]+row1[i] 拼接；
          若 row1[i] 与 row0[i] 相同或为空，则保留 row0[i]。
    例：row0=['支流名称及级别', ...]  row1=['1级', ...] → '支流名称及级别-1级'
    """
    result = []
    ncols = max(len(row0), len(row1))
    for i in range(ncols):
        h0 = str(row0[i]).strip() if i < len(row0) else ''
        h1 = str(row1[i]).strip() if i < len(row1) else ''
        if h1 and h1 != h0:
            result.append(f"{h0}-{h1}" if h0 else h1)
        else:
            result.append(h0)
    return result


def _parse_pipe_table(lines):
    """解析pipe表格，返回 (headers, data_rows)"""
    rows = []
    for line in lines:
        line = line.strip()
        if not line or not line.startswith('|'):
            continue
        if re.match(r'^\|[\s\-:]+\|', line):
            continue
        cells = [c.strip() for c in line.strip('|').split('|')]
        rows.append(cells)

    if len(rows) < 2:
        return [], []

    headers = rows[0]
    data_rows = []
    for row in rows[1:]:
        while len(row) < len(headers):
            row.append('')
        d = {}
        for i, h in enumerate(headers):
            if h:
                d[h] = row[i] if i < len(row) else ''
        data_rows.append(d)

    return headers, data_rows


def _table_row_to_text(headers, row):
    """把一行转成 '字段：值；字段：值' 格式。row 可以是列表或字典"""
    if isinstance(row, dict):
        parts = [f"{h}：{row.get(h,'')}"
                 for h in headers if h and row.get(h, '')]
    else:
        # row 是列表，跳过值与表头名相同的列（表头重复行）
        parts = []
        for i, v in enumerate(row):
            if not v or not v.strip():
                continue
            h = headers[i] if i < len(headers) else f"字段{i+1}"
            if not h:
                continue
            if v.strip() == h.strip():  # 跳过表头重复行（如"编号：编号"）
                continue
            parts.append(f"{h}：{v}")
    return '；'.join(parts) if parts else ''


# ══════════════════════════════════════════════════════════════
#  MD 解析
# ══════════════════════════════════════════════════════════════

def _parse_units(content, min_len=30):
    raw = content.split('\n')
    units = []
    current_heading = '（无标题）'
    i = 0

    while i < len(raw):
        line = raw[i]
        stripped = line.strip()

        # <!-- TABLE:tables/table_001.png --> 占位符（Step01新格式）
        m_tbl = re.match(r'^<!--\s*TABLE:(.*?)\s*-->$', stripped)
        if m_tbl:
            tbl_ref = m_tbl.group(1).strip()
            start = i
            # 跳过紧随其后的 ![表格](...) 行
            i += 1
            if i < len(raw) and re.match(r'^!\[.*\]\(.*\)', raw[i].strip()):
                i += 1
            title = _last_text_title(units, current_heading)
            units.append({'type': 'table', 'fmt': 'png_ref',
                          'title': title, 'heading': current_heading,
                          'lines': [line], 'start': start, 'end': i - 1,
                          'row_count': 5,  # 占位，批量LLM会处理
                          'tbl_ref': tbl_ref,
                          'headers': [], 'data_rows': []})
            continue

        # HTML表格
        if re.search(r'<table', stripped, re.IGNORECASE):
            start = i
            tlines = []
            while i < len(raw):
                tlines.append(raw[i])
                if re.search(r'</table', raw[i], re.IGNORECASE):
                    i += 1
                    break
                i += 1
            row_count = len(re.findall(r'<tr', '\n'.join(tlines), re.IGNORECASE)) - 1
            row_count = max(row_count, 1)
            title = _last_text_title(units, current_heading)
            units.append({'type': 'table', 'fmt': 'html',
                          'title': title, 'heading': current_heading,
                          'lines': tlines, 'start': start, 'end': i - 1,
                          'row_count': row_count,
                          'headers': [], 'data_rows': []})
            continue

        # Pipe 表格
        if stripped.startswith('|') and '|' in stripped[1:]:
            start = i
            tlines = []
            while i < len(raw) and raw[i].strip().startswith('|'):
                tlines.append(raw[i])
                i += 1
            headers, data_rows = _parse_pipe_table(tlines)
            row_count = max(len(data_rows) - 1, 1) if data_rows else 1
            title = _last_text_title(units, current_heading)
            units.append({'type': 'table', 'fmt': 'pipe',
                          'title': title, 'heading': current_heading,
                          'lines': tlines, 'start': start, 'end': i - 1,
                          'row_count': row_count,
                          'headers': headers, 'data_rows': data_rows})
            continue

        # 空行
        if not stripped:
            i += 1
            continue

        # 文本段
        start = i
        plines = []
        while i < len(raw):
            cur = raw[i].strip()
            if not cur:
                break
            if re.search(r'<table', cur, re.IGNORECASE):
                break
            if cur.startswith('|') and '|' in cur[1:]:
                break
            if re.match(r'^<!--\s*TABLE:', cur):
                break
            # 跳过标注行（处理已标注文件时不吃进正文）
            if cur == '####':
                break
            if cur.startswith('@@@') and cur.endswith('@@@'):
                i += 1; continue
            if re.match(r'^tags:\s*@@@', cur):
                i += 1; continue
            plines.append(raw[i])
            i += 1

        para = ' '.join(l.strip() for l in plines if l.strip())
        if not para:
            continue

        first = plines[0].strip()
        if re.match(r'^#{1,6}\s+', first):
            current_heading = re.sub(r'^#{1,6}\s+', '', first).strip()

        units.append({'type': 'text', 'para': para,
                      'heading': current_heading,
                      'lines': plines, 'start': start, 'end': i - 1})

    # 合并短碎片
    merged = []
    for u in units:
        if (u['type'] == 'text' and merged
                and merged[-1]['type'] == 'text'
                and len(u['para']) < min_len
                and not re.match(r'^#{1,6}\s+', u['para'])):
            merged[-1]['para'] += ' ' + u['para']
            merged[-1]['lines'].extend(u['lines'])
            merged[-1]['end'] = u['end']
        else:
            merged.append(u)
    return merged


def _last_text_title(units, fallback):
    for u in reversed(units):
        if u['type'] == 'text':
            return u['para'][:80]
    return fallback


# ══════════════════════════════════════════════════════════════
#  输出重建（旧格式：#### 分割 + @@@ 标签）
# ══════════════════════════════════════════════════════════════

def _html_table_to_rows(html_text):
    """从HTML表格提取 (headers, data_rows)，不依赖bs4"""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_text, "html.parser")
        table = soup.find("table")
        if not table:
            return [], []
        rows = table.find_all("tr")
        if not rows:
            return [], []
        headers = [td.get_text(strip=True) for td in rows[0].find_all(["th", "td"])]
        data_rows = []
        for tr in rows[1:]:
            cells = [td.get_text(strip=True) for td in tr.find_all(["th", "td"])]
            if any(c for c in cells):
                data_rows.append(cells)
        return headers, data_rows
    except Exception:
        # 回退：正则提取
        rows = re.findall(r'<tr[^>]*>(.*?)</tr>', html_text, re.IGNORECASE | re.DOTALL)
        result = []
        for r in rows:
            cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', r, re.IGNORECASE | re.DOTALL)
            cells = [re.sub(r'<[^>]+>', '', c).strip() for c in cells]
            result.append(cells)
        return (result[0] if result else []), result[1:]


def _pipe_rows_to_qa(headers, data_rows):
    """把pipe表格的行数据转为 Q:...;A:... 格式列表"""
    qa_list = []
    for row in data_rows:
        if not any(c.strip() for c in row):
            continue
        # 找出非空字段
        pairs = []
        for ci, val in enumerate(row):
            val = val.strip()
            if not val:
                continue
            hdr = headers[ci].strip() if ci < len(headers) else f"字段{ci+1}"
            pairs.append(f"{hdr}：{val}")
        if pairs:
            qa_text = "；".join(pairs)
            qa_list.append(qa_text)
    return qa_list


def _extract_question(answer_text: str) -> str:
    """
    从 answer 文本里自动提取主体，生成问题。
    优先取第一个「字段：值」对里的值作为主体。
    例："编号：1；干流名称：沅江；支流名称及级别-1级：溆水；..."
        → 主体="沅江" → "沅江的相关信息是？"
    例："刘小松是男性，学历本科，..."
        → 主体="刘小松" → "刘小松的相关信息是？"
    """
    import re as _re
    # 情况1：字段：值 格式，取第一个有意义的值（跳过纯数字编号）
    pairs = _re.findall(r'[\w（）/\-]+[：:]([\w\s（）、，,·.·]+?)(?:[；;]|$)', answer_text)
    for val in pairs:
        val = val.strip()
        if val and not _re.fullmatch(r'[\d.,]+', val) and len(val) <= 12:
            return f"{val}的相关信息是？"
    # 情况2：自然语言，取开头的人名/地名（2~5字汉字）
    m = _re.match(r'^([^\s，,是的有在为与]{2,6})[是的有在为与]', answer_text)
    if m:
        return f"{m.group(1)}的相关信息是？"
    # 兜底
    return "该记录的相关信息是？"


def _auto_tags_from_qa(question: str, answer: str) -> list:
    """
    当LLM没有返回tags或全被过滤时，从question和answer自动生成tags。
    目标3-5个，包含：主体名、核心数值、地名、检索问句。
    """
    import re as _re
    tags = []

    # 1. 主体名称（从question提取）
    subject = question.replace('的相关信息是？', '').replace('是什么？', '').strip()
    if subject:
        tags.append(subject)

    # 2. 从answer提取有意义的字段值
    if '：' in answer or ':' in answer:
        # 字段：值 格式
        pairs = _re.findall(r'([^：:；;，,\n]+)[：:]\s*([^；;：:，,\n]+)', answer)
        added = 0
        for field, val in pairs:
            field = field.strip()
            val   = val.strip()
            if not val or val == subject:
                continue
            # 跳过纯数字编号（无单位）
            if _re.fullmatch(r'\d+', val):
                continue
            # 字段名含单位时，把字段名里的单位带上（如"流域面积（km2）：3290" → "3290km2"）
            unit_in_field = _re.search(r'[（(]([^）)]+)[）)]', field)
            if unit_in_field and _re.fullmatch(r'[\d.]+', val):
                val = val + unit_in_field.group(1)
            # 优先：带单位的数值、地名
            if (_re.search(r'[\d.]+(?:km2?²?|m2?²?|‰|%|元|天|年|亩|万)', val) or
                    _re.search(r'县|市|省|乡|镇|村', val)):
                if len(val) <= 20 and val not in tags:
                    tags.append(val)
                    added += 1
                    if added >= 3:
                        break
    else:
        # 自然语言：提取带单位数字片段
        nums = _re.findall(r'\d+(?:\.\d+)?(?:km²?|‰|元|天|km|m|亩|万)', answer)
        for n in nums[:2]:
            if n not in tags:
                tags.append(n)

    # 3. question本身作为检索问句
    if question not in tags:
        tags.append(question)

    return tags[:5]


def _rebuild(orig_lines, units, results):
    """
    重建 annotated.md。
    输出格式（Dify标准）：
        ####
        原文正文（标题/段落/图片原样）
        tags: @@@tag1@@@tag2@@@tag3@@@
        ####

    表格格式：
        ####
        <!-- TABLE:tables/table_001.png -->
        ![表格](tables/table_001.png)
        ####
        ####
        Q:字段1：值1；字段2：值2
        tags: @@@tag1@@@tag2@@@
        ####
        ...（每行一个chunk）
    """
    line_to_unit = {}
    for idx, u in enumerate(units):
        for ln in range(u['start'], u['end'] + 1):
            line_to_unit[ln] = idx

    out = []
    done_units = set()
    i = 0

    def add_chunk(content_lines, tags):
        """
        输出一个 Dify chunk。格式：
            ####
            内容行
            tags: @@@...@@@
        相邻 chunk 共用分隔符，最终在文件末尾由 _rebuild 补一个收尾 ####。
        """
        out.append('####')
        for cl in content_lines:
            out.append(cl)
        if tags:
            out.append('tags: @@@' + '@@@'.join(tags) + '@@@')

    while i < len(orig_lines):
        uid = line_to_unit.get(i)
        if uid is not None and uid not in done_units:
            done_units.add(uid)
            unit = units[uid]
            result = results[uid]

            if unit['type'] == 'text':
                content_lines = [l for l in unit['lines'] if l.strip()]
                tags = result.get('tags', []) if result else []
                add_chunk(content_lines, tags)
                i = unit['end'] + 1

            elif unit['type'] == 'table':
                fmt = unit.get('fmt', '')

                # ① 表格图片引用 chunk（供读者看，也供 Word 插图）
                if fmt == 'png_ref':
                    tbl_ref = unit.get('tbl_ref', '')
                    fname = tbl_ref.split('/')[-1] if tbl_ref else ''
                    ref_lines = [
                        f"<!-- TABLE:{tbl_ref} -->",
                        f"![表格]({tbl_ref})",
                    ] if tbl_ref else unit['lines']
                    add_chunk(ref_lines, [])

                    # ② 每行 QA chunk
                    row_results = result.get('row_results', []) if result else []
                    for rr in row_results:
                        question = rr.get('question', '').strip()
                        answer   = (rr.get('answer', '') or rr.get('description', '') or
                                    rr.get('qa', '')).strip()
                        # tags 清洗：过滤掉超过25字的长文本
                        raw_tags = rr.get('tags', [])
                        tags = [t for t in raw_tags if isinstance(t, str) and 0 < len(t) <= 25]
                        if not answer:
                            continue
                        if not question:
                            question = _extract_question(answer)
                        # tags 兜底：LLM没返回或全被过滤，自动生成
                        if not tags:
                            tags = _auto_tags_from_qa(question, answer)
                        add_chunk([f"Q:{question}", f"A:{answer}"], tags)

                elif fmt in ('pipe', 'html'):
                    add_chunk(unit['lines'], [])

                    row_results = result.get('row_results', []) if result else []
                    if row_results:
                        for rr in row_results:
                            question = rr.get('question', '').strip()
                            answer   = (rr.get('answer', '') or rr.get('description', '') or
                                        rr.get('qa', '')).strip()
                            raw_tags = rr.get('tags', [])
                            tags = [t for t in raw_tags if isinstance(t, str) and 0 < len(t) <= 25]
                            if not answer:
                                continue
                            if not question:
                                question = _extract_question(answer)
                            if not tags:
                                tags = _auto_tags_from_qa(question, answer)
                            add_chunk([f"Q:{question}", f"A:{answer}"], tags)
                    else:
                        headers   = unit.get('headers', [])
                        data_rows = unit.get('data_rows', [])
                        if not headers and not data_rows:
                            headers, data_rows = _parse_pipe_table(unit['lines']) \
                                if fmt == 'pipe' else ([], [])
                        for row in data_rows:
                            answer = _table_row_to_text(headers, row)
                            if answer:
                                question = _extract_question(answer)
                                tags = _auto_tags_from_qa(question, answer)
                                add_chunk([f"Q:{question}", f"A:{answer}"], tags)

                i = unit['end'] + 1
        else:
            line = orig_lines[i]
            s = line.strip()
            # 跳过残留的旧格式标注行，避免重复
            if s == '####' or (s.startswith('@@@') and s.endswith('@@@')) \
                    or re.match(r'^tags:\s*@@@', s):
                i += 1
                continue
            out.append(line)
            i += 1

    # 收尾 ####（最后一个chunk的结束符）
    # 去掉末尾多余空行后补上
    while out and out[-1].strip() == '':
        out.pop()
    if out and out[-1] != '####':
        out.append('####')

    return out


# ══════════════════════════════════════════════════════════════
#  Tool3 面板
# ══════════════════════════════════════════════════════════════

class Tool3Panel(tk.Frame):
    def __init__(self, parent, shared_state: dict, status_bar,
                 navigate_cb=None, **kw):
        bg = kw.pop('bg', COLORS['bg_main'])
        super().__init__(parent, bg=bg, **kw)
        self._shared     = shared_state
        self._status_bar = status_bar
        self._navigate   = navigate_cb
        self._is_running = False
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=COLORS['bg_card'])
        hdr.pack(fill='x')
        if self._navigate:
            tk.Button(hdr, text='◈  主页',
                      bg=COLORS['bg_card'], fg=COLORS['accent'],
                      activebackground=COLORS['bg_hover'],
                      activeforeground=COLORS['accent'],
                      relief='flat', bd=0, cursor='hand2',
                      font=FONTS['sm'], padx=PADDING['md'],
                      command=lambda: self._navigate('home')
                      ).pack(side='right', pady=PADDING['sm'],
                             padx=PADDING['md'])
        tk.Label(hdr, text='数据标注',
                 bg=COLORS['bg_card'], fg=COLORS['tool3_color'],
                 font=FONTS['h1']).pack(side='left',
                                        padx=PADDING['xl'], pady=PADDING['lg'])
        tk.Label(hdr, text='Markdown → Dify 知识库（#### 分割 + @@@ 标签）',
                 bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                 font=FONTS['md']).pack(side='left', pady=PADDING['lg'])
        Divider(self).pack(fill='x')

        body = tk.Frame(self, bg=COLORS['bg_main'])
        body.pack(fill='both', expand=True)
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(2, weight=1)

        # 配置区
        cfg_frame = tk.LabelFrame(body, text=' 模型配置 ',
                                   bg=COLORS['bg_card'],
                                   fg=COLORS['tool3_color'],
                                   font=FONTS['h3'], relief='flat',
                                   highlightthickness=1,
                                   highlightbackground=COLORS['border'])
        cfg_frame.grid(row=0, column=0, columnspan=2, sticky='ew',
                       padx=PADDING['lg'], pady=PADDING['md'])
        inner_cfg = tk.Frame(cfg_frame, bg=COLORS['bg_card'])
        inner_cfg.pack(fill='x', padx=PADDING['md'], pady=PADDING['md'])
        inner_cfg.columnconfigure(1, weight=1)

        def lbl(t, r):
            tk.Label(inner_cfg, text=t, bg=COLORS['bg_card'],
                     fg=COLORS['text_secondary'],
                     font=FONTS['md'], width=12, anchor='w'
                     ).grid(row=r, column=0, sticky='w', pady=3)

        lbl('API URL', 0)
        self._url_entry = tk.Entry(inner_cfg,
                                    bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                                    insertbackground=COLORS['text_primary'],
                                    relief='flat', bd=0, highlightthickness=1,
                                    highlightbackground=COLORS['border'],
                                    highlightcolor=COLORS['tool3_color'],
                                    font=FONTS['mono_sm'])
        self._url_entry.grid(row=0, column=1, sticky='ew', padx=(8,0), pady=3, ipady=6)
        self._url_entry.insert(0, 'https://api.siliconflow.cn/v1/chat/completions')

        lbl('API Key', 1)
        key_row = tk.Frame(inner_cfg, bg=COLORS['bg_card'])
        key_row.grid(row=1, column=1, sticky='ew', padx=(8,0), pady=3)
        self._key_entry = tk.Entry(key_row, bg=COLORS['bg_input'],
                                    fg=COLORS['text_primary'],
                                    insertbackground=COLORS['text_primary'],
                                    relief='flat', bd=0, highlightthickness=1,
                                    highlightbackground=COLORS['border'],
                                    highlightcolor=COLORS['tool3_color'],
                                    font=FONTS['mono_sm'], show='•')
        self._key_entry.pack(side='left', fill='x', expand=True, ipady=6)
        sv = tk.BooleanVar(value=False)
        tk.Checkbutton(key_row, text='显示', variable=sv,
                       bg=COLORS['bg_card'], fg=COLORS['text_muted'],
                       activebackground=COLORS['bg_card'],
                       selectcolor=COLORS['bg_input'], font=FONTS['xs'],
                       command=lambda: self._key_entry.config(
                           show='' if sv.get() else '•')
                       ).pack(side='right', padx=(6,0))

        lbl('模型', 2)
        self._model_var = tk.StringVar(value='Pro/deepseek-ai/DeepSeek-V3')
        ttk.Combobox(inner_cfg, textvariable=self._model_var,
                     values=['Pro/deepseek-ai/DeepSeek-V3',
                             'deepseek-ai/DeepSeek-V3',
                             'Qwen/Qwen2.5-72B-Instruct',
                             'Qwen/Qwen2.5-32B-Instruct',
                             'Pro/Qwen/Qwen2.5-7B-Instruct',
                             'deepseek-ai/DeepSeek-R1'],
                     font=FONTS['sm'], style='Dark.TCombobox', width=38
                     ).grid(row=2, column=1, sticky='ew', padx=(8,0), pady=3)

        lbl('并发数', 3)
        self._concur_var = tk.IntVar(value=3)
        cr = tk.Frame(inner_cfg, bg=COLORS['bg_card'])
        cr.grid(row=3, column=1, sticky='w', padx=(8,0), pady=3)
        tk.Spinbox(cr, from_=1, to=10, textvariable=self._concur_var,
                   bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                   relief='flat', bd=0, highlightthickness=1,
                   highlightbackground=COLORS['border'],
                   font=FONTS['sm'], width=5).pack(side='left')
        tk.Label(cr, text='个任务同时处理',
                 bg=COLORS['bg_card'], fg=COLORS['text_muted'],
                 font=FONTS['sm']).pack(side='left', padx=(6,0))

        self._load_from_global_config()

        # 文件区
        ff = tk.LabelFrame(body, text=' 输入文件 ',
                            bg=COLORS['bg_card'], fg=COLORS['tool3_color'],
                            font=FONTS['h3'], relief='flat',
                            highlightthickness=1,
                            highlightbackground=COLORS['border'])
        ff.grid(row=1, column=0, columnspan=2, sticky='ew',
                padx=PADDING['lg'], pady=(0, PADDING['md']))
        inner_ff = tk.Frame(ff, bg=COLORS['bg_card'])
        inner_ff.pack(fill='x', padx=PADDING['md'], pady=PADDING['md'])

        self._file_row = FilePickRow(inner_ff, 'MD 文件',
                                     self._browse_file,
                                     COLORS['tool3_color'])
        self._file_row.pack(fill='x', pady=(0, PADDING['sm']))

        bottom_row = tk.Frame(inner_ff, bg=COLORS['bg_card'])
        bottom_row.pack(fill='x')
        self._file_info = tk.Label(bottom_row, text='文件统计: —',
                                    bg=COLORS['bg_card'],
                                    fg=COLORS['text_muted'], font=FONTS['sm'])
        self._file_info.pack(side='left')
        self._stop_btn = StyledButton(bottom_row, text='停止',
                                       style='danger', command=self._stop)
        self._stop_btn.pack(side='right', padx=(PADDING['sm'],0))
        self._stop_btn.config(state='disabled')
        self._start_btn = StyledButton(bottom_row, text='开始标注',
                                        style='orange', command=self._start_thread)
        self._start_btn.pack(side='right', padx=(PADDING['sm'],0))

        # 日志区
        log_frame = tk.Frame(body, bg=COLORS['bg_card'])
        log_frame.grid(row=2, column=0, columnspan=2, sticky='nsew',
                       padx=PADDING['lg'], pady=(0, PADDING['lg']))
        log_frame.grid_rowconfigure(1, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)
        tk.Label(log_frame, text='执行日志',
                 bg=COLORS['bg_card'], fg=COLORS['text_secondary'],
                 font=FONTS['xs']).grid(row=0, column=0, sticky='w')
        self._log = LogView(log_frame, height=12)
        self._log.grid(row=1, column=0, sticky='nsew', pady=(2,0))

    def _load_from_global_config(self):
        try:
            cfg = get_config()
            llm = cfg.get('llm', {})
            if llm.get('api_key'):
                self._key_entry.delete(0, 'end')
                self._key_entry.insert(0, llm['api_key'])
            if llm.get('base_url'):
                url = llm['base_url'].rstrip('/') + '/chat/completions'
                self._url_entry.delete(0, 'end')
                self._url_entry.insert(0, url)
            if llm.get('model'):
                self._model_var.set(llm['model'])
        except Exception:
            pass

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title='选择 Markdown 文件',
            filetypes=[('Markdown', '*.md'), ('所有文件', '*.*')])
        if path:
            self._file_row.set(path)
            self._update_file_info(path)
        elif self._shared.get('last_md_path'):
            path = self._shared['last_md_path']
            self._file_row.set(path)
            self._update_file_info(path)

    def _update_file_info(self, path):
        try:
            content = open(path, encoding='utf-8').read()
            units = _parse_units(content, 30)
            text_cnt  = sum(1 for u in units if u['type'] == 'text')
            table_cnt = sum(1 for u in units if u['type'] == 'table')
            total_rows = sum(u.get('row_count', 0)
                             for u in units if u['type'] == 'table')
            self._file_info.config(
                text=f'文本段: {text_cnt}  |  表格: {table_cnt}  |  表格行: {total_rows}')
        except Exception as e:
            self._file_info.config(text=f'读取失败: {e}')

    def _start_thread(self):
        path = self._file_row.get()
        if not path:
            path = self._shared.get('last_md_path', '')
            if path: self._file_row.set(path)
        if not path:
            messagebox.showerror('错误', '请选择 MD 文件', parent=self)
            return
        if not self._key_entry.get().strip():
            self._load_from_global_config()
        if not self._key_entry.get().strip():
            messagebox.showerror('错误',
                                  '请输入 API Key\n可在「模型设置」中统一配置',
                                  parent=self)
            return
        self._is_running = True
        self._start_btn.config(state='disabled')
        self._stop_btn.config(state='normal')
        self._log.clear()
        self._status_bar.set('数据标注处理中…', 'running')
        threading.Thread(target=self._worker, args=(path,), daemon=True).start()

    def _stop(self):
        self._is_running = False
        self._log.append('用户请求停止…', 'WARNING')
        self._stop_btn.config(state='disabled')
        self._status_bar.set('已停止', 'warning')

    def _worker(self, file_path):
        import traceback
        try:
            content = open(file_path, encoding='utf-8').read()
            orig_lines = content.split('\n')
            units = _parse_units(content, 30)

            text_cnt  = sum(1 for u in units if u['type'] == 'text')
            table_cnt = sum(1 for u in units if u['type'] == 'table')
            total_rows = sum(len(u.get('data_rows', [])) or u.get('row_count', 0)
                             for u in units if u['type'] == 'table')
            self._log.append(
                f'解析：{text_cnt} 文本段，{table_cnt} 表格，{total_rows} 表格数据行', 'INFO')

            results = [None] * len(units)
            total = len(units)
            done = [0]

            def _process(idx):
                if not self._is_running:
                    return idx, None
                unit = units[idx]
                try:
                    if unit['type'] == 'text':
                        return idx, self._process_text(unit)
                    else:
                        return idx, self._process_table(unit, md_path=file_path)
                except Exception as e:
                    self._log.append(f'  单元{idx+1}失败: {e}', 'ERROR')
                    return idx, None

            concur = self._concur_var.get()
            with ThreadPoolExecutor(max_workers=concur) as pool:
                futures = {pool.submit(_process, i): i for i in range(total)}
                for fut in as_completed(futures):
                    if not self._is_running:
                        break
                    idx, res = fut.result()
                    results[idx] = res
                    done[0] += 1
                    pct = done[0] / total * 100
                    unit = units[idx]
                    if unit['type'] == 'text':
                        tags = res.get('tags', []) if res else []
                        self._log.append(
                            f'  [{done[0]}/{total}] 文本 → {tags[:3]}', 'SUCCESS')
                    else:
                        rr = res.get('row_results', []) if res else []
                        self._log.append(
                            f'  [{done[0]}/{total}] 表格 → {len(rr)} 行已处理', 'SUCCESS')
                    self._status_bar.set(f'标注中 {int(pct)}%…', 'running')

            if not self._is_running:
                self._log.append('已停止', 'WARNING')
                return

            out_lines = _rebuild(orig_lines, units, results)

            from pathlib import Path
            base = Path(file_path)
            out_path = base.parent / (base.stem + '_annotated.md')
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(out_lines))

            chunk_count = out_lines.count('####')
            self._log.append('─' * 40, 'DIM')
            self._log.append(f'✓ 输出：{out_path}', 'SUCCESS')
            self._log.append(
                f'✓ 共 {chunk_count} 个 chunk（Dify 分段标识符设为 #### ）', 'GOLD')
            self._status_bar.set('标注完成 ✓', 'success')
            self.after(0, lambda: messagebox.showinfo(
                '标注完成',
                f'已生成：{out_path.name}\n'
                f'共 {chunk_count} 个 chunk\n\n'
                f'上传到 Dify 时，将分段标识符设为：####',
                parent=self))

        except Exception as e:
            self._log.append(f'异常：{traceback.format_exc()[:400]}', 'ERROR')
            self._status_bar.set('标注失败', 'error')
        finally:
            self._is_running = False
            self.after(0, lambda: (
                self._start_btn.config(state='normal'),
                self._stop_btn.config(state='disabled')))

    def _call_llm(self, system_prompt, user_msg):
        import requests as rq
        import time
        url   = self._url_entry.get().strip()
        key   = self._key_entry.get().strip()
        model = self._model_var.get().strip()
        h = {'Content-Type': 'application/json',
             'Authorization': f'Bearer {key}'}
        p = {'model': model,
             'messages': [{'role': 'system', 'content': system_prompt},
                           {'role': 'user',   'content': user_msg}],
             'max_tokens': 2048, 'temperature': 0}
        max_retries = 4
        wait = 3
        last_exc = None
        for attempt in range(max_retries):
            _rate_limiter.acquire()  # 全局限速，防止并发爆 API
            try:
                r = rq.post(url, headers=h, json=p, timeout=300)
                if r.status_code == 429:
                    # 限流：加抖动后重试（避免所有 worker 同步重试）
                    jitter = random.uniform(0, wait)
                    time.sleep(wait + jitter)
                    wait *= 2
                    continue
                r.raise_for_status()
                raw = r.json()['choices'][0]['message']['content'].strip()
                raw = re.sub(r'^```(?:json)?\s*', '', raw)
                raw = re.sub(r'\s*```$', '', raw).strip()
                return raw
            except Exception as exc:
                last_exc = exc
                if attempt < max_retries - 1:
                    jitter = random.uniform(0, wait)
                    time.sleep(wait + jitter)
                    wait *= 2
        raise last_exc

    def _process_text(self, unit):
        sys_p = (PROMPT_TEXT
                 .replace('__HEADING__',   unit['heading'])
                 .replace('__PARAGRAPH__', unit['para']))
        raw = self._call_llm(sys_p, '只输出JSON对象{"tags":[...]}')
        try:
            tags = json.loads(raw).get('tags', [])
        except Exception:
            tags = ['解析失败']
        return {'type': 'text', 'tags': tags[:6]}

    def _process_table(self, unit, md_path: str = ""):
        """表格处理：整表一次批量发给LLM，返回每行的QA+tags列表"""
        fmt       = unit.get('fmt', '')
        headers   = unit.get('headers', [])
        data_rows = unit.get('data_rows', [])

        # png_ref 类型：从旁边的 .json 文件加载 grid 数据
        if fmt == 'png_ref' and not data_rows:
            import json as _j2, re as _re2
            from pathlib import Path as _P
            tbl_ref = unit.get('tbl_ref', '')
            if tbl_ref and md_path:
                json_path = str(_P(md_path).parent / tbl_ref.replace('.png', '.json'))
                if _P(json_path).exists():
                    try:
                        grid = _j2.loads(_P(json_path).read_text(encoding='utf-8'))
                        if grid and len(grid) > 1:
                            raw_headers = grid[0]
                            if len(grid) > 2 and _is_sub_header(grid[1]):
                                # 完全的副表头：直接用行1
                                headers   = grid[1]
                                data_rows = grid[2:]
                            elif len(grid) > 2 and not _is_sub_header(grid[1]) \
                                    and not any(
                                        __import__('re').fullmatch(r'[\d.,]+', v)
                                        for v in grid[1] if v.strip()):
                                # 行1是展开的表头行（合并单元格）：合并行0和行1
                                headers   = _merge_header_rows(raw_headers, grid[1])
                                data_rows = grid[2:]
                            else:
                                headers   = raw_headers
                                data_rows = grid[1:]
                    except Exception:
                        pass

        if not data_rows:
            sys_p = (PROMPT_TEXT
                     .replace('__HEADING__',   unit['heading'])
                     .replace('__PARAGRAPH__', unit['title']))
            raw = self._call_llm(sys_p, '只输出JSON对象{"tags":[...]}')
            try:
                tags = json.loads(raw).get('tags', [])
            except Exception:
                tags = [unit['title']]
            return {'type': 'table', 'tags': tags, 'row_results': []}

        # 整表一次请求
        if fmt == 'png_ref':
            table_text = '\n'.join(['\t'.join(str(c) for c in r)
                                    for r in [headers] + list(data_rows)])
        else:
            table_text = '\n'.join(unit['lines'])

        sys_p = (PROMPT_TABLE_BATCH
                 .replace('__TITLE__', unit['title'])
                 .replace('__TABLE__', table_text))
        row_count = len(data_rows)
        try:
            raw = self._call_llm(
                sys_p,
                f'只输出JSON数组，共{row_count}个对象，每行数据对应一个，完整保留所有字段。')
            batch = json.loads(raw)
            if not isinstance(batch, list):
                raise ValueError('not a list')
        except Exception:
            # 规则兜底：无法生成自然语言，先用字段罗列作answer，question留空让_rebuild兼容
            batch = [{'question': '', 'answer': _table_row_to_text(headers, row),
                      'tags': [unit['title']]}
                     for row in data_rows]

        if len(batch) < row_count:
            for row in data_rows[len(batch):]:
                batch.append({'question': '', 'answer': _table_row_to_text(headers, row),
                               'tags': [unit['title']]})

        row_results = []
        for item in batch:
            if isinstance(item, dict):
                row_results.append({
                    'question': item.get('question', '').strip(),
                    'answer':   item.get('answer',   item.get('description', '')).strip(),
                    'tags':     item.get('tags', [unit['title']])[:5],
                })

        return {'type': 'table', 'row_results': row_results}
