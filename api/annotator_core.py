"""
HHPDI API — Standalone Annotation Core
完全独立，不依赖 tkinter / gui 模块。
纯函数从 tools/tool3_annotator.py 逐字复制，仅移除 GUI 绑定部分。
"""
from __future__ import annotations

import json
import random
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Callable, Optional

# ══════════════════════════════════════════════════════════════
#  全局限速器
# ══════════════════════════════════════════════════════════════

class _RateLimiter:
    def __init__(self, min_interval: float = 0.4):
        self._lock = threading.Lock()
        self._last = 0.0
        self._min_interval = min_interval

    def acquire(self):
        with self._lock:
            now = time.time()
            wait = self._min_interval - (now - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.time()

_rate_limiter = _RateLimiter(min_interval=0.4)

# ══════════════════════════════════════════════════════════════
#  提示词（与 tool3_annotator.py 保持一致）
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

### tags（检索标签）
- 3～5个**简短关键词**，每个不超过10个字
- 包含：主体名称、1～2个核心属性值、1个检索问句
- **严禁**把 answer 本身或超过10字的长句放进 tags

## 输出规范
- 只输出一个 JSON 数组，数组长度 = 数据行数
- 不要任何解释、不要 markdown 代码块、不要多余文字
- 每个对象必须包含且只包含 question、answer、tags 三个字段"""


# ══════════════════════════════════════════════════════════════
#  纯函数（从 tool3_annotator.py 复制，无 GUI 依赖）
# ══════════════════════════════════════════════════════════════

def _is_sub_header(row: list) -> bool:
    vals = [str(v).strip() for v in row if str(v).strip()]
    if not vals:
        return False
    has_number   = any(re.fullmatch(r'[\d.,]+', v) for v in vals)
    has_location = any(re.search(r'县|市|省|村|乡|镇|区', v) for v in vals)
    all_short    = all(len(v) <= 8 for v in vals)
    return all_short and not has_number and not has_location


def _merge_header_rows(row0: list, row1: list) -> list:
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


def _table_row_to_text(headers, row) -> str:
    if isinstance(row, dict):
        parts = [f"{h}：{row.get(h,'')}"
                 for h in headers if h and row.get(h, '')]
    else:
        parts = []
        for i, v in enumerate(row):
            if not v or not v.strip():
                continue
            h = headers[i] if i < len(headers) else f"字段{i+1}"
            if not h:
                continue
            if v.strip() == h.strip():
                continue
            parts.append(f"{h}：{v}")
    return '；'.join(parts) if parts else ''


def _last_text_title(units, fallback):
    for u in reversed(units):
        if u['type'] == 'text':
            return u['para'][:80]
    return fallback


def _parse_units(content: str, min_len: int = 30) -> list:
    raw = content.split('\n')
    units = []
    current_heading = '（无标题）'
    i = 0

    while i < len(raw):
        line = raw[i]
        stripped = line.strip()

        # <!-- TABLE:tables/table_001.png --> 占位符
        m_tbl = re.match(r'^<!--\s*TABLE:(.*?)\s*-->$', stripped)
        if m_tbl:
            tbl_ref = m_tbl.group(1).strip()
            start = i
            i += 1
            if i < len(raw) and re.match(r'^!\[.*\]\(.*\)', raw[i].strip()):
                i += 1
            title = _last_text_title(units, current_heading)
            units.append({'type': 'table', 'fmt': 'png_ref',
                          'title': title, 'heading': current_heading,
                          'lines': [line], 'start': start, 'end': i - 1,
                          'row_count': 5,
                          'tbl_ref': tbl_ref,
                          'headers': [], 'data_rows': []})
            continue

        # HTML 表格
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
            if cur == '####':
                break
            if cur.startswith('@@@') and cur.endswith('@@@'):
                i += 1
                continue
            if re.match(r'^tags:\s*@@@', cur):
                i += 1
                continue
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


def _extract_question(answer_text: str) -> str:
    pairs = re.findall(r'[\w（）/\-]+[：:]([\w\s（）、，,·.·]+?)(?:[；;]|$)', answer_text)
    for val in pairs:
        val = val.strip()
        if val and not re.fullmatch(r'[\d.,]+', val) and len(val) <= 12:
            return f"{val}的相关信息是？"
    m = re.match(r'^([^\s，,是的有在为与]{2,6})[是的有在为与]', answer_text)
    if m:
        return f"{m.group(1)}的相关信息是？"
    return "该记录的相关信息是？"


def _auto_tags_from_qa(question: str, answer: str) -> list:
    tags = []
    subject = question.replace('的相关信息是？', '').replace('是什么？', '').strip()
    if subject:
        tags.append(subject)

    if '：' in answer or ':' in answer:
        pairs = re.findall(r'([^：:；;，,\n]+)[：:]\s*([^；;：:，,\n]+)', answer)
        added = 0
        for field, val in pairs:
            field = field.strip()
            val   = val.strip()
            if not val or val == subject:
                continue
            if re.fullmatch(r'\d+', val):
                continue
            unit_in_field = re.search(r'[（(]([^）)]+)[）)]', field)
            if unit_in_field and re.fullmatch(r'[\d.]+', val):
                val = val + unit_in_field.group(1)
            if (re.search(r'[\d.]+(?:km2?²?|m2?²?|‰|%|元|天|年|亩|万)', val) or
                    re.search(r'县|市|省|乡|镇|村', val)):
                if len(val) <= 20 and val not in tags:
                    tags.append(val)
                    added += 1
                    if added >= 3:
                        break
    else:
        nums = re.findall(r'\d+(?:\.\d+)?(?:km²?|‰|元|天|km|m|亩|万)', answer)
        for n in nums[:2]:
            if n not in tags:
                tags.append(n)

    if question not in tags:
        tags.append(question)

    return tags[:5]


def _rebuild(orig_lines: list, units: list, results: list) -> list:
    line_to_unit = {}
    for idx, u in enumerate(units):
        for ln in range(u['start'], u['end'] + 1):
            line_to_unit[ln] = idx

    out = []
    done_units: set = set()
    i = 0

    def add_chunk(content_lines, tags):
        out.append('####')
        for cl in content_lines:
            out.append(cl)
        if tags:
            out.append('tags: @@@' + '@@@'.join(tags) + '@@@')

    while i < len(orig_lines):
        uid = line_to_unit.get(i)
        if uid is not None and uid not in done_units:
            done_units.add(uid)
            unit   = units[uid]
            result = results[uid]
            fmt    = unit.get('fmt', '')

            if unit['type'] == 'text':
                tags = result.get('tags', []) if result else []
                add_chunk(unit['lines'], tags)
                i = unit['end'] + 1

            else:  # table
                if fmt == 'png_ref':
                    tbl_ref = unit.get('tbl_ref', '')
                    ref_lines = (
                        [f"<!-- TABLE:{tbl_ref} -->", f"![表格]({tbl_ref})"]
                        if tbl_ref else unit['lines']
                    )
                    add_chunk(ref_lines, [])

                    row_results = result.get('row_results', []) if result else []
                    for rr in row_results:
                        question = rr.get('question', '').strip()
                        answer   = (rr.get('answer', '') or rr.get('description', '') or
                                    rr.get('qa', '')).strip()
                        raw_tags = rr.get('tags', [])
                        tags = [t for t in raw_tags
                                if isinstance(t, str) and 0 < len(t) <= 25]
                        if not answer:
                            continue
                        if not question:
                            question = _extract_question(answer)
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
                            tags = [t for t in raw_tags
                                    if isinstance(t, str) and 0 < len(t) <= 25]
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
                            headers, data_rows = (
                                _parse_pipe_table(unit['lines'])
                                if fmt == 'pipe' else ([], [])
                            )
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
            if (s == '####'
                    or (s.startswith('@@@') and s.endswith('@@@'))
                    or re.match(r'^tags:\s*@@@', s)):
                i += 1
                continue
            out.append(line)
            i += 1

    while out and out[-1].strip() == '':
        out.pop()
    if out and out[-1] != '####':
        out.append('####')

    return out


# ══════════════════════════════════════════════════════════════
#  LLM 调用（独立实现，不依赖 GUI）
# ══════════════════════════════════════════════════════════════

def _call_llm(url: str, key: str, model: str,
              system_prompt: str, user_msg: str) -> str:
    import requests as rq
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {key}',
    }
    payload = {
        'model': model,
        'messages': [
            {'role': 'system', 'content': system_prompt},
            {'role': 'user',   'content': user_msg},
        ],
        'max_tokens': 2048,
        'temperature': 0,
    }
    max_retries = 4
    wait = 3
    last_exc = None
    for attempt in range(max_retries):
        _rate_limiter.acquire()
        try:
            r = rq.post(url, headers=headers, json=payload, timeout=300)
            if r.status_code == 429:
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


def _process_text_unit(unit: dict, url: str, key: str, model: str) -> dict:
    sys_p = (PROMPT_TEXT
             .replace('__HEADING__',   unit['heading'])
             .replace('__PARAGRAPH__', unit['para']))
    raw = _call_llm(url, key, model, sys_p, '只输出JSON对象{"tags":[...]}')
    try:
        tags = json.loads(raw).get('tags', [])
    except Exception:
        tags = ['解析失败']
    return {'type': 'text', 'tags': tags[:6]}


def _process_table_unit(unit: dict, url: str, key: str, model: str,
                        md_path: str = '') -> dict:
    fmt       = unit.get('fmt', '')
    headers   = unit.get('headers', [])
    data_rows = unit.get('data_rows', [])

    # png_ref 类型：从旁边的 .json 文件加载 grid 数据
    if fmt == 'png_ref' and not data_rows:
        tbl_ref = unit.get('tbl_ref', '')
        if tbl_ref and md_path:
            json_path = Path(md_path).parent / tbl_ref.replace('.png', '.json')
            if json_path.exists():
                try:
                    grid = json.loads(json_path.read_text(encoding='utf-8'))
                    if grid and len(grid) > 1:
                        raw_headers = grid[0]
                        if len(grid) > 2 and _is_sub_header(grid[1]):
                            headers   = grid[1]
                            data_rows = grid[2:]
                        elif (len(grid) > 2
                              and not _is_sub_header(grid[1])
                              and not any(
                                  re.fullmatch(r'[\d.,]+', v)
                                  for v in grid[1] if str(v).strip())):
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
        raw = _call_llm(url, key, model, sys_p, '只输出JSON对象{"tags":[...]}')
        try:
            tags = json.loads(raw).get('tags', [])
        except Exception:
            tags = [unit['title']]
        return {'type': 'table', 'tags': tags, 'row_results': []}

    # 整表一次 LLM 请求
    if fmt == 'png_ref':
        table_text = '\n'.join(
            '\t'.join(str(c) for c in r)
            for r in [headers] + list(data_rows)
        )
    else:
        table_text = '\n'.join(unit['lines'])

    sys_p = (PROMPT_TABLE_BATCH
             .replace('__TITLE__', unit['title'])
             .replace('__TABLE__', table_text))
    row_count = len(data_rows)
    try:
        raw = _call_llm(
            url, key, model, sys_p,
            f'只输出JSON数组，共{row_count}个对象，每行数据对应一个，完整保留所有字段。'
        )
        batch = json.loads(raw)
        if not isinstance(batch, list):
            raise ValueError('not a list')
    except Exception:
        batch = [
            {'question': '', 'answer': _table_row_to_text(headers, row),
             'tags': [unit['title']]}
            for row in data_rows
        ]

    # 补齐行数
    if len(batch) < row_count:
        batch += [
            {'question': '', 'answer': _table_row_to_text(headers, data_rows[i]),
             'tags': [unit['title']]}
            for i in range(len(batch), row_count)
        ]

    return {'type': 'table', 'row_results': batch[:row_count]}


# ══════════════════════════════════════════════════════════════
#  公共入口
# ══════════════════════════════════════════════════════════════

def annotate_md(
    md_path: str,
    llm_url: str,
    llm_key: str,
    llm_model: str,
    concurrency: int = 3,
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
    cancel_event: Optional[threading.Event] = None,
) -> str:
    """
    对 Markdown 文件进行数据标注，输出 Dify 知识库格式。

    Args:
        md_path:      输入 .md 文件路径
        llm_url:      LLM 补全接口地址（OpenAI 兼容格式）
        llm_key:      API Key
        llm_model:    模型名称
        concurrency:  并发处理数（1–10）
        progress_cb:  进度回调 (current, total, message)
        cancel_event: 取消信号，调用 .set() 可中断

    Returns:
        输出文件路径（与输入同目录，文件名加 _annotated 后缀）
    """
    content = Path(md_path).read_text(encoding='utf-8')
    orig_lines = content.split('\n')
    units = _parse_units(content, min_len=30)
    total = len(units)
    results: list = [None] * total
    done = [0]

    def _process(idx: int):
        if cancel_event and cancel_event.is_set():
            return idx, None
        unit = units[idx]
        try:
            if unit['type'] == 'text':
                return idx, _process_text_unit(unit, llm_url, llm_key, llm_model)
            else:
                return idx, _process_table_unit(
                    unit, llm_url, llm_key, llm_model, md_path)
        except Exception:
            return idx, None

    with ThreadPoolExecutor(max_workers=max(1, concurrency)) as pool:
        futures = {pool.submit(_process, i): i for i in range(total)}
        for fut in as_completed(futures):
            if cancel_event and cancel_event.is_set():
                break
            idx, res = fut.result()
            results[idx] = res
            done[0] += 1
            if progress_cb:
                progress_cb(done[0], total, f'标注进度 {done[0]}/{total}')

    out_lines = _rebuild(orig_lines, units, results)
    base = Path(md_path)
    out_path = base.parent / (base.stem + '_annotated.md')
    out_path.write_text('\n'.join(out_lines), encoding='utf-8')
    return str(out_path)
