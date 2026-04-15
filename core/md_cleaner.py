"""
Markdown 清洗器
去除 VLM 解析时误识别进来的广告/水印内容：
  - 纯 URL 行（如大量重复的 jq.qq.com 链接）
  - 以 # 开头但内容全为 URL 的假标题行
  - 重复3次以上的相同短词/短句行（水印特征）
  - 连续超过2个空行压缩为1个
"""
from __future__ import annotations

import re
from typing import List

# 匹配裸 URL（不在 Markdown 图片/链接语法内）
_URL_RE = re.compile(r'https?://\S+', re.IGNORECASE)

# Markdown 图片/链接语法（保留）
_MD_LINK_RE = re.compile(r'!?\[.*?\]\(.*?\)')

# 保留行：以这些开头的行不做清洗
_KEEP_PREFIXES = (
    '![',          # Markdown 图片
    '<!-- ',       # 注释（bbox 信息）
    '<!-- TABLE:', # 表格占位符
    '|',           # 表格行
)


def _strip_urls(text: str) -> str:
    """移除字符串中所有裸 URL（保留 Markdown 图片/链接语法）"""
    # 先把 Markdown 语法占位，避免误删
    placeholders = {}
    counter = [0]

    def _protect(m):
        key = f'\x00MDLINK{counter[0]}\x00'
        placeholders[key] = m.group(0)
        counter[0] += 1
        return key

    protected = _MD_LINK_RE.sub(_protect, text)
    cleaned = _URL_RE.sub('', protected)
    for k, v in placeholders.items():
        cleaned = cleaned.replace(k, v)
    return cleaned


def _is_url_only(line: str) -> bool:
    """判断一行是否仅由 URL 和空白构成（不含 Markdown 图片/链接）"""
    stripped = line.strip()
    if not stripped:
        return False
    # 有 Markdown 图片语法则保留
    if _MD_LINK_RE.search(stripped):
        return False
    remainder = _URL_RE.sub('', stripped).strip()
    return remainder == '' and bool(_URL_RE.search(stripped))


def _is_spam_heading(line: str) -> bool:
    """判断 # 标题行是否内容全为 URL（废广告标题）"""
    m = re.match(r'^(#{1,6})\s+(.*)', line.strip())
    if not m:
        return False
    content = m.group(2).strip()
    if not content:
        return False
    remainder = _URL_RE.sub('', content).strip()
    return remainder == ''


def _is_repeated_spam(line: str, threshold: int = 3) -> bool:
    """
    判断一行是否为重复水印：同一短词出现 threshold 次以上。
    例：'联系96群QQ1563033835 联系96群QQ1563033835 联系96群QQ1563033835'
    """
    stripped = line.strip()
    if not stripped or len(stripped) > 500:
        return False
    # 按空白分词
    tokens = stripped.split()
    if len(tokens) < threshold:
        return False
    # 取最长出现的 token
    from collections import Counter
    counts = Counter(tokens)
    most_common_token, freq = counts.most_common(1)[0]
    # 如果最高频 token 占比超过 80% 且出现次数 >= threshold，认为是水印
    return freq >= threshold and freq / len(tokens) >= 0.8


def clean_markdown(text: str) -> str:
    """
    清洗 Markdown 文本，移除广告/水印行。

    Args:
        text: 原始 Markdown 字符串

    Returns:
        清洗后的 Markdown 字符串
    """
    lines = text.split('\n')
    cleaned: List[str] = []

    for line in lines:
        # 保留 Markdown 图片/注释/表格行，不做任何处理
        stripped = line.strip()
        if any(stripped.startswith(p) for p in _KEEP_PREFIXES):
            cleaned.append(line)
            continue

        # 过滤纯 URL 行
        if _is_url_only(line):
            continue

        # 过滤内容全为 URL 的假标题
        if _is_spam_heading(line):
            continue

        # 过滤重复水印行
        if _is_repeated_spam(line):
            continue

        cleaned.append(line)

    # 压缩连续空行（超过 2 个合并为 1 个）
    result = re.sub(r'\n{3,}', '\n\n', '\n'.join(cleaned))
    return result.strip()
