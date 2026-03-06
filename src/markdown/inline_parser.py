"""行内 Markdown 格式解析器"""

import re
from src.markdown.ir import InlineSpan, InlineType

# <br> / <br/> / <br /> 标签正则
RE_BR_TAG = re.compile(r'<br\s*/?\s*>', re.IGNORECASE)

# 按优先级排列（最长/最具体的在前）
INLINE_PATTERNS = [
    # ***bold italic***
    (re.compile(r'\*\*\*(.+?)\*\*\*'), InlineType.BOLD_ITALIC),
    # **bold**
    (re.compile(r'\*\*(.+?)\*\*'), InlineType.BOLD),
    # ~~strikethrough~~
    (re.compile(r'~~(.+?)~~'), InlineType.STRIKETHROUGH),
    # *italic* (不匹配 ** 内部的 *)
    (re.compile(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)'), InlineType.ITALIC),
    # ``code with backticks`` (双反引号优先)
    (re.compile(r'``(.+?)``'), InlineType.CODE),
    # `inline code`
    (re.compile(r'`([^`]+)`'), InlineType.CODE),
    # ![alt](url) — 图片（必须在 HYPERLINK 之前）
    (re.compile(r'!\[([^\]]*)\]\(([^)]+)\)'), InlineType.IMAGE),
    # [text](url)
    (re.compile(r'\[([^\]]+)\]\(([^)]+)\)'), InlineType.HYPERLINK),
    # [^id] — 脚注引用
    (re.compile(r'\[\^([A-Za-z0-9_-]+)\]'), InlineType.FOOTNOTE_REF),
    # $latex$ (行内公式)
    (re.compile(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)'), InlineType.LATEX_INLINE),
]


def parse_inline(text: str) -> list[InlineSpan]:
    """将一行文本解析为 InlineSpan 列表（位置扫描算法）

    支持: **bold**, *italic*, ***bold_italic***, ~~strike~~,
          `code`, [link](url), $latex$, <br> 标签
    """
    spans: list[InlineSpan] = []
    pos = 0

    while pos < len(text):
        # 优先检测 <br> 标签
        br_match = RE_BR_TAG.match(text, pos)
        if br_match:
            spans.append(InlineSpan(InlineType.LINE_BREAK, "\n"))
            pos = br_match.end()
            continue

        # 在所有模式中找最早匹配
        earliest_match = None
        earliest_type = None
        earliest_br = RE_BR_TAG.search(text, pos)

        for pattern, inline_type in INLINE_PATTERNS:
            m = pattern.search(text, pos)
            if m and (earliest_match is None
                      or m.start() < earliest_match.start()):
                earliest_match = m
                earliest_type = inline_type

        # 比较 <br> 和其他模式的位置
        if earliest_br and (earliest_match is None
                            or earliest_br.start() < earliest_match.start()):
            # <br> 更早出现
            if earliest_br.start() > pos:
                spans.append(InlineSpan(
                    InlineType.TEXT, text[pos:earliest_br.start()]))
            spans.append(InlineSpan(InlineType.LINE_BREAK, "\n"))
            pos = earliest_br.end()
            continue

        if earliest_match is None:
            remaining = text[pos:]
            if remaining:
                spans.append(InlineSpan(InlineType.TEXT, remaining))
            break

        # 匹配前的纯文本
        if earliest_match.start() > pos:
            spans.append(InlineSpan(
                InlineType.TEXT, text[pos:earliest_match.start()]))

        # 格式化片段
        if earliest_type == InlineType.IMAGE:
            spans.append(InlineSpan(
                earliest_type, earliest_match.group(1),
                url=earliest_match.group(2)))
        elif earliest_type == InlineType.HYPERLINK:
            spans.append(InlineSpan(
                earliest_type, earliest_match.group(1),
                url=earliest_match.group(2)))
        elif earliest_type == InlineType.FOOTNOTE_REF:
            spans.append(InlineSpan(
                earliest_type, "",
                footnote_id=earliest_match.group(1)))
        else:
            content = next(
                (g for g in earliest_match.groups() if g is not None), "")
            spans.append(InlineSpan(earliest_type, content))

        pos = earliest_match.end()

    return spans
