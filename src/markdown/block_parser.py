"""块级 Markdown 解析器"""

import re
import uuid
from src.markdown.ir import MarkdownBlock, BlockType, InlineType
from src.markdown.inline_parser import parse_inline

# 块级正则
RE_HEADING = re.compile(r'^(#{1,6})\s+(.+?)(?:\s+#+)?$')
RE_FENCED_CODE = re.compile(r'^(`{3,}|~{3,})(\w*)\s*$')
RE_BLOCKQUOTE = re.compile(r'^((?:>\s*)+)(.*)')
RE_HR = re.compile(r'^(?:---+|\*\*\*+|___+)\s*$')
RE_UNORDERED = re.compile(r'^(\s*)([-*+])\s+(.*)')
RE_ORDERED = re.compile(r'^(\s*)(\d+[.)])\s+(.*)')
# 兼容中文/中式有序列表标记：1、 / (1) / （1）
RE_ORDERED_CN_DUN = re.compile(r'^(\s*)(\d+、)\s*(.*)')
RE_ORDERED_CN_PAREN = re.compile(r'^(\s*)([（(]\d+[)）])\s*(.*)')
RE_TASK_LIST = re.compile(r'^(\s*)[-*+]\s+\[([ xX])\]\s+(.*)')
RE_LATEX_BLOCK = re.compile(r'^\$\$\s*$')
RE_LATEX_BLOCK_INLINE = re.compile(r'^\$\$\s*(.+?)\s*\$\$$')
RE_TABLE_ROW = re.compile(r'^\s*\|?.+\|.+\|?\s*$')
RE_TABLE_SEP = re.compile(r'^\|[\s:]*-{3,}[\s:]*')
RE_SOFT_BREAK = re.compile(r'  $')  # 行尾两个空格 = 硬换行
RE_FOOTNOTE_DEF = re.compile(r'^\[\^([A-Za-z0-9_-]+)\]:\s+(.*)')


# ── 表格辅助 ──

def _parse_table_row(text: str) -> list[str]:
    """解析一行表格为单元格列表（支持转义竖线 \\|）"""
    s = text.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|') and not s.endswith('\\|'):
        s = s[:-1]
    # 用唯一占位符保护转义竖线
    placeholder = f"__PIPE_{uuid.uuid4().hex[:8]}__"
    s = s.replace('\\|', placeholder)
    cells = [c.strip().replace(placeholder, '|') for c in s.split('|')]
    return cells


def _parse_alignment(sep_text: str) -> list[str]:
    """从分隔行解析对齐方式"""
    cells = sep_text.strip().strip('|').split('|')
    result = []
    for c in cells:
        c = c.strip()
        if c.startswith(':') and c.endswith(':'):
            result.append('center')
        elif c.endswith(':'):
            result.append('right')
        else:
            result.append('left')
    return result


def _is_table_sep(text: str) -> bool:
    """判断是否为表格分隔行 |---|---|"""
    s = text.strip()
    if '|' not in s:
        return False
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|') and not s.endswith('\\|'):
        s = s[:-1]
    cells = s.split('|')
    if len(cells) < 2:
        return False
    return all(re.match(r'\s*:?-{3,}:?\s*$', c) for c in cells)


def _table_col_count(text: str) -> int:
    """Return parsed table column count for a row-like line."""
    return len(_parse_table_row(text))


def _has_inline_markdown(text: str) -> bool:
    """判断文本是否含可渲染的行内 Markdown 语法。"""
    spans = parse_inline(text)
    return any(span.type != InlineType.TEXT for span in spans)


def _has_block_markdown(para_texts: list[tuple[int, str]]) -> bool:
    """预扫描：判断段落集合中是否存在块级 Markdown 语法。

    用于区分"Markdown 粘贴"和"纯文本论文"：
    有块级语法 → 启用段落合并；无 → 跳过合并。
    """
    for i, (_, text) in enumerate(para_texts):
        s = text.strip()
        if not s:
            continue
        if (
                RE_HEADING.match(s) or RE_FENCED_CODE.match(s)
                or RE_BLOCKQUOTE.match(s) or RE_HR.match(s)
                or RE_LATEX_BLOCK.match(s) or RE_LATEX_BLOCK_INLINE.match(s)
        ):
            return True
        # 表格要满足“表头行 + 分隔行”才视为块级 Markdown，避免普通“含竖线文本”误判。
        if RE_TABLE_ROW.match(s) and i + 1 < len(para_texts):
            next_s = (para_texts[i + 1][1] or "").strip()
            if next_s and _is_table_sep(next_s):
                if _table_col_count(s) == _table_col_count(next_s):
                    return True
    return False


def _match_ordered_item(text: str, *, allow_cn: bool = False):
    """匹配有序列表项（支持可选中文标记）。"""
    m = RE_ORDERED.match(text)
    if m:
        return m
    if allow_cn:
        m = RE_ORDERED_CN_DUN.match(text)
        if m:
            return m
        m = RE_ORDERED_CN_PAREN.match(text)
        if m:
            return m
    return None


def _has_unambiguous_markdown_list_cluster(
        para_texts: list[tuple[int, str]]) -> bool:
    """是否存在连续的“强 Markdown 特征”列表行。

    仅统计标准 Markdown 列表标记（-/*/+、1./1)、任务清单），
    用于判断 docx 粘贴文本是否应开启列表转换。
    """
    consecutive = 0
    for _, text in para_texts:
        if not (text or "").strip():
            consecutive = 0
            continue
        is_strong_list = bool(
            RE_TASK_LIST.match(text)
            or RE_UNORDERED.match(text)
            or RE_ORDERED.match(text)
        )
        if is_strong_list:
            consecutive += 1
            if consecutive >= 2:
                return True
        else:
            consecutive = 0
    return False


def _indent_to_level(indent: int) -> int:
    """将缩进空格数转换为列表嵌套层级 (0-based)"""
    return min(indent // 2, 8)  # 每 2 空格一级，最多 8 级


def _is_block_start(line: str, *, allow_cn_ordered: bool = False) -> bool:
    """判断一行是否为块级元素的起始行"""
    s = line.strip()
    if not s:
        return True
    if RE_HEADING.match(s):
        return True
    if RE_FENCED_CODE.match(s):
        return True
    if RE_BLOCKQUOTE.match(s):
        return True
    if RE_HR.match(s):
        return True
    if RE_TASK_LIST.match(s):
        return True
    if RE_UNORDERED.match(s) or _match_ordered_item(s, allow_cn=allow_cn_ordered):
        return True
    if RE_TABLE_ROW.match(s):
        return True
    if RE_LATEX_BLOCK.match(s) or RE_LATEX_BLOCK_INLINE.match(s):
        return True
    if RE_FOOTNOTE_DEF.match(s):
        return True
    return False


def _merge_paragraph_lines(lines: list[str]) -> list['InlineSpan']:
    """合并多行文本为一个段落的 InlineSpan 列表。

    规则：
    - 行尾两个空格 → 硬换行 (LINE_BREAK)
    - 否则行间插入空格连接
    """
    from src.markdown.ir import InlineSpan, InlineType

    all_spans: list[InlineSpan] = []
    prev_hard_break = False
    for idx, line in enumerate(lines):
        # 检查行尾是否有两个空格（硬换行）
        has_hard_break = bool(RE_SOFT_BREAK.search(line))
        clean = line.rstrip()

        spans = parse_inline(clean)
        if idx > 0 and all_spans:
            # 默认软换行按空格拼接；仅显式双空格才输出硬换行。
            if prev_hard_break:
                all_spans.append(InlineSpan(InlineType.LINE_BREAK, "\n"))
            else:
                all_spans.append(InlineSpan(InlineType.TEXT, " "))

        all_spans.extend(spans)
        prev_hard_break = has_hard_break

    return all_spans


# ── 独立转换模式：解析原始 .md 文本 ──

def parse_markdown_text(text: str) -> list[MarkdownBlock]:
    """解析原始 Markdown 文本为 MarkdownBlock 列表"""
    lines = text.split('\n')
    blocks: list[MarkdownBlock] = []
    i = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.rstrip()

        # 空行
        if not stripped:
            blocks.append(MarkdownBlock(type=BlockType.BLANK))
            i += 1
            continue

        # 围栏代码块
        m = RE_FENCED_CODE.match(stripped)
        if m:
            fence = m.group(1)
            lang = m.group(2) or ""
            code_lines = []
            j = i + 1
            closed = False
            while j < len(lines):
                if lines[j].rstrip().startswith(fence[0] * len(fence)):
                    closed = True
                    break
                code_lines.append(lines[j])
                j += 1
            if closed:
                blocks.append(MarkdownBlock(
                    type=BlockType.CODE_BLOCK, language=lang,
                    code_lines=code_lines,
                    raw_text='\n'.join(code_lines)))
                i = j + 1
                continue

        # LaTeX 块（单行或围栏）
        m_inline = RE_LATEX_BLOCK_INLINE.match(stripped)
        if m_inline:
            blocks.append(MarkdownBlock(
                type=BlockType.LATEX_BLOCK,
                raw_text=m_inline.group(1).strip()))
            i += 1
            continue

        # LaTeX 围栏块 $$
        if RE_LATEX_BLOCK.match(stripped):
            latex_lines = []
            j = i + 1
            while j < len(lines) and not RE_LATEX_BLOCK.match(lines[j].rstrip()):
                latex_lines.append(lines[j])
                j += 1
            blocks.append(MarkdownBlock(
                type=BlockType.LATEX_BLOCK,
                raw_text='\n'.join(latex_lines)))
            i = j + 1
            continue

        # 表格
        if RE_TABLE_ROW.match(stripped) and i + 1 < len(lines):
            if _is_table_sep(lines[i + 1].rstrip()):
                headers = _parse_table_row(stripped)
                aligns = _parse_alignment(lines[i + 1].rstrip())
                if len(headers) < 2 or len(aligns) != len(headers):
                    i += 1
                    continue
                rows = []
                j = i + 2
                while j < len(lines):
                    row_line = lines[j].rstrip()
                    if not RE_TABLE_ROW.match(row_line):
                        break
                    row_cells = _parse_table_row(row_line)
                    if len(row_cells) != len(headers):
                        break
                    rows.append(row_cells)
                    j += 1
                blocks.append(MarkdownBlock(
                    type=BlockType.TABLE,
                    table_headers=headers,
                    table_alignments=aligns,
                    table_rows=rows))
                i = j
                continue

        # 标题
        m = RE_HEADING.match(stripped)
        if m:
            level = len(m.group(1))
            content = m.group(2)
            blocks.append(MarkdownBlock(
                type=BlockType.HEADING, level=level,
                spans=parse_inline(content),
                raw_text=content))
            i += 1
            continue

        # 水平线
        if RE_HR.match(stripped):
            blocks.append(MarkdownBlock(type=BlockType.HORIZONTAL_RULE))
            i += 1
            continue

        # 引用块（支持嵌套 > > ）
        m = RE_BLOCKQUOTE.match(stripped)
        if m:
            quote_level = m.group(1).count('>')
            bq_lines = [m.group(2)]
            max_level = quote_level
            j = i + 1
            while j < len(lines):
                bm = RE_BLOCKQUOTE.match(lines[j].rstrip())
                if bm:
                    lvl = bm.group(1).count('>')
                    max_level = max(max_level, lvl)
                    bq_lines.append(bm.group(2))
                    j += 1
                else:
                    break
            content = '\n'.join(bq_lines)
            blocks.append(MarkdownBlock(
                type=BlockType.BLOCKQUOTE,
                spans=parse_inline(content),
                raw_text=content,
                quote_level=max_level))
            i = j
            continue

        # 脚注定义 [^id]: content
        fm = RE_FOOTNOTE_DEF.match(stripped)
        if fm:
            blocks.append(MarkdownBlock(
                type=BlockType.FOOTNOTE_DEF,
                raw_text=fm.group(2),
                spans=parse_inline(fm.group(2)),
                list_marker=fm.group(1)))  # 借用 list_marker 存 fn_id
            i += 1
            continue

        # 任务清单 - [x] / - [ ]
        tm = RE_TASK_LIST.match(stripped)
        if tm:
            indent = len(tm.group(1))
            checked = tm.group(2).lower() == 'x'
            content = tm.group(3)
            blocks.append(MarkdownBlock(
                type=BlockType.TASK_LIST_ITEM,
                list_indent=indent, checked=checked,
                list_level=_indent_to_level(indent),
                spans=parse_inline(content),
                raw_text=content))
            i += 1
            continue

        # 列表项（支持中文/中式有序标记）
        m = RE_UNORDERED.match(stripped) or _match_ordered_item(
            stripped, allow_cn=True)
        if m:
            indent = len(m.group(1))
            marker = m.group(2)
            content = m.group(3)
            blocks.append(MarkdownBlock(
                type=BlockType.LIST_ITEM,
                list_marker=marker, list_indent=indent,
                list_level=_indent_to_level(indent),
                spans=parse_inline(content),
                raw_text=content))
            i += 1
            continue

        # 普通段落（合并连续非空行）
        # 保留原始行尾空格，供“两个空格=硬换行”判断使用。
        para_lines = [line]
        j = i + 1
        while j < len(lines):
            next_line_raw = lines[j]
            next_line = next_line_raw.rstrip()
            if not next_line:
                break
            # 如果下一行是块级元素，停止合并
            if _is_block_start(next_line, allow_cn_ordered=True):
                break
            para_lines.append(next_line_raw)
            j += 1
        # 合并行：行尾两空格 → LINE_BREAK，否则空格连接
        merged_spans = _merge_paragraph_lines(para_lines)
        raw = '\n'.join(para_lines)
        blocks.append(MarkdownBlock(
            type=BlockType.PARAGRAPH,
            spans=merged_spans,
            raw_text=raw))
        i = j

    return blocks


# ── 粘贴修复模式：解析 DOCX 段落文本 ──

def parse_docx_paragraphs(
        para_texts: list[tuple[int, str]],
        *,
        markdown_paste_hint: bool = False) -> list[MarkdownBlock]:
    """解析 DOCX 段落文本为 MarkdownBlock 列表。

    输入: [(段落索引, 段落文本), ...]
    输出: MarkdownBlock 列表，source_para_indices 已填充
    """
    blocks: list[MarkdownBlock] = []
    i = 0

    # 预扫描：有块级 Markdown 语法时才启用段落合并
    # （区分"Markdown 粘贴"和"纯文本论文"）
    merge_enabled = _has_block_markdown(para_texts)
    # 列表转换仅在“疑似 Markdown 粘贴”语境下启用，减少误伤原生文档结构。
    list_mode_enabled = bool(markdown_paste_hint) or merge_enabled or _has_unambiguous_markdown_list_cluster(para_texts)

    while i < len(para_texts):
        para_idx, text = para_texts[i]
        stripped = text.strip()

        # 空段落只在疑似 Markdown 粘贴语境下发射 BLANK，避免误删普通正文留白。
        if not stripped:
            if markdown_paste_hint or merge_enabled or list_mode_enabled:
                blocks.append(MarkdownBlock(
                    type=BlockType.BLANK,
                    source_para_indices=[para_idx]))
            i += 1
            continue

        # 围栏代码块
        m = RE_FENCED_CODE.match(stripped)
        if m:
            fence = m.group(1)
            lang = m.group(2) or ""
            code_lines = []
            indices = [para_idx]
            j = i + 1
            closed = False
            while j < len(para_texts):
                j_idx, j_text = para_texts[j]
                indices.append(j_idx)
                if j_text.strip().startswith(fence[0] * len(fence)):
                    closed = True
                    break
                code_lines.append(j_text)
                j += 1
            if closed:
                blocks.append(MarkdownBlock(
                    type=BlockType.CODE_BLOCK, language=lang,
                    code_lines=code_lines,
                    raw_text='\n'.join(code_lines),
                    source_para_indices=indices))
                i = j + 1
                continue

        # LaTeX 块（单段或围栏）
        m_inline = RE_LATEX_BLOCK_INLINE.match(stripped)
        if m_inline:
            blocks.append(MarkdownBlock(
                type=BlockType.LATEX_BLOCK,
                raw_text=m_inline.group(1).strip(),
                source_para_indices=[para_idx]))
            i += 1
            continue

        if RE_LATEX_BLOCK.match(stripped):
            latex_lines = []
            indices = [para_idx]
            j = i + 1
            closed = False
            while j < len(para_texts):
                j_idx, j_text = para_texts[j]
                indices.append(j_idx)
                if RE_LATEX_BLOCK.match(j_text.strip()):
                    closed = True
                    break
                latex_lines.append(j_text)
                j += 1
            if closed:
                blocks.append(MarkdownBlock(
                    type=BlockType.LATEX_BLOCK,
                    raw_text='\n'.join(latex_lines).strip(),
                    source_para_indices=indices))
                i = j + 1
                continue

        # 表格（多段落）
        if RE_TABLE_ROW.match(stripped):
            table_items = [(para_idx, stripped)]
            header_cols = _table_col_count(stripped)
            j = i + 1
            while j < len(para_texts):
                j_idx, j_text = para_texts[j]
                row_text = j_text.strip()
                if not row_text:
                    break
                if not RE_TABLE_ROW.match(row_text):
                    break
                if _table_col_count(row_text) != header_cols:
                    break
                table_items.append((j_idx, row_text))
                j += 1
            texts = [t for _, t in table_items]
            if len(texts) >= 2 and _is_table_sep(texts[1]):
                headers = _parse_table_row(texts[0])
                aligns = _parse_alignment(texts[1])
                if len(headers) >= 2 and len(aligns) == len(headers):
                    rows = [_parse_table_row(t) for t in texts[2:]]
                    indices = [idx for idx, _ in table_items]
                    blocks.append(MarkdownBlock(
                        type=BlockType.TABLE,
                        table_headers=headers,
                        table_alignments=aligns,
                        table_rows=rows,
                        source_para_indices=indices))
                    i = j
                    continue

        # 标题
        m = RE_HEADING.match(stripped)
        if m:
            level = len(m.group(1))
            content = m.group(2)
            blocks.append(MarkdownBlock(
                type=BlockType.HEADING, level=level,
                spans=parse_inline(content),
                raw_text=content,
                source_para_indices=[para_idx]))
            i += 1
            continue

        # 水平线
        if RE_HR.match(stripped):
            blocks.append(MarkdownBlock(
                type=BlockType.HORIZONTAL_RULE,
                source_para_indices=[para_idx]))
            i += 1
            continue

        # 引用块（多段落，支持嵌套）
        m = RE_BLOCKQUOTE.match(stripped)
        if m:
            quote_level = m.group(1).count('>')
            bq_lines = [m.group(2)]
            max_level = quote_level
            indices = [para_idx]
            j = i + 1
            while j < len(para_texts):
                j_idx, j_text = para_texts[j]
                bm = RE_BLOCKQUOTE.match(j_text.strip())
                if bm:
                    lvl = bm.group(1).count('>')
                    max_level = max(max_level, lvl)
                    bq_lines.append(bm.group(2))
                    indices.append(j_idx)
                    j += 1
                else:
                    break
            content = '\n'.join(bq_lines)
            blocks.append(MarkdownBlock(
                type=BlockType.BLOCKQUOTE,
                spans=parse_inline(content),
                raw_text=content,
                quote_level=max_level,
                source_para_indices=indices))
            i = j
            continue

        # 脚注定义 [^id]: content
        fm = RE_FOOTNOTE_DEF.match(stripped)
        if fm:
            blocks.append(MarkdownBlock(
                type=BlockType.FOOTNOTE_DEF,
                raw_text=fm.group(2),
                spans=parse_inline(fm.group(2)),
                list_marker=fm.group(1),
                source_para_indices=[para_idx]))
            i += 1
            continue

        # 任务清单 - [x] / - [ ]（用 text 保留前导空格以检测缩进）
        tm = RE_TASK_LIST.match(text)
        if tm and list_mode_enabled:
            indent = len(tm.group(1))
            checked = tm.group(2).lower() == 'x'
            content = tm.group(3)
            blocks.append(MarkdownBlock(
                type=BlockType.TASK_LIST_ITEM,
                list_indent=indent, checked=checked,
                list_level=_indent_to_level(indent),
                spans=parse_inline(content),
                raw_text=content,
                source_para_indices=[para_idx]))
            i += 1
            continue

        # 列表项（仅在疑似 Markdown 语境启用，避免误改原生论文编号结构）
        m = None
        if list_mode_enabled:
            m = RE_UNORDERED.match(text) or _match_ordered_item(
                text, allow_cn=list_mode_enabled)
        if m is not None:
            indent = len(m.group(1))
            marker = m.group(2)
            content = m.group(3)
            blocks.append(MarkdownBlock(
                type=BlockType.LIST_ITEM,
                list_marker=marker, list_indent=indent,
                list_level=_indent_to_level(indent),
                spans=parse_inline(content),
                raw_text=content,
                source_para_indices=[para_idx]))
            i += 1
            continue

        # 无行内 Markdown → 跳过（避免误改普通正文段落）
        if not _has_inline_markdown(stripped):
            i += 1
            continue

        # 合并模式：向前合并连续的非空非块级段落
        # 非合并模式：仅处理当前段落的行内 Markdown
        para_lines = [stripped]
        merge_indices = [para_idx]
        if merge_enabled:
            j = i + 1
            while j < len(para_texts):
                j_idx, j_text = para_texts[j]
                j_stripped = j_text.strip()
                if not j_stripped:
                    break
                if _is_block_start(
                        j_stripped, allow_cn_ordered=list_mode_enabled):
                    break
                if not _has_inline_markdown(j_stripped):
                    break
                para_lines.append(j_stripped)
                merge_indices.append(j_idx)
                j += 1
        else:
            j = i + 1

        # 有行内 Markdown 或多段合并时才生成块
        if len(merge_indices) > 1 or _has_inline_markdown(stripped):
            merged_spans = _merge_paragraph_lines(para_lines)
            raw = '\n'.join(para_lines)
            blocks.append(MarkdownBlock(
                type=BlockType.PARAGRAPH,
                spans=merged_spans,
                raw_text=raw,
                source_para_indices=merge_indices))
        i = j

    return blocks
