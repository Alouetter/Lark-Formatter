"""Markdown 中间表示数据结构"""

from dataclasses import dataclass, field
from enum import Enum


class BlockType(Enum):
    PARAGRAPH = "paragraph"
    HEADING = "heading"
    CODE_BLOCK = "code_block"
    TABLE = "table"
    BLOCKQUOTE = "blockquote"
    HORIZONTAL_RULE = "hr"
    LIST_ITEM = "list_item"
    TASK_LIST_ITEM = "task_list_item"
    BLANK = "blank"
    LATEX_BLOCK = "latex_block"
    FOOTNOTE_DEF = "footnote_def"


class InlineType(Enum):
    TEXT = "text"
    BOLD = "bold"
    ITALIC = "italic"
    BOLD_ITALIC = "bold_italic"
    STRIKETHROUGH = "strikethrough"
    CODE = "code"
    HYPERLINK = "hyperlink"
    LATEX_INLINE = "latex_inline"
    LINE_BREAK = "line_break"
    IMAGE = "image"
    FOOTNOTE_REF = "footnote_ref"


@dataclass
class InlineSpan:
    """行内格式化片段"""
    type: InlineType
    text: str
    url: str = ""       # for HYPERLINK / IMAGE
    footnote_id: str = ""  # for FOOTNOTE_REF


@dataclass
class MarkdownBlock:
    """块级 Markdown 元素"""
    type: BlockType
    level: int = 0
    spans: list[InlineSpan] = field(default_factory=list)
    raw_text: str = ""
    language: str = ""
    code_lines: list[str] = field(default_factory=list)
    table_headers: list[str] = field(default_factory=list)
    table_alignments: list[str] = field(default_factory=list)
    table_rows: list[list[str]] = field(default_factory=list)
    list_marker: str = ""
    list_indent: int = 0
    list_level: int = 0                  # nested list depth (0-based)
    checked: bool | None = None          # task list: True/False/None
    children: list['MarkdownBlock'] = field(default_factory=list)  # nested list items
    quote_level: int = 1                 # blockquote nesting depth
    source_para_indices: list[int] = field(default_factory=list)
