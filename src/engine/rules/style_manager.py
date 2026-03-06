"""样式创建/更新规则"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from src.engine.rules.base import BaseRule
from src.engine.change_tracker import ChangeTracker
from src.scene.schema import SceneConfig, StyleConfig
from src.scene.heading_model import get_level_to_style_key, get_level_to_word_style
from src.utils.line_spacing import apply_line_spacing

ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# 不全局修改的样式（由 section_format / toc_format / page_setup 按段落直接应用）
SKIP_GLOBAL_STYLES = {
    "normal", "references_body", "acknowledgment_body",
    "abstract_body", "abstract_body_en",
    "abstract_title_cn", "abstract_title_en",
    "appendix_body", "resume_body", "symbol_table_body",
    "toc_title", "toc_chapter", "toc_level1", "toc_level2",
    "header_cn", "header_en", "page_number",
}


def _apply_style_config(style, sc: StyleConfig) -> None:
    """将 StyleConfig 应用到 Word 样式对象"""
    from lxml import etree
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    font = style.font
    font.name = sc.font_en
    font.size = Pt(sc.size_pt)
    font.bold = sc.bold if sc.bold else None
    font.italic = sc.italic if sc.italic else None

    # 中文字体 + 清除主题字体/颜色引用
    rpr = style.element.find(f'{W}rPr')
    if rpr is None:
        rpr = etree.SubElement(style.element, f'{W}rPr')

    # 强制黑色字体，清除文档主题色（如蓝色 accent1）
    color_el = rpr.find(f'{W}color')
    if color_el is not None:
        rpr.remove(color_el)
    color_el = etree.SubElement(rpr, f'{W}color')
    color_el.set(f'{W}val', '000000')
    rfonts = rpr.find(f'{W}rFonts')
    if rfonts is None:
        rfonts = etree.SubElement(rpr, f'{W}rFonts')
    # 设置显式字体
    rfonts.set(f'{W}eastAsia', sc.font_cn)
    rfonts.set(f'{W}ascii', sc.font_en)
    rfonts.set(f'{W}hAnsi', sc.font_en)
    # 移除主题字体引用（否则 theme 优先级高于显式字体名）
    for theme_attr in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
        rfonts.attrib.pop(f'{W}{theme_attr}', None)
    # 同步 szCs（font.size 只设置 sz，szCs 可能残留原值）
    sz_cs = rpr.find(f'{W}szCs')
    if sz_cs is not None:
        sz_cs.set(f'{W}val', str(int(sc.size_pt * 2)))
    else:
        sz_cs = etree.SubElement(rpr, f'{W}szCs')
        sz_cs.set(f'{W}val', str(int(sc.size_pt * 2)))

    # 段落格式
    pf = style.paragraph_format
    if sc.alignment in ALIGNMENT_MAP:
        pf.alignment = ALIGNMENT_MAP[sc.alignment]
    pf.space_before = Pt(sc.space_before_pt)
    pf.space_after = Pt(sc.space_after_pt)

    apply_line_spacing(pf, sc.line_spacing_type, sc.line_spacing_pt)

    if sc.first_line_indent_chars > 0:
        pf.first_line_indent = Pt(sc.size_pt * sc.first_line_indent_chars)


class StyleManagerRule(BaseRule):
    name = "style_manager"
    description = "创建/更新文档样式集"

    def apply(self, doc: Document, config: SceneConfig,
              tracker: ChangeTracker, context: dict) -> None:
        style_name_map = {}
        level_style_key_map = get_level_to_style_key(config)
        level_word_style_map = get_level_to_word_style(config)
        for level, style_key in level_style_key_map.items():
            word_style = level_word_style_map.get(level)
            if style_key and word_style:
                style_name_map[style_key] = word_style

        for style_key, sc in config.styles.items():
            # 跳过由 section_format 按段落直接应用的样式
            if style_key in SKIP_GLOBAL_STYLES:
                continue

            word_name = style_name_map.get(style_key, style_key)
            try:
                style = doc.styles[word_name]
            except KeyError:
                style = doc.styles.add_style(
                    word_name, WD_STYLE_TYPE.PARAGRAPH)

            _apply_style_config(style, sc)

            tracker.record(
                rule_name=self.name,
                target=f"Style '{word_name}'",
                section="global",
                change_type="style",
                before="(existing)",
                after=f"font={sc.font_cn}/{sc.font_en} {sc.size_pt}pt",
                paragraph_index=-1,
            )
