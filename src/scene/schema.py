"""场景配置 dataclass 定义"""

from dataclasses import dataclass, field

from src.utils.indent import sync_style_config_indent_fields


def heading_level_keys() -> list[str]:
    return [f"heading{i}" for i in range(1, 9)]


def heading_level_index(level_name: str) -> int:
    text = str(level_name or "").strip().lower()
    if text.startswith("heading"):
        try:
            value = int(text[7:])
        except ValueError:
            return 1
        if 1 <= value <= 8:
            return value
    return 1


def default_chain_id_for_level(level_name: str) -> str:
    level_idx = heading_level_index(level_name)
    if level_idx <= 1:
        return "current_only"
    parts = [f"l{i}" for i in range(1, level_idx)]
    return "_dot_".join(parts + ["current"])


@dataclass
class MarginConfig:
    top_cm: float = 3.8
    bottom_cm: float = 3.8
    left_cm: float = 3.2
    right_cm: float = 3.2


@dataclass
class PageSetupConfig:
    paper_size: str = "A4"
    margin: MarginConfig = field(default_factory=MarginConfig)
    gutter_cm: float = 0
    header_distance_cm: float = 3.0
    footer_distance_cm: float = 3.0


@dataclass
class NumberShellConfig:
    label: str = "{}"
    prefix: str = ""
    suffix: str = ""


@dataclass
class NumberCoreStyleConfig:
    label: str = ""
    sample: str = ""


@dataclass
class NumberChainSegmentConfig:
    type: str = "value"  # "value" | "literal"
    source: str = "current"
    text: str = ""


@dataclass
class NumberChainConfig:
    label: str = ""
    segments: list[NumberChainSegmentConfig] = field(default_factory=list)


@dataclass
class NumberPresetConfig:
    label: str = ""
    display_shell: str = "plain"
    display_core_style: str = "arabic"
    reference_core_style: str = ""
    chain: str = "current_only"


@dataclass
class HeadingLevelBindingConfig:
    enabled: bool = False
    display_shell: str = "plain"
    display_core_style: str = "arabic"
    reference_core_style: str = "arabic"
    chain: str = "current_only"
    title_separator: str = "\u3000"
    ooxml_separator_mode: str = "inline"  # "inline" | "suff" | "none"
    ooxml_suff: str | None = "nothing"    # "tab" | "space" | "nothing" | None
    ooxml_lvl_ind: dict[str, str] = field(default_factory=dict)
    start_at: int = 1
    restart_on: str | None = None
    include_in_toc: bool = True


REMOVED_NUMBER_CORE_STYLE_IDS = frozenset(
    {
        "arabic_fullwidth",
        "arabic_pad3",
        "roman_lower",
        "alpha_upper",
        "alpha_lower",
    }
)


def _default_shell_catalog() -> dict[str, NumberShellConfig]:
    return {
        "plain": NumberShellConfig(label="{}", prefix="", suffix=""),
        "dot_suffix": NumberShellConfig(label="{}.", prefix="", suffix="."),
        "chapter_cn": NumberShellConfig(label="第{}章", prefix="第", suffix="章"),
        "section_cn": NumberShellConfig(label="第{}节", prefix="第", suffix="节"),
        "dunhao_cn": NumberShellConfig(label="{}、", prefix="", suffix="、"),
        "paren_cn": NumberShellConfig(label="（{}）", prefix="（", suffix="）"),
        "paren_en": NumberShellConfig(label="({})", prefix="(", suffix=")"),
        "appendix_cn": NumberShellConfig(label="附录{}", prefix="附录", suffix=""),
    }


def _default_core_style_catalog() -> dict[str, NumberCoreStyleConfig]:
    return {
        "arabic": NumberCoreStyleConfig(label="1 2 3", sample="1"),
        "arabic_pad2": NumberCoreStyleConfig(label="01 02 03", sample="01"),
        "cn_lower": NumberCoreStyleConfig(label="一 二 三", sample="一"),
        "cn_upper": NumberCoreStyleConfig(label="壹 贰 叁", sample="壹"),
        "roman_upper": NumberCoreStyleConfig(label="I II III", sample="I"),
        "circled": NumberCoreStyleConfig(label="① ② ③", sample="①"),
        "circled_paren": NumberCoreStyleConfig(label="⑴ ⑵ ⑶", sample="⑴"),
    }


def _chain_segments_for_level(level_idx: int) -> list[NumberChainSegmentConfig]:
    if level_idx <= 1:
        return [NumberChainSegmentConfig(type="value", source="current")]
    segments: list[NumberChainSegmentConfig] = []
    for idx in range(1, level_idx):
        if segments:
            segments.append(NumberChainSegmentConfig(type="literal", text="."))
        segments.append(NumberChainSegmentConfig(type="value", source=f"level{idx}"))
    segments.append(NumberChainSegmentConfig(type="literal", text="."))
    segments.append(NumberChainSegmentConfig(type="value", source="current"))
    return segments


def _default_chain_catalog() -> dict[str, NumberChainConfig]:
    labels = {
        1: "一级",
        2: "二级",
        3: "三级",
        4: "四级",
        5: "五级",
        6: "六级",
        7: "七级",
    }
    catalog = {
        "current_only": NumberChainConfig(
            label="仅当前",
            segments=[NumberChainSegmentConfig(type="value", source="current")],
        )
    }
    for level_idx in range(2, 9):
        chain_id = default_chain_id_for_level(f"heading{level_idx}")
        names = [labels[idx] for idx in range(1, level_idx)]
        catalog[chain_id] = NumberChainConfig(
            label=".".join(names + ["当前"]),
            segments=_chain_segments_for_level(level_idx),
        )
    return catalog


def _default_preset_catalog() -> dict[str, NumberPresetConfig]:
    return {
        "chapter_cn": NumberPresetConfig(
            label="第一章",
            display_shell="chapter_cn",
            display_core_style="cn_lower",
            reference_core_style="arabic",
            chain="current_only",
        ),
        "section_cn_arabic": NumberPresetConfig(
            label="第1节",
            display_shell="section_cn",
            display_core_style="arabic",
            reference_core_style="arabic",
            chain="current_only",
        ),
        "ordinal_cn": NumberPresetConfig(
            label="一、",
            display_shell="dunhao_cn",
            display_core_style="cn_lower",
            reference_core_style="cn_lower",
            chain="current_only",
        ),
        "paren_cn": NumberPresetConfig(
            label="（一）",
            display_shell="paren_cn",
            display_core_style="cn_lower",
            reference_core_style="cn_lower",
            chain="current_only",
        ),
        "decimal_2": NumberPresetConfig(
            label="1.1",
            display_shell="plain",
            display_core_style="arabic",
            reference_core_style="arabic",
            chain=default_chain_id_for_level("heading2"),
        ),
        "decimal_3": NumberPresetConfig(
            label="1.1.1",
            display_shell="plain",
            display_core_style="arabic",
            reference_core_style="arabic",
            chain=default_chain_id_for_level("heading3"),
        ),
    }


def _default_level_bindings() -> dict[str, HeadingLevelBindingConfig]:
    bindings: dict[str, HeadingLevelBindingConfig] = {}
    for level_name in heading_level_keys():
        level_idx = heading_level_index(level_name)
        bindings[level_name] = HeadingLevelBindingConfig(
            enabled=False,
            display_shell="plain",
            display_core_style="arabic",
            reference_core_style="arabic",
            chain=default_chain_id_for_level(level_name),
            title_separator="\u3000",
            start_at=1,
            restart_on=None if level_idx <= 1 else f"heading{level_idx - 1}",
            include_in_toc=level_idx <= 3,
        )

    bindings["heading1"] = HeadingLevelBindingConfig(
        enabled=True,
        display_shell="chapter_cn",
        display_core_style="cn_lower",
        reference_core_style="arabic",
        chain="current_only",
        title_separator="\u3000",
        start_at=1,
        restart_on=None,
        include_in_toc=True,
    )
    bindings["heading2"] = HeadingLevelBindingConfig(
        enabled=True,
        display_shell="plain",
        display_core_style="arabic",
        reference_core_style="arabic",
        chain=default_chain_id_for_level("heading2"),
        title_separator="\u3000",
        start_at=1,
        restart_on="heading1",
        include_in_toc=True,
    )
    bindings["heading3"] = HeadingLevelBindingConfig(
        enabled=True,
        display_shell="paren_cn",
        display_core_style="cn_lower",
        reference_core_style="cn_lower",
        chain="current_only",
        title_separator="\u3000",
        start_at=1,
        restart_on="heading2",
        include_in_toc=True,
    )
    bindings["heading4"] = HeadingLevelBindingConfig(
        enabled=True,
        display_shell="plain",
        display_core_style="arabic",
        reference_core_style="arabic",
        chain=default_chain_id_for_level("heading4"),
        title_separator="\u3000",
        start_at=1,
        restart_on="heading3",
        include_in_toc=False,
    )
    return bindings


@dataclass
class HeadingNumberingV2Config:
    enabled: bool = True
    shell_catalog: dict[str, NumberShellConfig] = field(default_factory=_default_shell_catalog)
    core_style_catalog: dict[str, NumberCoreStyleConfig] = field(default_factory=_default_core_style_catalog)
    chain_catalog: dict[str, NumberChainConfig] = field(default_factory=_default_chain_catalog)
    preset_catalog: dict[str, NumberPresetConfig] = field(default_factory=_default_preset_catalog)
    level_bindings: dict[str, HeadingLevelBindingConfig] = field(default_factory=_default_level_bindings)


@dataclass
class HeadingLevelConfig:
    format: str = "arabic_dotted"
    template: str = "{current}"
    separator: str = "\u3000"
    custom_separator: str | None = None
    alignment: str = ""          # "" 使用样式默认, "left"/"center"/"right"/"justify"
    left_indent_chars: float = 0  # 左缩进（汉字数）

    @property
    def effective_separator(self) -> str:
        return self.custom_separator if self.custom_separator is not None else self.separator


@dataclass
class EnforcementConfig:
    ban_tab: bool = True
    ban_double_halfwidth_space: bool = True
    ban_mixed_separator_per_level: bool = True
    auto_fix: bool = True


@dataclass
class HeadingRiskGuardConfig:
    """标题识别防风险兜底参数（仅在分区异常时触发）。"""
    enabled: bool = True
    no_body_min_candidates: int = 2
    no_body_min_chapters: int = 1
    no_chapter_min_outside_chapters: int = 1
    tiny_body_max_abs_paras: int = 4
    tiny_body_max_ratio: float = 0.08
    tiny_body_min_candidates: int = 2
    tiny_body_primary_margin: int = 1
    keep_after_first_chapter: bool = True


@dataclass
class HeadingNumberingConfig:
    mode: str = "B"  # "A" or "B"
    scheme: str = "1"  # "1" 阿拉伯数字 / "2" 中文编号
    levels: dict[str, HeadingLevelConfig] = field(default_factory=dict)
    schemes: dict[str, dict[str, HeadingLevelConfig]] = field(default_factory=dict)
    enforcement: EnforcementConfig = field(default_factory=EnforcementConfig)
    risk_guard: HeadingRiskGuardConfig = field(default_factory=HeadingRiskGuardConfig)

    def apply_scheme(self, scheme_id: str | None = None) -> None:
        """将指定方案的 levels 设为当前活动 levels"""
        sid = scheme_id or self.scheme
        if sid in self.schemes:
            self.levels = dict(self.schemes[sid])
            self.scheme = sid


@dataclass
class HeadingModelConfig:
    """标题语义模型：统一层级映射、分区标题样式和无编号策略。"""
    level_to_word_style: dict[str, str] = field(default_factory=lambda: {
        "heading1": "Heading 1",
        "heading2": "Heading 2",
        "heading3": "Heading 3",
        "heading4": "Heading 4",
        "heading5": "Heading 5",
        "heading6": "Heading 6",
        "heading7": "Heading 7",
        "heading8": "Heading 8",
    })
    level_to_style_key: dict[str, str] = field(default_factory=lambda: {
        "heading1": "heading1",
        "heading2": "heading2",
        "heading3": "heading3",
        "heading4": "heading4",
        "heading5": "heading5",
        "heading6": "heading6",
        "heading7": "heading7",
        "heading8": "heading8",
    })
    style_alias_to_level: dict[str, str] = field(default_factory=lambda: {
        "一级标题": "heading1",
        "二级标题": "heading2",
        "三级标题": "heading3",
        "四级标题": "heading4",
        "五级标题": "heading5",
        "六级标题": "heading6",
        "七级标题": "heading7",
        "八级标题": "heading8",
        "章标题": "heading1",
        "节标题": "heading2",
        "标题 1": "heading1",
        "标题 2": "heading2",
        "标题 3": "heading3",
        "标题 4": "heading4",
    })
    section_title_style_map: dict[str, str] = field(default_factory=lambda: {
        "abstract_cn": "abstract_title_cn",
        "abstract_en": "abstract_title_en",
        "references": "heading1",
        "errata": "heading1",
        "appendix": "heading1",
        "acknowledgment": "heading1",
        "resume": "heading1",
    })
    non_numbered_title_sections: list[str] = field(default_factory=lambda: [
        "references", "errata", "appendix", "acknowledgment", "resume",
    ])
    non_numbered_title_texts: list[str] = field(default_factory=lambda: [
        "参考文献", "勘误页", "勘误", "附录", "致谢",
        "个人简历", "在学期间发表的学术论文与研究成果",
    ])
    front_matter_title_texts: list[str] = field(default_factory=lambda: [
        "摘要", "摘要。", "摘要：", "摘要:",
        "abstract", "目录", "目录：", "tableofcontents",
    ])
    post_section_types: list[str] = field(default_factory=lambda: [
        "references", "errata", "appendix", "acknowledgment", "resume",
    ])
    non_numbered_heading_style_name: str = "Heading 1 Unnumbered"
    header_front_text: dict[str, str] = field(default_factory=lambda: {
        "abstract_cn": "摘要",
        "abstract_en": "Abstract",
        "toc": "目录",
    })
    header_back_types: list[str] = field(default_factory=lambda: [
        "references", "errata", "acknowledgment", "appendix", "resume",
    ])


@dataclass
class StyleConfig:
    font_cn: str = "宋体"
    font_en: str = "Times New Roman"
    size_pt: float = 12
    size_display: str = ""
    bold: bool = False
    italic: bool = False
    alignment: str = "justify"
    first_line_indent_chars: float = 0
    first_line_indent_unit: str = "chars"
    left_indent_chars: float = 0
    left_indent_unit: str = "chars"
    right_indent_chars: float = 0
    right_indent_unit: str = "chars"
    hanging_indent_chars: float = 0
    hanging_indent_unit: str = "chars"
    special_indent_mode: str = "none"
    special_indent_value: float = 0
    special_indent_unit: str = "chars"
    line_spacing_type: str = "exact"  # exact=固定磅值; multiple=多倍行距（1.0/1.5/2.0/0.8...）
    line_spacing_pt: float = 20  # exact 时单位 pt；multiple 时表示倍数
    space_before_pt: float = 0
    space_after_pt: float = 0

    def __post_init__(self):
        sync_style_config_indent_fields(self)


@dataclass
class FormatScopeConfig:
    """排版作用域配置"""
    mode: str = "auto"  # "auto" 自动识别 / "manual" 手动指定
    page_ranges_text: str = ""  # 手动模式：修正页码范围（示例：27-40,44-56）
    body_start_index: int | None = None  # 手动模式：正文起始段落索引（直接指定，优先级最高）
    body_start_page: int | None = None   # 手动模式：正文起始页码（物理页，从1开始）
    body_start_keyword: str = ""  # 保留兼容：正文起始关键字（最低优先级）
    # 各分区是否参与排版
    sections: dict[str, bool] = field(default_factory=lambda: {
        "body": True,
        "references": True,
        "errata": True,
        "acknowledgment": True,
        "appendix": False,
        "abstract_cn": False,
        "abstract_en": False,
        "toc": False,
        "resume": False,
    })

    def is_section_enabled(self, section_type: str) -> bool:
        """检查某分区是否启用排版"""
        if str(section_type or "").strip().lower() == "cover":
            return False
        return self.sections.get(section_type, False)


@dataclass
class TocConfig:
    # word_native: Word 原生 TOC 域；plain: 普通目录条目（兼容旧实现）
    mode: str = "word_native"


@dataclass
class CaptionConfig:
    """图表题注配置"""
    enabled: bool = True
    auto_insert: bool = True           # 自动插入缺失题注（表头图尾）
    format_inserted: bool = False      # 插入题注默认不使用域代码（用纯文本编号）
    figure_prefix: str = "图"
    table_prefix: str = "表"
    separator: str = "\u3000"          # 序号与题目间全角空格
    placeholder: str = "[待补充]"      # 缺失题注的占位文本
    numbering_format: str = "chapter.seq"  # 章.序号


@dataclass
class ChemTypographyConfig:
    """化学式上下角标恢复（实验室功能）"""
    enabled: bool = False
    scopes: dict[str, bool] = field(default_factory=lambda: {
        "references": False,
        "body": False,
        "headings": False,
        "abstract_cn": False,
        "abstract_en": False,
        "captions": False,
        "tables": False,
    })
    # Force-allow list (higher priority than ignore/heuristics).
    allow_tokens: list[str] = field(default_factory=list)
    allow_patterns: list[str] = field(default_factory=list)
    ignore_tokens: list[str] = field(default_factory=list)
    ignore_patterns: list[str] = field(default_factory=list)
    # Per-token manual style override mask:
    # '^' => superscript, '_' => subscript, other chars => no style.
    manual_overrides: dict[str, str] = field(default_factory=dict)


@dataclass
class MdCleanupConfig:
    """Markdown 文本修复细项开关。"""
    enabled: bool = False
    # Preserve pre-existing Word list marker paragraphs when removing blank lines.
    preserve_existing_word_lists: bool = True
    # Collapse duplicated formula strings caused by Markdown->Word paste noise.
    formula_copy_noise_cleanup: bool = True
    # Remove fake ordered-list marker lines (e.g. "1.", "2.") around formula lines.
    suppress_formula_fake_lists: bool = True
    # List marker separator: "tab" | "half_space" | "full_space".
    list_marker_separator: str = "tab"
    # Ordered list style preset.
    # "mixed" | "decimal_dot" | "decimal_paren_right" | "decimal_cn_dun" | "decimal_full_paren"
    ordered_list_style: str = "mixed"
    # Unordered list style preset.
    # "word_default" | "bullet_dot" | "bullet_circle" | "bullet_square" | "bullet_dash"
    unordered_list_style: str = "word_default"


@dataclass
class WhitespaceNormalizeConfig:
    """空白与全半角规范（实验室功能，子项可选）"""
    enabled: bool = False
    normalize_space_variants: bool = True
    convert_tabs: bool = True
    remove_zero_width: bool = True
    collapse_multiple_spaces: bool = True
    trim_paragraph_edges: bool = True
    smart_full_half_convert: bool = True
    punctuation_by_context: bool = True
    bracket_by_inner_language: bool = True
    fullwidth_alnum_to_halfwidth: bool = True
    quote_by_context: bool = False
    protect_reference_numbering: bool = True
    context_min_confidence: int = 2


@dataclass
class CitationLinkConfig:
    """正文参考文献域关联（实验室功能）"""
    enabled: bool = True
    # Rebuild reference labels with SEQ fields and keep body citation numbers
    # synchronized via REF fields when references are inserted/reordered.
    auto_number_reference_entries: bool = True
    # Treat trailing page markers after bracket citations as citation superscript,
    # e.g. [320]198 -> superscript "198" when confidence is high.
    superscript_outer_page_numbers: bool = False


@dataclass
class FormulaConvertConfig:
    """Formula recognition and conversion configuration."""
    enabled: bool = False
    # "word_native" | "latex"
    output_mode: str = "word_native"
    # User requirement: low-confidence formulas are skipped and marked by default.
    low_confidence_policy: str = "skip_and_mark"
    # Optional post-save fallback: use local Word/MathType OLE subsystem to
    # extract MathML for unresolved MathType objects, then rewrite as OMML.
    office_fallback_enabled: bool = False
    office_fallback_timeout_sec: int = 30


@dataclass
class FormulaToTableConfig:
    """Formula-to-table conversion configuration."""
    enabled: bool = False
    # MVP stage only processes display/block equations.
    block_only: bool = True


@dataclass
class FormulaStyleConfig:
    """Formula style normalization configuration."""
    enabled: bool = False
    unify_font: bool = True
    unify_size: bool = True
    unify_spacing: bool = True


@dataclass
class EquationTableFormatConfig:
    """Equation-table numbering configuration."""
    enabled: bool = False
    numbering_format: str = "chapter.seq"


@dataclass
class FormulaTableConfig:
    """Formula-table visual configuration shared by formula rules."""
    formula_font_name: str = "Cambria Math"
    formula_font_size_pt: float = 12.0
    formula_font_size_display: str = ""
    formula_line_spacing: float = 1.0
    formula_space_before_pt: float = 0.0
    formula_space_after_pt: float = 0.0
    block_alignment: str = "center"
    table_alignment: str = "center"
    formula_cell_alignment: str = "center"
    number_alignment: str = "right"
    number_font_name: str = "Times New Roman"
    number_font_size_pt: float = 10.5
    number_font_size_display: str = ""
    auto_shrink_number_column: bool = True


@dataclass
class OutputConfig:
    final_docx: bool = True
    compare_docx: bool = True
    compare_text: bool = True
    compare_formatting: bool = True
    report_json: bool = True
    report_markdown: bool = True


# 默认执行流程步骤（单一来源，pipeline.py 也引用此列表）
DEFAULT_PIPELINE_STEPS: list[str] = [
    "page_setup", "md_cleanup", "style_manager",
    "heading_detect", "heading_numbering", "toc_format",
    "caption_format", "table_format", "section_format",
    "header_footer", "validation",
]


@dataclass
class SceneConfig:
    version: str = "1.0"
    name: str = ""
    description: str = ""
    format_signature: str = ""
    category: str = "general"
    category_label: str = "通用文档"
    capabilities: dict[str, bool] = field(default_factory=lambda: {
        "heading_numbering": True,
        "section_detection": True,
        "caption": True,
        "header_footer": True,
        "toc_rebuild": True,
        "equation_numbering": True,
        "md_cleanup": True,
        "whitespace_normalize": True,
        "citation_link_restore": True,
        "chem_typography_restore": True,
    })
    available_sections: list[str] = field(default_factory=lambda: [
        "body", "references", "errata", "acknowledgment", "appendix",
        "resume", "abstract_cn", "abstract_en", "toc",
    ])
    page_setup: PageSetupConfig = field(default_factory=PageSetupConfig)
    heading_numbering: HeadingNumberingConfig = field(default_factory=HeadingNumberingConfig)
    heading_numbering_v2: HeadingNumberingV2Config = field(default_factory=HeadingNumberingV2Config)
    heading_model: HeadingModelConfig = field(default_factory=HeadingModelConfig)
    toc: TocConfig = field(default_factory=TocConfig)
    caption: CaptionConfig = field(default_factory=CaptionConfig)
    chem_typography: ChemTypographyConfig = field(default_factory=ChemTypographyConfig)
    md_cleanup: MdCleanupConfig = field(default_factory=MdCleanupConfig)
    whitespace_normalize: WhitespaceNormalizeConfig = field(default_factory=WhitespaceNormalizeConfig)
    citation_link: CitationLinkConfig = field(default_factory=CitationLinkConfig)
    formula_convert: FormulaConvertConfig = field(default_factory=FormulaConvertConfig)
    formula_to_table: FormulaToTableConfig = field(default_factory=FormulaToTableConfig)
    equation_table_format: EquationTableFormatConfig = field(default_factory=EquationTableFormatConfig)
    formula_style: FormulaStyleConfig = field(default_factory=FormulaStyleConfig)
    formula_table: FormulaTableConfig = field(default_factory=FormulaTableConfig)
    format_scope: FormatScopeConfig = field(default_factory=FormatScopeConfig)
    # 常规表格排版风格：compact=适宜压缩（默认），full=铺满页面
    normal_table_layout_mode: str = "smart"
    normal_table_smart_levels: int = 4
    # 常规表格边框样式：full_grid=全框线, three_line=三线表, keep=不改边框
    normal_table_border_mode: str = "three_line"
    # 全框线线宽（磅）
    table_border_width_pt: float = 0.5
    # 三线表：表头线宽（磅，首行上下均使用）
    three_line_header_width_pt: float = 1.0
    # 三线表：表尾线宽（磅）
    three_line_bottom_width_pt: float = 0.5
    # 常规表格内文字行距：single=单倍（默认）, one_half=1.5倍, double=双倍
    normal_table_line_spacing_mode: str = "single"
    # 常规表格跨页时重复首行表头（Word“重复标题行”）
    normal_table_repeat_header: bool = False
    update_header: bool = True         # 自动更新页眉
    update_page_number: bool = True    # 自动更新页码
    update_header_line: bool = True    # 页眉底部横线
    styles: dict[str, StyleConfig] = field(default_factory=dict)
    output: OutputConfig = field(default_factory=OutputConfig)
    pipeline: list[str] = field(default_factory=lambda: list(DEFAULT_PIPELINE_STEPS))
    # 严格模式：关键规则失败时，PipelineResult.success=False
    pipeline_strict_mode: bool = True
    # 关键规则清单（可通过场景配置覆盖）
    pipeline_critical_rules: list[str] = field(default_factory=lambda: list(DEFAULT_PIPELINE_STEPS))
