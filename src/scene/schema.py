"""场景配置 dataclass 定义"""

from dataclasses import dataclass, field


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
class HeadingLevelConfig:
    format: str = "arabic_dotted"
    template: str = "{n}"
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
    bold: bool = False
    italic: bool = False
    alignment: str = "justify"
    first_line_indent_chars: float = 0
    left_indent_chars: float = 0
    line_spacing_type: str = "exact"  # exact=固定磅值; multiple=多倍行距（1.0/1.5/2.0/0.8...）
    line_spacing_pt: float = 20  # exact 时单位 pt；multiple 时表示倍数
    space_before_pt: float = 0
    space_after_pt: float = 0


@dataclass
class FormatScopeConfig:
    """排版作用域配置"""
    mode: str = "auto"  # "auto" 自动识别 / "manual" 手动指定
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
        return self.sections.get(section_type, False)


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
    enabled: bool = True
    scopes: dict[str, bool] = field(default_factory=lambda: {
        "references": True,
        "body": False,
        "headings": False,
        "abstract_cn": False,
        "abstract_en": False,
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
class WhitespaceNormalizeConfig:
    """空白字符统一（实验室功能，子项可选）"""
    enabled: bool = True
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
    heading_model: HeadingModelConfig = field(default_factory=HeadingModelConfig)
    caption: CaptionConfig = field(default_factory=CaptionConfig)
    chem_typography: ChemTypographyConfig = field(default_factory=ChemTypographyConfig)
    whitespace_normalize: WhitespaceNormalizeConfig = field(default_factory=WhitespaceNormalizeConfig)
    citation_link: CitationLinkConfig = field(default_factory=CitationLinkConfig)
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
