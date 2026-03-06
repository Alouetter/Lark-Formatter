"""常量定义：OOXML 命名空间、样式名映射等"""

# OOXML 命名空间
OOXML_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}

# python-docx 内置样式名 → 我们的逻辑名映射
STYLE_MAP = {
    "heading1": "Heading 1",
    "heading2": "Heading 2",
    "heading3": "Heading 3",
    "heading4": "Heading 4",
    "heading5": "Heading 5",
    "heading6": "Heading 6",
    "heading7": "Heading 7",
    "heading8": "Heading 8",
    "normal": "Normal",
}

# 逻辑层级名 → Word Heading 层级
LEVEL_TO_HEADING = {
    "heading1": 1,
    "heading2": 2,
    "heading3": 3,
    "heading4": 4,
    "heading5": 5,
    "heading6": 6,
    "heading7": 7,
    "heading8": 8,
}

# 单位转换
CM_TO_EMU = 914400 / 2.54  # 1cm = 360000 EMU
CM_TO_TWIPS = 567  # 1cm ≈ 567 twips
PT_TO_HALF_PT = 2  # python-docx 用半磅
PT_TO_TWIPS = 20  # 1pt = 20 twips

# 修订标记作者
REVISION_AUTHOR = "DocFormatter"
