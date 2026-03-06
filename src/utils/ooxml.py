"""OOXML 底层操作工具：修订标记注入、XML 元素构建等"""

from datetime import datetime, timezone
from lxml import etree
from src.utils.constants import OOXML_NS, REVISION_AUTHOR

# 全局修订 ID 计数器
_revision_id_counter = 0


def next_revision_id() -> int:
    global _revision_id_counter
    _revision_id_counter += 1
    return _revision_id_counter


def reset_revision_counter():
    global _revision_id_counter
    _revision_id_counter = 0


def revision_date() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ── 命名空间工具 ──

W_NS = OOXML_NS["w"]


def _w(tag: str) -> str:
    """生成 w: 命名空间限定标签名"""
    return f"{{{W_NS}}}{tag}"


# ── 多级列表编号 ──

# 场景 format → OOXML numFmt
FORMAT_TO_NUMFMT = {
    "chinese_chapter": "chineseCounting",
    "chinese_section": "chineseCounting",
    "chinese_ordinal": "chineseCounting",
    "chinese_ordinal_paren": "chineseCounting",
    "arabic_dotted": "decimal",
    "arabic": "decimal",
}

# 层级名 → ilvl
LEVEL_NAME_TO_ILVL = {
    "heading1": 0,
    "heading2": 1,
    "heading3": 2,
    "heading4": 3,
    "heading5": 4,
    "heading6": 5,
    "heading7": 6,
    "heading8": 7,
}


def build_lvl_text(level_name: str, template: str) -> str:
    """将场景模板转换为 OOXML lvlText 格式。

    例:
      chapter, "第{cn}章"     → "第%1章"
      level1,  "{parent}.{n}" → "%1.%2"
      level1,  "第{cn}节"     → "第%2节"
      level2,  "{cn}、"       → "%3、"
    """
    ilvl = LEVEL_NAME_TO_ILVL[level_name]
    current = f"%{ilvl + 1}"
    # {cn} 和 {n} 都映射到当前层级的编号引用
    result = template.replace("{cn}", current).replace("{n}", current)
    # {parent} 展开为上层编号引用链
    if "{parent}" in template:
        parent_parts = ".".join(f"%{i+1}" for i in range(ilvl))
        result = result.replace("{parent}", parent_parts)
    return result


def _add_lvl_ppr(lvl, style_config, *, alignment: str = "",
                  left_indent_chars: float = 0) -> None:
    """为 lvl 添加段落格式 pPr（间距、对齐、缩进）"""
    ppr = etree.SubElement(lvl, _w("pPr"))
    # 间距
    spacing = etree.SubElement(ppr, _w("spacing"))
    before = int(style_config.space_before_pt * 20)
    after = int(style_config.space_after_pt * 20)
    if before:
        spacing.set(_w("before"), str(before))
    if after:
        spacing.set(_w("after"), str(after))
    # 对齐
    if alignment:
        jc = etree.SubElement(ppr, _w("jc"))
        jc.set(_w("val"), alignment)
    # 缩进：显式归零防止 Word 添加默认悬挂缩进
    ind = etree.SubElement(ppr, _w("ind"))
    if left_indent_chars > 0:
        ind.set(_w("leftChars"), str(int(left_indent_chars * 100)))
    else:
        ind.set(_w("left"), "0")
    ind.set(_w("hanging"), "0")


def _add_lvl_rpr(lvl, style_config) -> None:
    """为 lvl 添加字体格式 rPr（字体、字号、加粗）"""
    rpr = etree.SubElement(lvl, _w("rPr"))
    # 字体
    rfonts = etree.SubElement(rpr, _w("rFonts"))
    if style_config.font_en:
        rfonts.set(_w("ascii"), style_config.font_en)
        rfonts.set(_w("hAnsi"), style_config.font_en)
    if style_config.font_cn:
        rfonts.set(_w("eastAsia"), style_config.font_cn)
    # 字号 (半磅单位)
    if style_config.size_pt:
        sz = etree.SubElement(rpr, _w("sz"))
        sz.set(_w("val"), str(int(style_config.size_pt * 2)))
        sz_cs = etree.SubElement(rpr, _w("szCs"))
        sz_cs.set(_w("val"), str(int(style_config.size_pt * 2)))
    # 加粗
    if style_config.bold:
        etree.SubElement(rpr, _w("b"))
        etree.SubElement(rpr, _w("bCs"))


def _build_lvl_element(ilvl: int, num_fmt: str, lvl_text: str,
                       separator: str, *,
                       is_legal: bool = False,
                       style_config=None,
                       alignment: str = "",
                       left_indent_chars: float = 0) -> etree._Element:
    """构建单个 <w:lvl> 元素。

    is_legal: 为 True 时添加 <w:isLgl/>，强制父级编号用 decimal 显示。
    style_config: StyleConfig，用于设置该级标题的段落/字体格式。
    alignment: 段落对齐方式覆盖（"left"/"center"/"right"/"justify"）。
    left_indent_chars: 左缩进（汉字数）。
    """
    lvl = etree.SubElement(etree.Element("dummy"), _w("lvl"))
    lvl.set(_w("ilvl"), str(ilvl))

    start = etree.SubElement(lvl, _w("start"))
    start.set(_w("val"), "1")

    fmt = etree.SubElement(lvl, _w("numFmt"))
    fmt.set(_w("val"), num_fmt)

    if is_legal:
        etree.SubElement(lvl, _w("isLgl"))

    full_text = lvl_text + separator
    lt = etree.SubElement(lvl, _w("lvlText"))
    lt.set(_w("val"), full_text)

    jc = etree.SubElement(lvl, _w("lvlJc"))
    jc.set(_w("val"), "left")

    # suff=nothing：分隔符已嵌入 lvlText，不需要 Word 额外添加
    suff = etree.SubElement(lvl, _w("suff"))
    suff.set(_w("val"), "nothing")

    # 段落格式 pPr
    if style_config:
        _add_lvl_ppr(lvl, style_config,
                     alignment=alignment,
                     left_indent_chars=left_indent_chars)

    # 字体格式 rPr
    if style_config:
        _add_lvl_rpr(lvl, style_config)

    return lvl


def build_abstract_num(abstract_num_id: int,
                       levels_config: dict,
                       styles_config: dict | None = None,
                       level_order=["heading1", "heading2", "heading3", "heading4", "heading5", "heading6", "heading7", "heading8"],
                       level_to_style_key: dict[str, str] | None = None,
                       ) -> etree._Element:
    """构建完整的 <w:abstractNum> 元素。

    levels_config: {level_name: HeadingLevelConfig, ...}
    styles_config: {level_name: StyleConfig, ...} 用于设置各级段落/字体格式
    """
    # 层级名 → 样式配置名的映射
    default_style_key_map = {
        "heading1": "heading1",
        "heading2": "heading2",
        "heading3": "heading3",
        "heading4": "heading4",
        "heading5": "heading5",
        "heading6": "heading6",
        "heading7": "heading7",
        "heading8": "heading8",
    }
    style_key_map = dict(default_style_key_map)
    if isinstance(level_to_style_key, dict):
        for raw_level, raw_key in level_to_style_key.items():
            level = str(raw_level or "").strip()
            key = str(raw_key or "").strip()
            if level and key:
                style_key_map[level] = key

    abs_num = etree.Element(_w("abstractNum"))
    abs_num.set(_w("abstractNumId"), str(abstract_num_id))

    mlt = etree.SubElement(abs_num, _w("multiLevelType"))
    mlt.set(_w("val"), "multilevel")

    for level_name in level_order:
        lc = levels_config.get(level_name)
        if not lc:
            continue
        ilvl = LEVEL_NAME_TO_ILVL[level_name]
        num_fmt = FORMAT_TO_NUMFMT.get(lc.format, "decimal")
        lvl_text = build_lvl_text(level_name, lc.template)
        sep = lc.effective_separator

        # 仅当 lvlText 引用了父级编号时才需要 isLgl
        # 例如 "%1.%2" 在 ilvl=1 时引用了 %1（父级）
        is_legal = False
        if ilvl > 0:
            for parent_ilvl in range(ilvl):
                if f"%{parent_ilvl + 1}" in lvl_text:
                    is_legal = True
                    break

        # 获取该级别对应的样式配置
        sc = None
        if styles_config:
            style_key = style_key_map.get(level_name)
            sc = styles_config.get(style_key)

        lvl_el = _build_lvl_element(
            ilvl, num_fmt, lvl_text, sep,
            is_legal=is_legal, style_config=sc,
            alignment=lc.alignment,
            left_indent_chars=lc.left_indent_chars,
        )
        abs_num.append(lvl_el)

    return abs_num


def _next_id(numbering_el, tag: str, attr: str) -> int:
    """获取 numbering.xml 中下一个可用 ID"""
    max_id = 0
    for el in numbering_el.findall(_w(tag)):
        val = int(el.get(_w(attr), "0"))
        if val > max_id:
            max_id = val
    return max_id + 1


def register_numbering(doc, levels_config: dict,
                       styles_config: dict | None = None,
                       *,
                       level_to_style_key: dict[str, str] | None = None) -> int:
    """在文档中注册多级列表编号定义，返回 numId。

    自动处理 numbering part 的获取/创建。
    styles_config: 场景的 styles 字典，用于设置各级段落格式。
    """
    numbering_part = doc.part.numbering_part
    numbering_el = numbering_part.element

    abs_id = _next_id(numbering_el, "abstractNum", "abstractNumId")
    num_id = _next_id(numbering_el, "num", "numId")

    # 构建并插入 abstractNum
    abs_num = build_abstract_num(
        abs_id,
        levels_config,
        styles_config,
        level_to_style_key=level_to_style_key,
    )
    # abstractNum 必须在 num 之前
    first_num = numbering_el.find(_w("num"))
    if first_num is not None:
        first_num.addprevious(abs_num)
    else:
        numbering_el.append(abs_num)

    # 构建并插入 num
    num_el = etree.SubElement(numbering_el, _w("num"))
    num_el.set(_w("numId"), str(num_id))
    abs_ref = etree.SubElement(num_el, _w("abstractNumId"))
    abs_ref.set(_w("val"), str(abs_id))

    return num_id


def link_style_to_numbering(doc, style_name: str, num_id: int,
                            ilvl: int) -> None:
    """将 Word 样式链接到多级列表编号。

    在样式的 pPr 中添加 <w:numPr> 引用。
    """
    try:
        style = doc.styles[style_name]
    except KeyError:
        return

    style_el = style.element
    ppr = style_el.find(_w("pPr"))
    if ppr is None:
        ppr = etree.SubElement(style_el, _w("pPr"))

    # 移除已有的 numPr
    old = ppr.find(_w("numPr"))
    if old is not None:
        ppr.remove(old)

    num_pr = etree.SubElement(ppr, _w("numPr"))
    ilvl_el = etree.SubElement(num_pr, _w("ilvl"))
    ilvl_el.set(_w("val"), str(ilvl))
    nid = etree.SubElement(num_pr, _w("numId"))
    nid.set(_w("val"), str(num_id))
