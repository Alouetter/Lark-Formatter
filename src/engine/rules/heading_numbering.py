"""标题编号修复规则：模式A(保留原编号) / 模式B(重建编号，原生多级列表)"""

import re
from docx import Document
from src.engine.rules.base import BaseRule
from src.engine.change_tracker import ChangeTracker
from src.scene.schema import SceneConfig
from src.scene.heading_model import get_level_to_style_key, get_level_to_word_style
from docx.enum.text import WD_ALIGN_PARAGRAPH
from src.utils.ooxml import (
    register_numbering, link_style_to_numbering, LEVEL_NAME_TO_ILVL,
)

ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# 编号前缀拆分：中文编号与阿拉伯编号分开处理，避免误拆 “2024年工作总结” 这类正文。
_CN_NUMBERING_SPLIT_RE = re.compile(
    r"^(?P<num>"
    r"第[一二三四五六七八九十百]+章"
    r"|第\d+章"
    r"|第[一二三四五六七八九十百]+节"
    r"|第\d+节"
    r"|（[一二三四五六七八九十百]+）"
    r"|\([一二三四五六七八九十百]+\)"
    r"|[一二三四五六七八九十百]+、"
    r")(?P<sep>[\s\u3000\t]*)(?P<title>.+)$"
)
_ARABIC_MULTI_LEVEL_SPLIT_RE = re.compile(
    r"^(?P<num>\d+\.\d+(?:\.\d+)*)"
    r"(?P<sep>[\s\u3000\t]+)"
    r"(?P<title>.+)$"
)
_ARABIC_INTEGER_SPLIT_RE = re.compile(
    r"^(?P<num>\d+)"
    r"(?P<sep>(?:[\s\u3000\t]+|(?:[、．)]\s*)|(?:\.\s+)))"
    r"(?P<title>.+)$"
)

def split_heading_text(text: str) -> tuple[str, str, str]:
    """将标题文本拆分为 (编号, 分隔符, 标题内容)"""
    raw = text.strip()
    if not raw:
        return "", "", ""

    m = _CN_NUMBERING_SPLIT_RE.match(raw)
    if m:
        return m.group("num"), m.group("sep"), m.group("title")

    m = _ARABIC_MULTI_LEVEL_SPLIT_RE.match(raw)
    if m:
        return m.group("num"), m.group("sep"), m.group("title")

    m = _ARABIC_INTEGER_SPLIT_RE.match(raw)
    if m:
        return m.group("num"), m.group("sep"), m.group("title")

    return "", "", raw


class HeadingNumberingRule(BaseRule):
    name = "heading_numbering"
    description = "标题编号修复（模式A/B，原生多级列表）"

    def apply(self, doc: Document, config: SceneConfig,
              tracker: ChangeTracker, context: dict) -> None:
        headings = context.get("headings", [])
        if not headings:
            return

        mode = config.heading_numbering.mode
        levels_config = config.heading_numbering.levels
        enforcement = config.heading_numbering.enforcement

        if mode == "B":
            self._apply_mode_b(doc, headings, levels_config,
                               enforcement, tracker, config)
        else:
            self._apply_mode_a(doc, headings, levels_config,
                               enforcement, tracker, config)

    def _apply_mode_b(self, doc, headings, levels_config,
                      enforcement, tracker, config):
        """模式B：用 Word 原生多级列表重建编号体系"""
        level_style_map = get_level_to_word_style(config)
        level_style_key_map = get_level_to_style_key(config)
        # 1. 在 numbering.xml 中注册多级列表定义（含样式格式）
        num_id = register_numbering(
            doc,
            levels_config,
            config.styles,
            level_to_style_key=level_style_key_map,
        )

        # 2. 将 Heading 样式链接到该编号定义
        for level_name, ilvl in LEVEL_NAME_TO_ILVL.items():
            if level_name not in levels_config:
                continue
            style_name = level_style_map.get(level_name)
            if not style_name:
                continue
            link_style_to_numbering(doc, style_name, num_id, ilvl)
        # 按方案覆盖样式对齐（否则样式默认对齐会覆盖编号层定义）
        self._sync_heading_style_alignment(doc, levels_config, level_style_map)

        # 3. 逐段落：剥离旧编号文本，应用 Heading 样式
        for h in headings:
            para = doc.paragraphs[h.para_index]
            old_text = para.text
            number_part, _, title_text = split_heading_text(old_text)

            # 剥离编号前缀，只保留标题纯文本（保留 run 结构以保持下标等格式）
            if number_part:
                strip_len = len(number_part) + len(_)
                self._strip_prefix_runs(para, strip_len)
            else:
                # 无编号前缀时，也要剥离前导全角/半角空白（避免与编号定义的分隔符重复）
                self._lstrip_runs(para)

            # 应用对应 Heading 样式（样式已绑定编号）
            style_name = level_style_map.get(h.level)
            if style_name:
                try:
                    para.style = doc.styles[style_name]
                except KeyError:
                    pass
                # 清除段落级 numPr 覆盖，让样式的 numPr 生效
                self._clear_para_numpr(para)
                # 清除 run 级别字体覆盖，让样式字体生效
                self._clear_run_fonts(para)

            tracker.record(
                rule_name=self.name,
                target=f"段落 #{h.para_index}",
                section="body",
                change_type="numbering",
                before=old_text[:80],
                after=f"原生编号 {style_name} ilvl={LEVEL_NAME_TO_ILVL.get(h.level, 0)}, 文本=\"{title_text[:40]}\"",
                paragraph_index=h.para_index,
            )

    def _apply_mode_a(self, doc, headings, levels_config,
                      enforcement, tracker, config):
        """模式A：保留原编号文本，仅应用 Heading 样式"""
        level_style_map = get_level_to_word_style(config)
        # 模式A也应尊重编号级别中的 alignment 配置（例如 right）。
        self._sync_heading_style_alignment(doc, levels_config, level_style_map)
        for h in headings:
            para = doc.paragraphs[h.para_index]
            old_text = para.text

            # 应用样式
            style_name = level_style_map.get(h.level)
            if style_name:
                try:
                    para.style = doc.styles[style_name]
                except KeyError:
                    pass
                self._clear_run_fonts(para)

            tracker.record(
                rule_name=self.name,
                target=f"段落 #{h.para_index}",
                section="body",
                change_type="style",
                before=old_text[:80],
                after=f"样式→{style_name}",
                paragraph_index=h.para_index,
            )

    @staticmethod
    def _sync_heading_style_alignment(doc, levels_config, level_style_map: dict[str, str]) -> None:
        for level_name, style_name in level_style_map.items():
            lc = levels_config.get(level_name)
            if lc is None:
                continue
            align_key = str(getattr(lc, "alignment", "") or "").strip().lower()
            if align_key not in ALIGNMENT_MAP:
                continue
            try:
                style = doc.styles[style_name]
            except KeyError:
                continue
            style.paragraph_format.alignment = ALIGNMENT_MAP[align_key]

    @staticmethod
    def _clear_para_numpr(para) -> None:
        """清除段落级 numPr，避免覆盖样式中的编号定义。

        原文档中部分段落自带 numPr（指向旧编号），会覆盖样式的 numPr，
        导致编号显示异常。
        """
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ppr = para._element.find(f"{{{W}}}pPr")
        if ppr is not None:
            numpr = ppr.find(f"{{{W}}}numPr")
            if numpr is not None:
                ppr.remove(numpr)

    @staticmethod
    def _clear_run_fonts(para) -> None:
        """清除段落所有 run 的字体/字号/颜色覆盖，让 Heading 样式生效。

        run 级别的 rFonts/sz/color 会覆盖样式定义，
        导致标题文本字体不一致或颜色异常。
        含 vertAlign（上下标）的 run 保留 sz，避免破坏 H2O2 等化学式格式。
        """
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        # 清除段落级 rPr 直改格式（会影响自动编号的字号/上下标等显示）
        ppr = para._element.find(f"{{{W}}}pPr")
        if ppr is not None:
            para_rpr = ppr.find(f"{{{W}}}rPr")
            if para_rpr is not None:
                ppr.remove(para_rpr)
        # 清除 run 级覆盖
        for run in para.runs:
            rpr = run._element.find(f"{{{W}}}rPr")
            if rpr is None:
                continue
            has_vert = rpr.find(f"{{{W}}}vertAlign") is not None
            for tag in ("rFonts", "color"):
                el = rpr.find(f"{{{W}}}{tag}")
                if el is not None:
                    rpr.remove(el)
            # 有上下标的 run 保留 sz，否则清除
            if not has_vert:
                for tag in ("sz", "szCs"):
                    el = rpr.find(f"{{{W}}}{tag}")
                    if el is not None:
                        rpr.remove(el)

    @staticmethod
    def _replace_paragraph_text(para, new_text: str) -> None:
        """替换段落文本，尽量保留第一个 run 的格式"""
        if not para.runs:
            para.text = new_text
            return
        first_run = para.runs[0]
        for run in para.runs[1:]:
            run.text = ""
        first_run.text = new_text

    @staticmethod
    def _strip_prefix_runs(para, strip_len: int) -> None:
        """从段落 runs 头部剥离 strip_len 个字符，保留剩余 run 结构（下标等格式不丢失）。

        剥离完成后，自动去除第一个非空 run 的前导空白（避免与编号定义的分隔符重复）。
        """
        remaining = strip_len
        for run in para.runs:
            if remaining <= 0:
                break
            txt = run.text
            if len(txt) <= remaining:
                remaining -= len(txt)
                run.text = ""
            else:
                run.text = txt[remaining:]
                remaining = 0

        # 去除第一个非空 run 的前导全角/半角空白
        for run in para.runs:
            if run.text:
                run.text = run.text.lstrip(" \t\u3000")
                break

    @staticmethod
    def _lstrip_runs(para) -> None:
        """剥离段落 runs 头部的前导全角/半角空白"""
        for run in para.runs:
            if run.text:
                run.text = run.text.lstrip(" \t\u3000")
                break
