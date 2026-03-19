"""页面设置规则"""

from docx import Document
from docx.shared import Cm
from src.engine.rules.base import BaseRule, ValidationIssue
from src.engine.change_tracker import ChangeTracker
from src.engine.rules.header_footer import _apply_part_style
from src.scene.schema import SceneConfig
from src.utils.constants import CM_TO_TWIPS

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_PAPER_SIZE_CM = {
    "A4": (21.0, 29.7),
    "LETTER": (21.59, 27.94),
    "A3": (29.7, 42.0),
}


def _w(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"


class PageSetupRule(BaseRule):
    name = "page_setup"
    description = "设置页面尺寸、边距、页眉页脚距"

    def apply(self, doc: Document, config: SceneConfig,
              tracker: ChangeTracker, context: dict) -> None:
        ps = config.page_setup
        paper_key_raw = str(getattr(ps, "paper_size", "A4")).strip().upper()
        paper_key = paper_key_raw if paper_key_raw in _PAPER_SIZE_CM else "A4"
        paper_w_cm, paper_h_cm = _PAPER_SIZE_CM[paper_key]
        for section in doc.sections:
            old_top = section.top_margin
            is_landscape = section.page_width > section.page_height
            if is_landscape:
                section.page_width = Cm(paper_h_cm)
                section.page_height = Cm(paper_w_cm)
            else:
                section.page_width = Cm(paper_w_cm)
                section.page_height = Cm(paper_h_cm)
            section.top_margin = Cm(ps.margin.top_cm)
            section.bottom_margin = Cm(ps.margin.bottom_cm)
            section.left_margin = Cm(ps.margin.left_cm)
            section.right_margin = Cm(ps.margin.right_cm)
            section.gutter = Cm(ps.gutter_cm)
            section.header_distance = Cm(ps.header_distance_cm)
            section.footer_distance = Cm(ps.footer_distance_cm)

            tracker.record(
                rule_name=self.name,
                target=f"Section",
                section="global",
                change_type="format",
                before=f"top={old_top}",
                after=f"paper={paper_key}, top={section.top_margin}",
                paragraph_index=-1,
            )

        # 格式化页眉/页脚字体
        self._format_headers_footers(doc, config, tracker)

    def _format_headers_footers(self, doc, config, tracker):
        """格式化已有页眉/页脚中的文字字体（不创建新的页眉页脚）"""
        header_sc = config.styles.get("header_cn")
        footer_sc = config.styles.get("page_number")
        if not header_sc and not footer_sc:
            return

        count = 0
        # 直接遍历 XML 中已存在的 headerReference / footerReference
        # 通过 python-docx 的 part 访问已有的页眉页脚
        for section in doc.sections:
            # 页眉：检查 sectPr 中是否有 headerReference
            sect_el = section._sectPr
            for ref in sect_el.findall(_w("headerReference")):
                r_id = ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if r_id and header_sc:
                    try:
                        part = doc.part.related_parts[r_id]
                        _apply_part_style(part.element, header_sc, default_alignment="center")
                        count += 1
                    except (KeyError, AttributeError):
                        pass

            # 页脚
            for ref in sect_el.findall(_w("footerReference")):
                r_id = ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if r_id and footer_sc:
                    try:
                        part = doc.part.related_parts[r_id]
                        _apply_part_style(part.element, footer_sc, default_alignment="center")
                        count += 1
                    except (KeyError, AttributeError):
                        pass

        if count:
            tracker.record(
                rule_name=self.name,
                target=f"{count} 个页眉/页脚",
                section="global",
                change_type="format",
                before="(mixed fonts)",
                after=f"header={header_sc.size_pt if header_sc else '?'}pt",
                paragraph_index=-1,
            )
