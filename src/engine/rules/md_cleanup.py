"""Markdown 粘贴修复规则：检测并转换段落中的 Markdown 语法"""

import re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree

from src.engine.rules.base import BaseRule
from src.engine.change_tracker import ChangeTracker
from src.scene.schema import SceneConfig, StyleConfig
from src.scene.heading_model import get_level_to_word_style
from src.engine.doc_tree import DocTree
from src.markdown.block_parser import parse_docx_paragraphs
from src.markdown.inline_parser import parse_inline
from src.markdown.ir import MarkdownBlock, BlockType, InlineSpan, InlineType
from src.markdown.word_render import (
    write_spans, set_no_proof, add_hyperlink, write_table_cell,
    register_list_numbering, apply_num_pr, apply_blockquote_border,
    _apply_inline_format,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Markdown 标题级别 → 场景 heading level key
HEADING_LEVEL_KEY_MAP = {
    1: "heading1", 2: "heading2",
    3: "heading3", 4: "heading4",
    5: "heading4", 6: "heading4",
}


class MdCleanupRule(BaseRule):
    name = "md_cleanup"
    description = "检测并转换 Markdown 语法"

    def apply(self, doc: Document, config: SceneConfig,
              tracker: ChangeTracker, context: dict) -> None:
        doc_tree: DocTree = context.get("doc_tree")
        body = doc_tree.get_section("body") if doc_tree else None

        # 无 body 分区时不处理，防止跨分区误操作
        if not body:
            return

        # 先把段落中的手动换行（Shift+Enter）拆成真实段落（Enter）。
        # 这一步可避免后续规则把“下箭头换行”继续传递到最终文档。
        split_count = self._split_manual_break_paragraphs(
            doc, body.start_index, body.end_index)
        body_end_index = min(len(doc.paragraphs) - 1, body.end_index + split_count)
        if split_count:
            tracker.record(
                rule_name=self.name,
                target=f"body 段落手动换行拆分",
                section="body",
                change_type="format",
                before=f"{split_count} 处段内换行",
                after="→ 拆分为真实段落（Enter）",
                paragraph_index=body.start_index,
            )

        # 收集作用范围内的段落文本（跳过已有 Heading 样式的段落）
        para_texts: list[tuple[int, str]] = []
        for i, para in enumerate(doc.paragraphs):
            if not (body.start_index <= i <= body_end_index):
                continue
            # 已有 Heading 样式的段落不参与 Markdown 解析
            style_name = para.style.name if para.style else ""
            if style_name.startswith("Heading"):
                continue
            para_texts.append((i, para.text))

        # 解析为 IR
        blocks = parse_docx_paragraphs(para_texts)

        # 预注册列表编号
        bullet_num_id = register_list_numbering(doc, "bullet")
        decimal_num_id = register_list_numbering(doc, "decimal")

        # 收集脚注定义
        footnote_defs: dict[str, str] = {}
        for block in blocks:
            if block.type == BlockType.FOOTNOTE_DEF:
                footnote_defs[block.list_marker] = block.raw_text

        # 从场景配置获取正文字体（供 write_spans 使用）
        normal_sc = config.styles.get("normal")
        base_font_name = normal_sc.font_en if normal_sc else None
        base_font_size = Pt(normal_sc.size_pt) if normal_sc and normal_sc.size_pt else None

        render_ctx = {
            "bullet_num_id": bullet_num_id,
            "decimal_num_id": decimal_num_id,
            "footnote_defs": footnote_defs,
            "base_font_name": base_font_name,
            "base_font_size": base_font_size,
            "heading_style_map": get_level_to_word_style(config),
        }

        # 过滤有效块（包含 BLANK / FOOTNOTE_DEF），倒序处理（避免索引偏移）
        md_blocks = [b for b in blocks if b.source_para_indices]
        md_blocks.sort(
            key=lambda b: b.source_para_indices[0], reverse=True)

        for block in md_blocks:
            self._apply_block(doc, block, config, tracker, render_ctx)

        # md_cleanup 删除/合并了段落，doc_tree 索引已失效
        # 重建 doc_tree 供后续规则（section_format 等）使用
        if doc_tree:
            scope = context.get("format_scope", config.format_scope)
            body_start = None
            if scope.mode == "manual":
                if scope.body_start_index is not None:
                    body_start = scope.body_start_index
                elif scope.body_start_page is not None:
                    body_start = self._find_page_start_index(
                        doc, scope.body_start_page)
                elif scope.body_start_keyword:
                    body_start = self._find_keyword_index(
                        doc, scope.body_start_keyword)
            doc_tree.build(doc, body_start_index=body_start)
            context["doc_tree"] = doc_tree

    def _apply_block(self, doc, block, config, tracker, ctx):
        """按类型分发处理"""
        handlers = {
            BlockType.HEADING: self._apply_heading,
            BlockType.CODE_BLOCK: self._apply_code_block,
            BlockType.TABLE: self._apply_table,
            BlockType.BLOCKQUOTE: self._apply_blockquote,
            BlockType.HORIZONTAL_RULE: self._apply_hr,
            BlockType.LIST_ITEM: self._apply_list_item,
            BlockType.TASK_LIST_ITEM: self._apply_task_list,
            BlockType.PARAGRAPH: self._apply_inline,
            BlockType.BLANK: self._apply_blank,
            BlockType.FOOTNOTE_DEF: self._apply_footnote_def,
        }
        handler = handlers.get(block.type)
        if handler:
            handler(doc, block, config, tracker, ctx)

    def _apply_heading(self, doc, block, config, tracker, ctx):
        """剥离 # 标记，应用 Heading 样式"""
        idx = block.source_para_indices[0]
        para = doc.paragraphs[idx]
        old_text = para.text
        style_map = ctx.get("heading_style_map") or {}
        level_key = HEADING_LEVEL_KEY_MAP.get(block.level, "heading4")
        style_name = style_map.get(level_key, f"Heading {min(block.level, 4)}")

        write_spans(para, block.spans,
                    doc=doc,
                    footnote_defs=ctx.get("footnote_defs"),
                    base_font_name=ctx.get("base_font_name"),
                    base_font_size=ctx.get("base_font_size"))
        try:
            para.style = doc.styles[style_name]
        except KeyError:
            pass

        tracker.record(
            rule_name=self.name, target=f"段落 #{idx}",
            section="body", change_type="text",
            before=old_text[:80],
            after=f"→ {style_name}: {block.raw_text[:50]}",
            paragraph_index=idx)

    def _apply_code_block(self, doc, block, config, tracker, ctx):
        """围栏代码块：删除 ``` 行，内容段落应用等宽字体 + noProof"""
        indices = block.source_para_indices
        if len(indices) < 2:
            return

        code_sc = config.styles.get("code_block", StyleConfig(
            font_cn="Consolas", font_en="Consolas", size_pt=10,
            alignment="left", first_line_indent_chars=0,
            line_spacing_type="exact", line_spacing_pt=16))

        content_indices = indices[1:-1] if len(indices) > 2 else []
        fence_indices = [indices[0], indices[-1]]

        for ci in content_indices:
            if ci >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[ci]
            pf = para.paragraph_format
            pf.first_line_indent = Pt(0)
            pf.left_indent = Cm(1.0)
            for run in para.runs:
                run.font.name = code_sc.font_en
                run.font.size = Pt(code_sc.size_pt)
                set_no_proof(run)

        for fi in sorted(fence_indices, reverse=True):
            if fi >= len(doc.paragraphs):
                continue
            p = doc.paragraphs[fi]
            p._element.getparent().remove(p._element)

        tracker.record(
            rule_name=self.name,
            target=f"段落 #{indices[0]}-#{indices[-1]}",
            section="body", change_type="format",
            before=f"```{block.language}...```",
            after=f"→ code_block ({len(content_indices)} 行)",
            paragraph_index=indices[0])

    def _apply_table(self, doc, block, config, tracker, ctx):
        """Markdown 表格 → Word 原生表格（支持单元格换行和行内格式）"""
        indices = block.source_para_indices
        if not indices:
            return

        num_cols = len(block.table_headers)
        num_rows = 1 + len(block.table_rows)

        table = doc.add_table(rows=num_rows, cols=num_cols)

        for j, header in enumerate(block.table_headers):
            if j < num_cols:
                cell = table.rows[0].cells[j]
                write_table_cell(cell, header)
                for run in cell.paragraphs[0].runs:
                    run.bold = True

        for i, row_data in enumerate(block.table_rows):
            for j, cell_text in enumerate(row_data):
                if j < num_cols:
                    write_table_cell(table.rows[i + 1].cells[j], cell_text)

        # 应用列对齐
        align_map = {
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'left': WD_ALIGN_PARAGRAPH.LEFT,
        }
        if block.table_alignments:
            for row in table.rows:
                for j, cell in enumerate(row.cells):
                    if j < len(block.table_alignments):
                        align = align_map.get(block.table_alignments[j])
                        if align and cell.paragraphs:
                            cell.paragraphs[0].alignment = align

        # 移动表格到原位置
        first_para = doc.paragraphs[indices[0]]
        first_para._element.addprevious(table._tbl)

        # 删除原始 | 段落（倒序）
        for pi in sorted(indices, reverse=True):
            if pi >= len(doc.paragraphs):
                continue
            p = doc.paragraphs[pi]
            p._element.getparent().remove(p._element)

        tracker.record(
            rule_name=self.name,
            target=f"段落 #{indices[0]}-#{indices[-1]}",
            section="body", change_type="text",
            before=f"{len(indices)} 行 Markdown 表格",
            after=f"→ Word 表格 {num_rows}x{num_cols}",
            paragraph_index=indices[0])

    def _apply_blockquote(self, doc, block, config, tracker, ctx):
        """引用块：剥离 > 前缀，添加左边框+底色"""
        for idx in block.source_para_indices:
            if idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            old_text = para.text
            clean = re.sub(r'^(?:>\s*)+', '', para.text)
            self._replace_text(para, clean)
            apply_blockquote_border(para, block.quote_level)
            tracker.record(
                rule_name=self.name, target=f"段落 #{idx}",
                section="body", change_type="text",
                before=old_text[:60],
                after=f"→ blockquote (level {block.quote_level})",
                paragraph_index=idx)

    def _apply_blank(self, doc, block, config, tracker, ctx):
        """删除 Markdown 空行产生的空段落（保留含图片的段落）"""
        for idx in sorted(block.source_para_indices, reverse=True):
            if idx < len(doc.paragraphs):
                p = doc.paragraphs[idx]
                if not p.text.strip() and not self._para_has_content(p):
                    p._element.getparent().remove(p._element)

    def _apply_footnote_def(self, doc, block, config, tracker, ctx):
        """删除脚注定义源行，正文引用已在 spans 渲染阶段处理。"""
        indices = block.source_para_indices
        for idx in sorted(indices, reverse=True):
            if idx < len(doc.paragraphs):
                para = doc.paragraphs[idx]
                para._element.getparent().remove(para._element)

        if indices:
            tracker.record(
                rule_name=self.name,
                target=f"段落 #{indices[0]}",
                section="body",
                change_type="text",
                before=f"[^{block.list_marker}]: {block.raw_text[:40]}",
                after="(脚注定义已转为引用并删除定义行)",
                paragraph_index=indices[0],
            )

    @staticmethod
    def _para_has_content(para) -> bool:
        """检查段落是否包含非文本内容（图片、公式等）"""
        el = para._element
        W = f"{{{W_NS}}}"
        if el.findall(f"{W}r/{W}drawing"):
            return True
        if el.findall(f"{W}r/{W}pict"):
            return True
        return False

    def _apply_hr(self, doc, block, config, tracker, ctx):
        """水平线：删除 --- 段落"""
        idx = block.source_para_indices[0]
        if idx >= len(doc.paragraphs):
            return
        para = doc.paragraphs[idx]
        para._element.getparent().remove(para._element)
        tracker.record(
            rule_name=self.name, target=f"段落 #{idx}",
            section="body", change_type="text",
            before="---", after="(已删除)",
            paragraph_index=idx)

    def _apply_list_item(self, doc, block, config, tracker, ctx):
        """列表项：剥离标记，应用 Word 原生编号"""
        idx = block.source_para_indices[0]
        if idx >= len(doc.paragraphs):
            return
        para = doc.paragraphs[idx]
        old_text = para.text

        base_name, base_size = ctx.get("base_font_name"), ctx.get("base_font_size")
        write_spans(para, block.spans,
                    doc=doc,
                    footnote_defs=ctx.get("footnote_defs"),
                    base_font_name=base_name, base_font_size=base_size)
        is_ordered = block.list_marker and block.list_marker[-1] == '.'
        num_id = ctx["decimal_num_id"] if is_ordered else ctx["bullet_num_id"]
        apply_num_pr(para, num_id, block.list_level)

        tracker.record(
            rule_name=self.name, target=f"段落 #{idx}",
            section="body", change_type="text",
            before=old_text[:60],
            after=f"→ list (level {block.list_level}): {block.raw_text[:40]}",
            paragraph_index=idx)

    def _apply_task_list(self, doc, block, config, tracker, ctx):
        """任务清单：剥离 - [x] 标记，添加勾选符号"""
        idx = block.source_para_indices[0]
        if idx >= len(doc.paragraphs):
            return
        para = doc.paragraphs[idx]
        old_text = para.text
        prefix = "☑ " if block.checked else "☐ "
        # 清除现有 runs
        for run in para.runs:
            run._element.getparent().remove(run._element)
        para.add_run(prefix)
        write_spans(para, block.spans,
                    doc=doc,
                    footnote_defs=ctx.get("footnote_defs"),
                    clear_existing=False)
        tracker.record(
            rule_name=self.name, target=f"段落 #{idx}",
            section="body", change_type="text",
            before=old_text[:60],
            after=f"→ task: {prefix}{block.raw_text[:30]}",
            paragraph_index=idx)

    def _apply_inline(self, doc, block, config, tracker, ctx):
        """普通段落：处理行内 Markdown 格式 + 多段落合并"""
        indices = block.source_para_indices
        idx = indices[0]
        if idx >= len(doc.paragraphs):
            return

        base_name, base_size = ctx.get("base_font_name"), ctx.get("base_font_size")
        # 关键修正：
        # 多个 source 段落时不再压缩成“单段 + 手动换行(w:br)”，
        # 直接逐段渲染，保留真正段落边界（Enter）。
        if len(indices) > 1:
            for para_idx in indices:
                if para_idx >= len(doc.paragraphs):
                    continue
                para = doc.paragraphs[para_idx]
                spans = parse_inline(para.text or "")
                write_spans(
                    para, spans,
                    doc=doc,
                    footnote_defs=ctx.get("footnote_defs"),
                    base_font_name=base_name, base_font_size=base_size)
            after_desc = f"→ inline formatted ({len(indices)} para preserved)"
            before_text = doc.paragraphs[idx].text[:60]
        else:
            para = doc.paragraphs[idx]
            old_text = para.text
            write_spans(para, block.spans,
                        doc=doc,
                        footnote_defs=ctx.get("footnote_defs"),
                        base_font_name=base_name, base_font_size=base_size)
            before_text = old_text[:60]
            after_desc = "→ inline formatted (1 para)"

        tracker.record(
            rule_name=self.name, target=f"段落 #{idx}",
            section="body", change_type="format",
            before=before_text,
            after=after_desc,
            paragraph_index=idx)

    # ── 辅助方法 ──

    @staticmethod
    def _replace_text(para, new_text: str):
        """替换段落文本，保留第一个 run 的格式"""
        if not para.runs:
            para.text = new_text
            return
        first_run = para.runs[0]
        for run in para.runs[1:]:
            run.text = ""
        first_run.text = new_text

    @staticmethod
    def _get_base_font(para):
        """从段落第一个 run 提取基础字体信息"""
        if para.runs:
            r = para.runs[0]
            return r.font.name, r.font.size
        return None, None

    @staticmethod
    def _find_keyword_index(doc: Document, keyword: str) -> int | None:
        """在文档中查找包含关键字的段落索引"""
        kw = re.sub(r'\s+', '', keyword or "")
        if not kw:
            return None
        for i, para in enumerate(doc.paragraphs):
            text = re.sub(r'\s+', '', (para.text or "").strip())
            if kw in text:
                return i
        return None

    @staticmethod
    def _find_page_start_index(doc: Document, target_page: int) -> int | None:
        """根据显式分页符/分节符估算目标页起始段落（1-based page）。"""
        if target_page is None or target_page < 1:
            return None
        if target_page == 1:
            return 0

        current_page = 1
        para_idx = -1

        for child in doc.element.body:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag != "p":
                continue
            para_idx += 1

            for br in child.iter(f"{{{W_NS}}}br"):
                if br.get(f"{{{W_NS}}}type", "") == "page":
                    current_page += 1
                    if current_page == target_page:
                        return para_idx + 1

            ppr = child.find(f"{{{W_NS}}}pPr")
            if ppr is None:
                continue
            sect = ppr.find(f"{{{W_NS}}}sectPr")
            if sect is None:
                continue
            sect_type_el = sect.find(f"{{{W_NS}}}type")
            sect_type = ""
            if sect_type_el is not None:
                sect_type = sect_type_el.get(f"{{{W_NS}}}val", "")
            if sect_type in ("nextPage", "oddPage", "evenPage", ""):
                current_page += 1
                if current_page == target_page:
                    return para_idx + 1

        return None

    def _split_manual_break_paragraphs(
            self, doc: Document, start_idx: int, end_idx: int) -> int:
        """将段落中的 '\\n' 手动换行拆分为多个真实段落。"""
        if start_idx is None or end_idx is None:
            return 0
        if start_idx < 0 or end_idx < start_idx:
            return 0

        inserted = 0
        end_idx = min(end_idx, len(doc.paragraphs) - 1)
        for idx in range(end_idx, start_idx - 1, -1):
            if idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            if self._para_has_content(para):
                continue
            text = para.text or ""
            if "\n" not in text and "\r" not in text:
                continue

            parts = re.split(r'\r?\n', text)
            if len(parts) <= 1:
                continue

            src_style = para.style
            src_pf = para.paragraph_format
            self._replace_text(para, parts[0])

            anchor = para
            for part in parts[1:]:
                new_para = self._insert_paragraph_after(anchor, part)
                try:
                    new_para.style = src_style
                except Exception:
                    pass
                new_pf = new_para.paragraph_format
                new_pf.alignment = src_pf.alignment
                new_pf.first_line_indent = src_pf.first_line_indent
                new_pf.left_indent = src_pf.left_indent
                new_pf.right_indent = src_pf.right_indent
                new_pf.space_before = src_pf.space_before
                new_pf.space_after = src_pf.space_after
                new_pf.line_spacing = src_pf.line_spacing
                new_pf.line_spacing_rule = src_pf.line_spacing_rule
                anchor = new_para
                inserted += 1

        return inserted

    @staticmethod
    def _insert_paragraph_after(para, text: str = ""):
        """在指定段落后插入新段落。"""
        new_p = OxmlElement("w:p")
        para._element.addnext(new_p)
        new_para = Paragraph(new_p, para._parent)
        if text:
            new_para.add_run(text)
        return new_para
