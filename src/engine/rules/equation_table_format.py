"""公式表格识别与调整规则（实验室可选）"""

from docx import Document

from src.engine.change_tracker import ChangeTracker
from src.engine.rules import table_format as table_helpers
from src.engine.rules.base import BaseRule
from src.scene.schema import SceneConfig


class EquationTableFormatRule(BaseRule):
    name = "equation_table_format"
    description = "公式表格识别与调整"

    def apply(self, doc: Document, config: SceneConfig,
              tracker: ChangeTracker, context: dict) -> None:
        avail_w = table_helpers._available_width(config)
        body_range = None

        doc_tree = context.get("doc_tree")
        if doc_tree:
            body_sec = doc_tree.get_section("body")
            if body_sec:
                start, end = body_sec.start_index, body_sec.end_index
                if end >= start:
                    body_range = (start, end)

        headings = context.get("headings", [])
        body_end = body_range[1] if body_range else (len(doc.paragraphs) - 1)
        chapter_ranges = table_helpers._build_chapter_ranges(headings, body_end)
        chapter_eq_counts = {}

        body_el = doc.element.body
        tbl_para_pos = []
        para_idx = -1
        for child in body_el:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                para_idx += 1
            elif tag == "tbl":
                tbl_para_pos.append(para_idx if para_idx >= 0 else 0)

        eq_table_count = 0
        eq_autonum_count = 0
        eq_renumber_count = 0

        for tbl_idx, tbl in enumerate(doc.tables):
            tbl_el = tbl._tbl
            pos = tbl_para_pos[tbl_idx] if tbl_idx < len(tbl_para_pos) else -1

            if body_range is not None:
                if pos >= 0 and (pos < body_range[0] or pos > body_range[1]):
                    continue

            if not table_helpers._is_equation_table(tbl_el):
                continue

            chap = table_helpers._get_chapter_num(pos, chapter_ranges)
            if chap <= 0:
                chap = 1

            eq_rows_in_table = 0
            for row_idx, tr in enumerate(tbl_el.findall(table_helpers._w("tr"))):
                if not table_helpers._is_equation_row(tr):
                    continue

                eq_rows_in_table += 1
                current_seq = chapter_eq_counts.get(chap, 0)
                expected_seq = current_seq + 1
                expected_number = f"({chap}.{expected_seq})"
                existing_id, pure_cell, number_cell = (
                    table_helpers._find_existing_equation_number_in_row(tr)
                )

                chapter_eq_counts[chap] = expected_seq

                if existing_id:
                    parsed = table_helpers._parse_chapter_seq(existing_id)
                    if parsed == (chap, expected_seq):
                        continue

                    before_number = f"({existing_id})"
                    updated_cell = table_helpers._rewrite_existing_equation_number_in_row(
                        tr,
                        expected_number,
                        pure_cell=pure_cell,
                        number_cell=number_cell,
                    )
                    if updated_cell is None:
                        updated_cell = table_helpers._insert_equation_number_cell_in_row(
                            tr, expected_number
                        )
                    if updated_cell is None:
                        continue

                    if table_helpers._RE_EQ_NUM.match(
                        (table_helpers._get_cell_text(updated_cell) or "").strip()
                    ):
                        table_helpers._format_equation_number_cell(updated_cell)
                    eq_renumber_count += 1
                    tracker.record(
                        rule_name=self.name,
                        target=f"公式表格 #{tbl_idx} 第{row_idx + 1}行",
                        section="body",
                        change_type="numbering",
                        before=before_number,
                        after=expected_number,
                        paragraph_index=pos,
                    )
                    continue

                inserted_cell = table_helpers._insert_equation_number_cell_in_row(
                    tr, expected_number
                )
                if inserted_cell is None:
                    continue

                table_helpers._format_equation_number_cell(inserted_cell)
                eq_autonum_count += 1
                tracker.record(
                    rule_name=self.name,
                    target=f"公式表格 #{tbl_idx} 第{row_idx + 1}行",
                    section="body",
                    change_type="numbering",
                    before="(无序号)",
                    after=expected_number,
                        paragraph_index=pos,
                    )

            # 防止误判：若没有任何“公式行”，不对该表应用公式表格样式
            # （否则可能把普通数据表误改为无边框样式）。
            if eq_rows_in_table <= 0:
                continue

            eq_table_count += 1
            table_helpers._format_equation_table(tbl_el, avail_w)

        if eq_table_count:
            tracker.record(
                rule_name=self.name,
                target=f"{eq_table_count} 个公式表格",
                section="body",
                change_type="format",
                before="默认表格样式",
                after="已应用公式表格识别与专用格式",
                paragraph_index=-1,
            )
        if eq_autonum_count:
            tracker.record(
                rule_name=self.name,
                target=f"{eq_autonum_count} 个公式行",
                section="body",
                change_type="numbering",
                before="(缺失序号)",
                after="已自动补全并应用右侧编号格式",
                paragraph_index=-1,
            )
        if eq_renumber_count:
            tracker.record(
                rule_name=self.name,
                target=f"{eq_renumber_count} 个公式行",
                section="body",
                change_type="numbering",
                before="(序号不连续/错误)",
                after="已按顺序自动纠正并递推后续序号",
                paragraph_index=-1,
            )
