"""Validation rule: check whether formatted result matches scene constraints."""

from docx import Document

from src.engine.rules.base import BaseRule, ValidationIssue
from src.engine.change_tracker import ChangeTracker
from src.scene.heading_model import get_level_to_word_style
from src.scene.schema import SceneConfig
from src.utils.toc_entry import (
    looks_like_numbered_toc_entry_with_page_suffix,
    looks_like_toc_entry_line,
)


_BROKEN_TOC_BOOKMARK_HINTS = (
    "error! bookmark not defined",
    "bookmark not defined",
    "\u672a\u5b9a\u4e49\u4e66\u7b7e",
)


class ValidationRule(BaseRule):
    name = "validation"
    description = "校验排版结果"

    def apply(
        self,
        doc: Document,
        config: SceneConfig,
        tracker: ChangeTracker,
        context: dict,
    ) -> None:
        issues = self.validate(doc, config, context)
        context["validation_issues"] = issues
        for issue in issues:
            tracker.record(
                rule_name=self.name,
                target=issue.location,
                section="global",
                change_type="format",
                before=issue.message,
                after=f"[{issue.level}]",
                paragraph_index=-1,
                success=(issue.level != "error"),
                failure_reason=issue.message if issue.level == "error" else None,
            )

    def validate(
        self,
        doc: Document,
        config: SceneConfig,
        context: dict | None = None,
    ) -> list[ValidationIssue]:
        issues: list[ValidationIssue] = []
        ctx = context or {}
        issues.extend(self._check_page_setup(doc, config))
        issues.extend(self._check_heading_styles(doc, config, ctx))
        issues.extend(self._check_separator_violations(doc, config, ctx))
        return issues

    def _check_page_setup(self, doc, config):
        issues: list[ValidationIssue] = []
        ps = config.page_setup
        for i, section in enumerate(doc.sections):
            expected_top = round(ps.margin.top_cm * 360000)
            actual_top = section.top_margin
            if actual_top and abs(actual_top - expected_top) > 5000:
                issues.append(
                    ValidationIssue(
                        level="warning",
                        rule_name=self.name,
                        message=f"页边距(上)不符: 期望 {ps.margin.top_cm}cm",
                        location=f"Section {i + 1}",
                    )
                )
        return issues

    def _check_heading_styles(self, doc, config, context):
        issues: list[ValidationIssue] = []
        headings = context.get("headings", [])
        level_style = get_level_to_word_style(config)
        for h in headings:
            para = doc.paragraphs[h.para_index]
            expected = level_style.get(h.level, "")
            actual = para.style.name if para.style else ""
            if actual != expected:
                issues.append(
                    ValidationIssue(
                        level="warning",
                        rule_name=self.name,
                        message=f"标题样式不符: 期望 '{expected}', 实际 '{actual}'",
                        location=f"段落 #{h.para_index}",
                    )
                )
        return issues

    def _is_toc_placeholder_heading(self, text: str) -> bool:
        raw = (text or "").strip()
        if "\t" not in raw:
            return False
        low = raw.lower()
        if looks_like_toc_entry_line(raw) or looks_like_numbered_toc_entry_with_page_suffix(raw):
            return True
        if any(hint in low for hint in _BROKEN_TOC_BOOKMARK_HINTS):
            return True
        return ("\u4e66\u7b7e" in low and "\u672a\u5b9a\u4e49" in low)

    def _check_separator_violations(self, doc, config, context):
        issues: list[ValidationIssue] = []
        headings = context.get("headings", [])
        enforcement = config.heading_numbering.enforcement
        for h in headings:
            text = doc.paragraphs[h.para_index].text or ""
            if self._is_toc_placeholder_heading(text):
                continue
            if enforcement.ban_tab and "\t" in text:
                issues.append(
                    ValidationIssue(
                        level="error",
                        rule_name=self.name,
                        message="标题中仍存在 Tab 字符",
                        location=f"段落 #{h.para_index}",
                    )
                )
            if enforcement.ban_double_halfwidth_space and "  " in text:
                issues.append(
                    ValidationIssue(
                        level="error",
                        rule_name=self.name,
                        message="标题中仍存在连续两个半角空格",
                        location=f"段落 #{h.para_index}",
                    )
                )
        return issues
