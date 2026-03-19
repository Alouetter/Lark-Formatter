from __future__ import annotations

import argparse
import json
import subprocess
import sys
from dataclasses import asdict
from hashlib import sha256
from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as _DocumentType
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.scene.manager import load_scene_from_data
from src.utils.heading_numbering_v2 import legacy_levels_from_v2


def iter_block_items(parent: _DocumentType | _Cell):
    if isinstance(parent, _DocumentType):
        parent_elm = parent.element.body
    else:
        parent_elm = parent._tc
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def normalize_text(text: str) -> str:
    return " ".join((text or "").replace("\r", "\n").split())


def cell_text(cell: _Cell) -> str:
    lines: list[str] = []
    for block in iter_block_items(cell):
        if isinstance(block, Paragraph):
            text = normalize_text(block.text)
            if text:
                lines.append(text)
        elif isinstance(block, Table):
            nested = table_lines(block)
            if nested:
                lines.extend(nested)
    return " / ".join(lines)


def table_lines(table: Table) -> list[str]:
    lines: list[str] = []
    for row in table.rows:
        cells = [cell_text(cell) for cell in row.cells]
        if any(cells):
            lines.append(" | ".join(cells))
    return lines


def extract_docx_text(path: Path) -> str:
    doc = Document(path)
    lines: list[str] = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = normalize_text(block.text)
            if text:
                lines.append(text)
        elif isinstance(block, Table):
            rows = table_lines(block)
            if rows:
                lines.extend(rows)
    return "\n".join(lines)


def bool_literal(value: bool) -> str:
    return "true" if value else "false"


def build_prompt(
    prompt_template: str,
    baseline_text: str,
    spec_text: str,
    baseline_data: dict[str, Any],
) -> str:
    replacements = {
        "{{SCENE_NAME}}": "南开大学研究生学位论文写作规范（2026版）",
        "{{SCENE_DESCRIPTION}}": "根据《南开大学研究生学位论文写作规范（2026版）》生成的场景配置回归测试",
        "{{CATEGORY}}": str(baseline_data.get("category", "thesis")),
        "{{CATEGORY_LABEL}}": str(baseline_data.get("category_label", "学位论文")),
        "{{BASELINE_JSON}}": baseline_text,
        "{{UPDATE_HEADER}}": bool_literal(bool(baseline_data.get("update_header", True))),
        "{{UPDATE_PAGE_NUMBER}}": bool_literal(bool(baseline_data.get("update_page_number", True))),
        "{{UPDATE_HEADER_LINE}}": bool_literal(bool(baseline_data.get("update_header_line", True))),
        "{{ENABLE_WHITESPACE_NORMALIZE}}": bool_literal(
            bool((baseline_data.get("whitespace_normalize") or {}).get("enabled", False))
        ),
        "{{ENABLE_FORMULA_CONVERT}}": bool_literal(
            bool((baseline_data.get("formula_convert") or {}).get("enabled", False))
        ),
        "{{ENABLE_FORMULA_TO_TABLE}}": bool_literal(
            bool((baseline_data.get("formula_to_table") or {}).get("enabled", False))
        ),
        "{{ENABLE_EQUATION_TABLE_FORMAT}}": bool_literal(
            bool((baseline_data.get("equation_table_format") or {}).get("enabled", False))
        ),
        "{{ENABLE_FORMULA_STYLE}}": bool_literal(
            bool((baseline_data.get("formula_style") or {}).get("enabled", False))
        ),
        "{{SPEC_TEXT}}": spec_text,
    }
    prompt_text = prompt_template
    for placeholder, value in replacements.items():
        prompt_text = prompt_text.replace(placeholder, value)
    return prompt_text


def canonical_json(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def display_from_size(size_pt: Any) -> str:
    if not isinstance(size_pt, (int, float)):
        return ""
    numeric = float(size_pt)
    if numeric.is_integer():
        return str(int(numeric))
    return f"{numeric:.4f}".rstrip("0").rstrip(".")


def approx_equal(left: Any, right: Any, tol: float = 1e-6) -> bool:
    try:
        return abs(float(left) - float(right)) <= tol
    except Exception:
        return False


def compare_values(
    path: str,
    actual: Any,
    expected: Any,
    issues: list[str],
    *,
    tol: float = 1e-6,
) -> None:
    if isinstance(expected, float):
        if not approx_equal(actual, expected, tol=tol):
            issues.append(f"{path}: expected {expected}, got {actual!r}")
        return
    if actual != expected:
        issues.append(f"{path}: expected {expected!r}, got {actual!r}")


def validate_heading_projection(cfg: Any, issues: list[str]) -> None:
    # Mirror runtime behavior: when payloads provide both the legacy
    # heading_numbering structure and heading_numbering_v2, runtime keeps
    # legacy-only layout fields (for example alignment / indent) and only
    # reprojects numbering-related fields from v2.
    projected = legacy_levels_from_v2(
        cfg.heading_numbering_v2,
        existing_levels=cfg.heading_numbering.levels,
    )
    for level_name, projected_level in projected.items():
        legacy_level = cfg.heading_numbering.levels.get(level_name)
        if legacy_level is None:
            issues.append(f"heading_numbering.levels.{level_name}: missing")
            continue
        projected_data = asdict(projected_level)
        legacy_data = asdict(legacy_level)
        if projected_data != legacy_data:
            issues.append(
                f"heading_numbering sync mismatch at {level_name}: "
                f"expected {projected_data}, got {legacy_data}"
            )


def validate_pipeline(cfg: Any, issues: list[str]) -> None:
    pipeline = list(cfg.pipeline or [])
    if len(pipeline) != len(set(pipeline)):
        issues.append("pipeline: contains duplicated steps")

    whitespace_enabled = bool(getattr(cfg.whitespace_normalize, "enabled", False))
    whitespace_present = "whitespace_normalize" in pipeline
    if whitespace_enabled != whitespace_present:
        issues.append(
            f"pipeline/whitespace_normalize mismatch: enabled={whitespace_enabled}, present={whitespace_present}"
        )
    if whitespace_enabled and whitespace_present:
        if pipeline.index("whitespace_normalize") <= pipeline.index("md_cleanup"):
            issues.append("pipeline: whitespace_normalize must be after md_cleanup")

    formula_enabled = any(
        bool(getattr(getattr(cfg, key), "enabled", False))
        for key in ("formula_convert", "formula_to_table", "equation_table_format", "formula_style")
    )
    formula_steps = ["formula_convert", "formula_to_table", "equation_table_format", "formula_style"]
    present_formula_steps = [step for step in formula_steps if step in pipeline]
    if formula_enabled:
        if present_formula_steps != formula_steps:
            issues.append(
                f"pipeline: formula cluster mismatch, expected {formula_steps}, got {present_formula_steps}"
            )
        else:
            indexes = [pipeline.index(step) for step in formula_steps]
            if indexes != sorted(indexes):
                issues.append("pipeline: formula cluster order is incorrect")
            if max(indexes) - min(indexes) != len(formula_steps) - 1:
                issues.append("pipeline: formula cluster must be contiguous")
            if "table_format" in pipeline and pipeline.index("formula_convert") <= pipeline.index("table_format"):
                issues.append("pipeline: formula cluster must be after table_format")
            if "section_format" in pipeline and pipeline.index("formula_style") >= pipeline.index("section_format"):
                issues.append("pipeline: formula cluster must be before section_format")


def validate_anchor_values(data: dict[str, Any], cfg: Any, issues: list[str]) -> None:
    compare_values("category", cfg.category, "thesis", issues)
    compare_values("category_label", cfg.category_label, "学位论文", issues)

    compare_values("page_setup.paper_size", cfg.page_setup.paper_size, "A4", issues)
    compare_values("page_setup.margin.top_cm", cfg.page_setup.margin.top_cm, 3.8, issues)
    compare_values("page_setup.margin.bottom_cm", cfg.page_setup.margin.bottom_cm, 3.8, issues)
    compare_values("page_setup.margin.left_cm", cfg.page_setup.margin.left_cm, 3.2, issues)
    compare_values("page_setup.margin.right_cm", cfg.page_setup.margin.right_cm, 3.2, issues)
    compare_values("page_setup.header_distance_cm", cfg.page_setup.header_distance_cm, 3.0, issues)
    compare_values("page_setup.footer_distance_cm", cfg.page_setup.footer_distance_cm, 3.0, issues)
    compare_values("page_setup.gutter_cm", cfg.page_setup.gutter_cm, 0.0, issues)

    compare_values("update_header", cfg.update_header, True, issues)
    compare_values("update_page_number", cfg.update_page_number, True, issues)
    compare_values("update_header_line", cfg.update_header_line, True, issues)

    styles = cfg.styles
    compare_values("styles.normal.font_cn", styles["normal"].font_cn, "宋体", issues)
    compare_values("styles.normal.font_en", styles["normal"].font_en, "Times New Roman", issues)
    compare_values("styles.normal.size_pt", styles["normal"].size_pt, 12.0, issues)
    compare_values("styles.normal.alignment", styles["normal"].alignment, "justify", issues)
    compare_values("styles.normal.first_line_indent_chars", styles["normal"].first_line_indent_chars, 2.0, issues)
    compare_values("styles.normal.first_line_indent_unit", styles["normal"].first_line_indent_unit, "chars", issues)
    compare_values("styles.normal.line_spacing_type", styles["normal"].line_spacing_type, "exact", issues)
    compare_values("styles.normal.line_spacing_pt", styles["normal"].line_spacing_pt, 20.0, issues)

    heading_expectations = {
        "heading1": {
            "alignment": "center",
            "left_indent_chars": 0.0,
            "size_pt": 16.0,
            "bold": True,
            "space_before_pt": 24.0,
            "space_after_pt": 18.0,
        },
        "heading2": {
            "alignment": "center",
            "left_indent_chars": 0.0,
            "size_pt": 14.0,
            "bold": True,
            "space_before_pt": 24.0,
            "space_after_pt": 6.0,
        },
        "heading3": {
            "alignment": "left",
            "left_indent_chars": 2.0,
            "size_pt": 13.0,
            "bold": False,
            "space_before_pt": 12.0,
            "space_after_pt": 6.0,
        },
        "heading4": {
            "alignment": "left",
            "left_indent_chars": 2.0,
            "size_pt": 12.0,
            "bold": False,
            "space_before_pt": 12.0,
            "space_after_pt": 6.0,
        },
    }
    for level_name, expected in heading_expectations.items():
        style = styles[level_name]
        compare_values(f"styles.{level_name}.font_cn", style.font_cn, "黑体", issues)
        compare_values(f"styles.{level_name}.line_spacing_type", style.line_spacing_type, "single", issues)
        for field_name, value in expected.items():
            compare_values(f"styles.{level_name}.{field_name}", getattr(style, field_name), value, issues)

    legacy_heading_expectations = {
        "heading1": {"alignment": "center", "left_indent_chars": 0},
        "heading2": {"alignment": "center", "left_indent_chars": 0},
        "heading3": {"alignment": "left", "left_indent_chars": 2},
        "heading4": {"alignment": "left", "left_indent_chars": 2},
    }
    legacy_levels = ((data.get("heading_numbering") or {}).get("levels") or {})
    legacy_scheme_2 = (((data.get("heading_numbering") or {}).get("schemes") or {}).get("2") or {})
    for level_name, expected in legacy_heading_expectations.items():
        level_cfg = legacy_levels.get(level_name) or {}
        compare_values(
            f"heading_numbering.levels.{level_name}.alignment",
            level_cfg.get("alignment"),
            expected["alignment"],
            issues,
        )
        compare_values(
            f"heading_numbering.levels.{level_name}.left_indent_chars",
            level_cfg.get("left_indent_chars"),
            expected["left_indent_chars"],
            issues,
        )
        scheme_level_cfg = legacy_scheme_2.get(level_name) or {}
        compare_values(
            f"heading_numbering.schemes.2.{level_name}.alignment",
            scheme_level_cfg.get("alignment"),
            expected["alignment"],
            issues,
        )
        compare_values(
            f"heading_numbering.schemes.2.{level_name}.left_indent_chars",
            scheme_level_cfg.get("left_indent_chars"),
            expected["left_indent_chars"],
            issues,
        )

    compare_values("styles.abstract_title_cn.font_cn", styles["abstract_title_cn"].font_cn, "黑体", issues)
    compare_values("styles.abstract_title_cn.size_pt", styles["abstract_title_cn"].size_pt, 18.0, issues)
    compare_values("styles.abstract_title_cn.bold", styles["abstract_title_cn"].bold, True, issues)
    compare_values("styles.abstract_title_cn.alignment", styles["abstract_title_cn"].alignment, "center", issues)
    compare_values("styles.abstract_title_en.font_en", styles["abstract_title_en"].font_en, "Arial", issues)
    compare_values("styles.abstract_title_en.size_pt", styles["abstract_title_en"].size_pt, 18.0, issues)
    compare_values("styles.abstract_title_en.bold", styles["abstract_title_en"].bold, True, issues)
    compare_values("styles.abstract_title_en.alignment", styles["abstract_title_en"].alignment, "center", issues)

    compare_values("styles.toc_title.font_cn", styles["toc_title"].font_cn, "黑体", issues)
    compare_values("styles.toc_title.size_pt", styles["toc_title"].size_pt, 16.0, issues)
    compare_values("styles.toc_title.bold", styles["toc_title"].bold, True, issues)
    compare_values("styles.toc_title.alignment", styles["toc_title"].alignment, "center", issues)
    compare_values("styles.toc_chapter.size_pt", styles["toc_chapter"].size_pt, 14.0, issues)
    compare_values("styles.toc_level1.size_pt", styles["toc_level1"].size_pt, 12.0, issues)
    compare_values("styles.toc_level1.left_indent_chars", styles["toc_level1"].left_indent_chars, 1.0, issues)
    compare_values("styles.toc_level2.size_pt", styles["toc_level2"].size_pt, 10.5, issues)
    compare_values("styles.toc_level2.left_indent_chars", styles["toc_level2"].left_indent_chars, 2.0, issues)

    compare_values("styles.figure_caption.size_pt", styles["figure_caption"].size_pt, 10.5, issues)
    compare_values("styles.figure_caption.alignment", styles["figure_caption"].alignment, "center", issues)
    compare_values("styles.figure_caption.space_before_pt", styles["figure_caption"].space_before_pt, 6.0, issues)
    compare_values("styles.figure_caption.space_after_pt", styles["figure_caption"].space_after_pt, 12.0, issues)
    compare_values("styles.table_caption.size_pt", styles["table_caption"].size_pt, 10.5, issues)
    compare_values("styles.table_caption.alignment", styles["table_caption"].alignment, "center", issues)
    compare_values("styles.table_caption.space_before_pt", styles["table_caption"].space_before_pt, 6.0, issues)
    compare_values("styles.table_caption.space_after_pt", styles["table_caption"].space_after_pt, 6.0, issues)

    compare_values("styles.references_body.size_pt", styles["references_body"].size_pt, 10.5, issues)
    compare_values("styles.references_body.line_spacing_type", styles["references_body"].line_spacing_type, "exact", issues)
    compare_values("styles.references_body.line_spacing_pt", styles["references_body"].line_spacing_pt, 16.0, issues)
    compare_values("styles.acknowledgment_body.font_cn", styles["acknowledgment_body"].font_cn, "仿宋", issues)
    compare_values("styles.acknowledgment_body.size_pt", styles["acknowledgment_body"].size_pt, 12.0, issues)
    compare_values("styles.acknowledgment_body.line_spacing_pt", styles["acknowledgment_body"].line_spacing_pt, 16.0, issues)
    compare_values("styles.resume_body.size_pt", styles["resume_body"].size_pt, 10.5, issues)
    compare_values("styles.resume_body.line_spacing_pt", styles["resume_body"].line_spacing_pt, 16.0, issues)
    compare_values("styles.symbol_table_body.size_pt", styles["symbol_table_body"].size_pt, 10.5, issues)
    compare_values("styles.symbol_table_body.line_spacing_pt", styles["symbol_table_body"].line_spacing_pt, 16.0, issues)
    compare_values("styles.appendix_body.size_pt", styles["appendix_body"].size_pt, 12.0, issues)
    compare_values("styles.appendix_body.first_line_indent_chars", styles["appendix_body"].first_line_indent_chars, 2.0, issues)
    compare_values("styles.appendix_body.line_spacing_pt", styles["appendix_body"].line_spacing_pt, 20.0, issues)
    compare_values("styles.header_en.font_cn", styles["header_en"].font_cn, "Times New Roman", issues)
    compare_values("styles.header_en.font_en", styles["header_en"].font_en, "Times New Roman", issues)

    compare_values("caption.figure_prefix", cfg.caption.figure_prefix, "图", issues)
    compare_values("caption.table_prefix", cfg.caption.table_prefix, "表", issues)
    compare_values("caption.separator", cfg.caption.separator, "\u3000", issues)

    compare_values("normal_table_border_mode", cfg.normal_table_border_mode, "three_line", issues)
    compare_values("normal_table_repeat_header", cfg.normal_table_repeat_header, True, issues)
    compare_values("table_border_width_pt", cfg.table_border_width_pt, 0.5, issues)
    compare_values("three_line_header_width_pt", cfg.three_line_header_width_pt, 1.0, issues)
    compare_values("three_line_bottom_width_pt", cfg.three_line_bottom_width_pt, 0.5, issues)

    v2 = cfg.heading_numbering_v2.level_bindings
    compare_values("heading_numbering_v2.heading1.display_shell", v2["heading1"].display_shell, "chapter_cn", issues)
    compare_values("heading_numbering_v2.heading2.display_shell", v2["heading2"].display_shell, "section_cn", issues)
    compare_values("heading_numbering_v2.heading3.display_shell", v2["heading3"].display_shell, "dunhao_cn", issues)
    compare_values("heading_numbering_v2.heading4.display_shell", v2["heading4"].display_shell, "paren_cn", issues)
    for level_name in ("heading1", "heading2", "heading3", "heading4"):
        compare_values(
            f"heading_numbering_v2.{level_name}.title_separator",
            v2[level_name].title_separator,
            "\u3000",
            issues,
        )


def flatten_json(data: Any, prefix: str = "") -> dict[str, str]:
    items: dict[str, str] = {}
    if isinstance(data, dict):
        for key in sorted(data):
            path = f"{prefix}.{key}" if prefix else str(key)
            items.update(flatten_json(data[key], path))
    elif isinstance(data, list):
        for idx, value in enumerate(data):
            path = f"{prefix}[{idx}]"
            items.update(flatten_json(value, path))
    else:
        items[prefix] = canonical_json(data)
    return items


def diff_paths(left: Any, right: Any) -> list[str]:
    left_flat = flatten_json(left)
    right_flat = flatten_json(right)
    paths = set(left_flat) | set(right_flat)
    return sorted(path for path in paths if left_flat.get(path) != right_flat.get(path))


def validate_round(
    raw_text: str,
    baseline_data: dict[str, Any],
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "ok": False,
        "issues": [],
        "hash": None,
        "data": None,
    }
    stripped = raw_text.lstrip("\ufeff").strip()
    try:
        data = json.loads(stripped)
    except Exception as exc:
        result["issues"].append(f"json parse failed: {exc}")
        return result

    baseline_keys = set(baseline_data.keys())
    data_keys = set(data.keys())
    if data_keys != baseline_keys:
        missing = sorted(baseline_keys - data_keys)
        extra = sorted(data_keys - baseline_keys)
        if missing:
            result["issues"].append(f"top-level missing keys: {missing}")
        if extra:
            result["issues"].append(f"top-level extra keys: {extra}")

    if "format_file" in data:
        result["issues"].append("top-level unexpected key: format_file")

    baseline_style_keys = set((baseline_data.get("styles") or {}).keys())
    data_style_keys = set((data.get("styles") or {}).keys())
    if data_style_keys != baseline_style_keys:
        missing = sorted(baseline_style_keys - data_style_keys)
        extra = sorted(data_style_keys - baseline_style_keys)
        if missing:
            result["issues"].append(f"styles missing keys: {missing}")
        if extra:
            result["issues"].append(f"styles extra keys: {extra}")

    style_field_keys = set((baseline_data.get("styles") or {}).get("normal", {}).keys())
    for style_name, style_data in sorted((data.get("styles") or {}).items()):
        if not isinstance(style_data, dict):
            result["issues"].append(f"styles.{style_name}: expected object, got {type(style_data).__name__}")
            continue
        actual_keys = set(style_data.keys())
        if actual_keys != style_field_keys:
            missing = sorted(style_field_keys - actual_keys)
            extra = sorted(actual_keys - style_field_keys)
            if missing:
                result["issues"].append(f"styles.{style_name}: missing fields {missing}")
            if extra:
                result["issues"].append(f"styles.{style_name}: extra fields {extra}")
        if "size_pt" in style_data and "size_display" in style_data:
            expected_display = display_from_size(style_data["size_pt"])
            if expected_display and style_data["size_display"] != expected_display:
                result["issues"].append(
                    f"styles.{style_name}.size_display: expected {expected_display!r}, got {style_data['size_display']!r}"
                )

    try:
        cfg = load_scene_from_data(data)
    except Exception as exc:
        result["issues"].append(f"scene parse failed: {exc}")
        return result

    if str(getattr(cfg, "_heading_numbering_v2_source", "")) != "payload":
        result["issues"].append(
            f"heading_numbering_v2 source mismatch: {getattr(cfg, '_heading_numbering_v2_source', None)!r}"
        )

    validate_heading_projection(cfg, result["issues"])
    validate_pipeline(cfg, result["issues"])
    validate_anchor_values(data, cfg, result["issues"])

    result["ok"] = not result["issues"]
    result["data"] = data
    result["hash"] = sha256(canonical_json(data).encode("utf-8")).hexdigest()
    return result


def run_codex_round(
    prompt_text: str,
    out_dir: Path,
    round_idx: int,
    *,
    model: str | None,
) -> tuple[int, str, str, Path]:
    raw_path = out_dir / f"round_{round_idx:02d}_last_message.txt"
    cmd = [
        "codex",
        "exec",
        "-C",
        str(ROOT),
        "--skip-git-repo-check",
        "--ephemeral",
        "--color",
        "never",
        "-s",
        "danger-full-access",
        "-o",
        str(raw_path),
    ]
    if model:
        cmd.extend(["-m", model])
    cmd.append("-")
    completed = subprocess.run(
        cmd,
        input=prompt_text,
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(ROOT),
    )
    stdout = completed.stdout or ""
    stderr = completed.stderr or ""
    return completed.returncode, stdout, stderr, raw_path


def summarize_variations(successful_runs: list[dict[str, Any]]) -> dict[str, list[str]]:
    if not successful_runs:
        return {}
    values_by_path: dict[str, set[str]] = {}
    for run in successful_runs:
        for path, value in flatten_json(run["data"]).items():
            values_by_path.setdefault(path, set()).add(value)
    return {
        path: sorted(values)
        for path, values in sorted(values_by_path.items())
        if len(values) > 1
    }


def build_report(
    prompt_file: Path,
    spec_docx: Path,
    rounds: int,
    round_results: list[dict[str, Any]],
) -> dict[str, Any]:
    successful = [item for item in round_results if item.get("ok")]
    exact_hashes = sorted({item["hash"] for item in successful if item.get("hash")})
    variation_paths = summarize_variations(successful)
    failed = [item for item in round_results if not item.get("ok")]
    report = {
        "prompt_file": str(prompt_file),
        "spec_docx": str(spec_docx),
        "rounds_requested": rounds,
        "rounds_completed": len(round_results),
        "success_count": len(successful),
        "failure_count": len(failed),
        "exact_match_across_successes": len(exact_hashes) <= 1 and len(successful) == rounds,
        "unique_success_hashes": exact_hashes,
        "varying_paths": variation_paths,
        "rounds": round_results,
    }
    if len(successful) >= 2:
        first = successful[0]["data"]
        report["diff_paths_against_first_success"] = {
            str(item["round"]): diff_paths(first, item["data"])
            for item in successful[1:]
            if item["hash"] != successful[0]["hash"]
        }
    else:
        report["diff_paths_against_first_success"] = {}
    return report


def write_markdown_summary(report: dict[str, Any], path: Path) -> None:
    lines = [
        "# Scene Prompt Regression",
        "",
        f"- Prompt: `{report['prompt_file']}`",
        f"- Spec: `{report['spec_docx']}`",
        f"- Rounds requested: {report['rounds_requested']}",
        f"- Success count: {report['success_count']}",
        f"- Failure count: {report['failure_count']}",
        f"- Exact match across successes: {report['exact_match_across_successes']}",
        "",
        "## Rounds",
        "",
    ]
    for item in report["rounds"]:
        status = "PASS" if item["ok"] else "FAIL"
        lines.append(f"- Round {item['round']:02d}: {status}")
        if item["issues"]:
            for issue in item["issues"][:10]:
                lines.append(f"  - {issue}")
    if report["varying_paths"]:
        lines.extend(["", "## Varying Paths", ""])
        for path_name, values in list(report["varying_paths"].items())[:80]:
            lines.append(f"- `{path_name}`")
            for value in values[:5]:
                lines.append(f"  - `{value}`")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run multi-round regression for scene prompt generation.")
    parser.add_argument(
        "--prompt-file",
        type=Path,
        default=ROOT / "src" / "scene" / "presets" / "AI_TEMPLATE_TO_SCENE_JSON_PROMPT_V2.md",
    )
    parser.add_argument(
        "--spec-docx",
        type=Path,
        default=ROOT / "tests" / "former-templates" / "1 南开大学研究生学位论文写作规范（2026版）.docx",
    )
    parser.add_argument(
        "--baseline-json",
        type=Path,
        default=ROOT / "src" / "scene" / "presets" / "default_format.json",
    )
    parser.add_argument("--rounds", type=int, default=10)
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=ROOT / "tests" / "_scene_prompt_regression_nankai_2026_v2",
    )
    parser.add_argument("--model", type=str, default=None)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    out_dir = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    prompt_template = args.prompt_file.read_text(encoding="utf-8")
    baseline_text = args.baseline_json.read_text(encoding="utf-8")
    baseline_data = json.loads(baseline_text)
    spec_text = extract_docx_text(args.spec_docx)
    prompt_text = build_prompt(prompt_template, baseline_text, spec_text, baseline_data)

    (out_dir / "prompt_filled.txt").write_text(prompt_text, encoding="utf-8")
    (out_dir / "spec_extracted.txt").write_text(spec_text, encoding="utf-8")

    round_results: list[dict[str, Any]] = []
    for round_idx in range(1, args.rounds + 1):
        returncode, stdout, stderr, raw_path = run_codex_round(
            prompt_text,
            out_dir,
            round_idx,
            model=args.model,
        )
        (out_dir / f"round_{round_idx:02d}.log").write_text(
            "\n".join(
                [
                    f"returncode={returncode}",
                    "stdout:",
                    stdout.rstrip(),
                    "stderr:",
                    stderr.rstrip(),
                ]
            )
            + "\n",
            encoding="utf-8",
        )
        round_result: dict[str, Any] = {
            "round": round_idx,
            "returncode": returncode,
            "ok": False,
            "issues": [],
            "hash": None,
        }
        if returncode != 0:
            round_result["issues"].append(f"codex exec failed with returncode={returncode}")
        elif not raw_path.exists():
            round_result["issues"].append("codex exec did not produce output-last-message file")
        else:
            raw_text = raw_path.read_text(encoding="utf-8")
            validation = validate_round(raw_text, baseline_data)
            round_result.update(
                {
                    "ok": validation["ok"],
                    "issues": validation["issues"],
                    "hash": validation["hash"],
                }
            )
            if validation["data"] is not None:
                normalized_path = out_dir / f"round_{round_idx:02d}.json"
                normalized_path.write_text(
                    json.dumps(validation["data"], ensure_ascii=False, indent=2, sort_keys=True) + "\n",
                    encoding="utf-8",
                )
                round_result["data"] = validation["data"]
        round_results.append(round_result)

    report = build_report(args.prompt_file, args.spec_docx, args.rounds, round_results)
    (out_dir / "summary.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    write_markdown_summary(report, out_dir / "summary.md")

    print(json.dumps(
        {
            "success_count": report["success_count"],
            "failure_count": report["failure_count"],
            "exact_match_across_successes": report["exact_match_across_successes"],
            "varying_path_count": len(report["varying_paths"]),
            "summary_json": str(out_dir / "summary.json"),
            "summary_md": str(out_dir / "summary.md"),
        },
        ensure_ascii=False,
        indent=2,
    ))
    return 0 if report["failure_count"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
