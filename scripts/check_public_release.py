from __future__ import annotations

import argparse
import json
from dataclasses import asdict, dataclass
from pathlib import Path
import sys


@dataclass(frozen=True)
class Rule:
    category: str
    pattern: str
    note: str
    severity: str = "blocker"


ROOT_DIR_RULES = [
    Rule("local_validation_dirs", "e2e_*", "Local E2E output directory."),
    Rule("local_validation_dirs", "exec-clone-check-*", "Clone-check output directory."),
    Rule("local_validation_dirs", "scene-upgrade-check-*", "Scene upgrade validation directory."),
    Rule("local_validation_dirs", "tmp_prompt_regression_exec_*", "Prompt regression execution output."),
    Rule("local_validation_dirs", "tmp_md_audit", "Markdown audit output."),
    Rule("local_validation_dirs", "tmp_md_symbol_audit", "Markdown symbol audit output."),
    Rule("local_validation_dirs", "tmp_toc_probe*", "TOC probe output."),
    Rule("local_validation_dirs", "scan_reports", "Scan report output."),
    Rule("local_build_dirs", "build", "Build output directory."),
    Rule("local_build_dirs", "dist", "Distribution output directory."),
    Rule("local_env_dirs", ".venv", "Local virtual environment."),
    Rule("local_env_dirs", ".pytest_cache", "pytest cache."),
]


EXACT_PATH_RULES = [
    Rule("private_test_materials", "tests/00", "Likely real thesis/document samples."),
    Rule("private_test_materials", "tests/clone", "Clone samples may include private or third-party files."),
    Rule("private_test_materials", "tests/former-templates", "Third-party template/spec originals."),
    Rule("private_test_materials", "tests/test-1", "Historical local test samples."),
    Rule("private_test_materials", "tests/test-equation", "Equation sample documents."),
    Rule("private_test_materials", "tests/test-markdown", "Markdown sample documents."),
    Rule("private_test_materials", "tests/目录", "Directory-sample documents."),
    Rule(
        "brand_assets_to_review",
        "src/ui/icons/github.svg",
        "Confirm redistribution rights or replace with a generic external-link icon.",
        severity="caution",
    ),
    Rule(
        "brand_assets_to_review",
        "src/ui/icons/bilibili.svg",
        "Confirm redistribution rights or replace with a generic external-link icon.",
        severity="caution",
    ),
]


RECURSIVE_GLOB_RULES = [
    Rule("test_temp_dirs", "tests/_tmp_*", "Temporary test directory."),
    Rule("test_temp_dirs", "tests/_clone_*", "Temporary clone-test directory."),
    Rule("prompt_regression_dirs", "tests/_scene_prompt_regression_*", "May include extracted spec text or prompt artifacts."),
    Rule("test_temp_dirs", "tests/lf_*", "Local test directory."),
    Rule("generated_outputs", "*_new.docx", "Generated comparison/output document."),
    Rule("generated_outputs", "*_对比稿.docx", "Generated comparison draft."),
    Rule("generated_outputs", "*_报告.json", "Generated JSON report."),
    Rule("generated_outputs", "*_报告.md", "Generated Markdown report."),
    Rule("generated_outputs", "tmp_*.docx", "Temporary DOCX artifact."),
    Rule("generated_outputs", "_tmp_*.docx", "Temporary DOCX artifact."),
    Rule("generated_outputs", "copytest-*.docx", "Intermediate copy-test document."),
    Rule("sensitive_extracted_text", "tests/_scene_prompt_regression_*/**/prompt_filled.txt", "May contain extracted third-party spec text."),
    Rule("sensitive_extracted_text", "tests/_scene_prompt_regression_*/**/spec_extracted.txt", "May contain extracted third-party spec text."),
    Rule("sensitive_extracted_text", "tests/_scene_prompt_regression_*/**/summary.json", "Regression summary artifact."),
    Rule("sensitive_extracted_text", "tests/_scene_prompt_regression_*/**/summary.md", "Regression summary artifact."),
    Rule("sensitive_extracted_text", "tests/_scene_prompt_regression_*/**/round_*.json", "Regression intermediate JSON."),
    Rule("sensitive_extracted_text", "tests/_scene_prompt_regression_*/**/round_*.log", "Regression log."),
    Rule("sensitive_extracted_text", "tests/_scene_prompt_regression_*/**/round_*_last_message.txt", "Regression last-message snapshot."),
    Rule("sensitive_extracted_text", "tmp_prompt_regression_exec_*/**/prompt_filled.txt", "May contain extracted third-party spec text."),
    Rule("sensitive_extracted_text", "tmp_prompt_regression_exec_*/**/spec_extracted.txt", "May contain extracted third-party spec text."),
    Rule("sensitive_extracted_text", "tmp_prompt_regression_exec_*/**/summary.json", "Regression summary artifact."),
    Rule("sensitive_extracted_text", "tmp_prompt_regression_exec_*/**/summary.md", "Regression summary artifact."),
    Rule("sensitive_extracted_text", "tmp_prompt_regression_exec_*/**/round_*.json", "Regression intermediate JSON."),
    Rule("sensitive_extracted_text", "tmp_prompt_regression_exec_*/**/round_*.log", "Regression log."),
    Rule("sensitive_extracted_text", "tmp_prompt_regression_exec_*/**/round_*_last_message.txt", "Regression last-message snapshot."),
]


@dataclass
class Finding:
    severity: str
    category: str
    path: str
    note: str
    pattern: str


def collect_findings(root: Path) -> list[Finding]:
    findings: list[Finding] = []
    seen: set[tuple[str, str, str]] = set()

    def add(rule: Rule, path: Path) -> None:
        try:
            rel = path.relative_to(root).as_posix()
        except ValueError:
            rel = path.as_posix()
        key = (rule.severity, rule.category, rel)
        if key in seen:
            return
        seen.add(key)
        findings.append(Finding(rule.severity, rule.category, rel, rule.note, rule.pattern))

    for rule in ROOT_DIR_RULES:
        for match in sorted(root.glob(rule.pattern)):
            if match.exists():
                add(rule, match)

    for rule in EXACT_PATH_RULES:
        path = root / rule.pattern
        if path.exists():
            add(rule, path)

    for rule in RECURSIVE_GLOB_RULES:
        for match in sorted(root.glob(rule.pattern)):
            if match.exists():
                add(rule, match)

    findings.sort(key=lambda item: (item.severity, item.category, item.path))
    return findings


def print_human(findings: list[Finding]) -> None:
    blockers = [f for f in findings if f.severity == "blocker"]
    cautions = [f for f in findings if f.severity == "caution"]

    if not findings:
        print("[OK] No predefined public-release blockers were found.")
        print("Please still review newly added fixtures, templates, and third-party assets manually.")
        return

    print("Public release scan results")
    print("=" * 72)

    current_group: tuple[str, str] | None = None
    for finding in findings:
        group = (finding.severity, finding.category)
        if group != current_group:
            label = "BLOCKER" if finding.severity == "blocker" else "CAUTION"
            print(f"\n[{label}] {finding.category}")
            current_group = group
        print(f"- {finding.path}")
        print(f"  note: {finding.note}")

    print("\n" + "=" * 72)
    print(f"blockers: {len(blockers)} | cautions: {len(cautions)}")
    print("Recommendation: run `python scripts/check_public_release.py --strict` before publishing.")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Scan the repository for files and directories that should not be uploaded publicly."
    )
    parser.add_argument("--strict", action="store_true", help="Return a non-zero exit code when blockers are found.")
    parser.add_argument("--json", action="store_true", help="Print results as JSON.")
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[1]
    findings = collect_findings(root)

    if args.json:
        payload = {
            "root": root.as_posix(),
            "blockers": [asdict(f) for f in findings if f.severity == "blocker"],
            "cautions": [asdict(f) for f in findings if f.severity == "caution"],
        }
        json.dump(payload, sys.stdout, ensure_ascii=False, indent=2)
        sys.stdout.write("\n")
    else:
        print_human(findings)

    if args.strict and any(f.severity == "blocker" for f in findings):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
