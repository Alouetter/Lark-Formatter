"""场景加载/保存/导入导出"""

import copy
from dataclasses import asdict
import json
import os
import re
import sys
import tempfile
from pathlib import Path
from src.scene.schema import (
    SceneConfig, PageSetupConfig, MarginConfig,
    HeadingNumberingConfig, HeadingLevelConfig,
    HeadingNumberingV2Config,
    NumberShellConfig, NumberCoreStyleConfig,
    NumberChainSegmentConfig, NumberChainConfig,
    NumberPresetConfig, HeadingLevelBindingConfig,
    HeadingModelConfig,
    EnforcementConfig, HeadingRiskGuardConfig,
    StyleConfig, OutputConfig,
    FormatScopeConfig, TocConfig, CaptionConfig, ChemTypographyConfig, MdCleanupConfig,
    WhitespaceNormalizeConfig, CitationLinkConfig,
    FormulaConvertConfig, FormulaToTableConfig, EquationTableFormatConfig, FormulaStyleConfig,
    FormulaTableConfig,
    REMOVED_NUMBER_CORE_STYLE_IDS,
    default_chain_id_for_level, heading_level_index, heading_level_keys,
)
from src.utils.heading_numbering_v2 import legacy_levels_from_v2
from src.utils.indent import sync_style_config_indent_fields


def _builtin_seed_candidates() -> list[Path]:
    """Candidate directories for built-in presets (highest priority first)."""
    candidates: list[Path] = []
    if getattr(sys, "frozen", False):
        meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        exe_dir = Path(sys.executable).resolve().parent
        candidates.extend(
            [
                meipass / "src" / "scene" / "presets",
                meipass / "templates",
                meipass / "_internal" / "templates",
                exe_dir / "_internal" / "src" / "scene" / "presets",
                exe_dir / "_internal" / "templates",
                exe_dir / "src" / "scene" / "presets",
                exe_dir / "presets",
                exe_dir / "templates",
            ]
        )
    candidates.append(Path(__file__).resolve().parent / "presets")

    deduped: list[Path] = []
    seen: set[str] = set()
    for p in candidates:
        key = str(p)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(p)
    return deduped


def _resolve_presets_dir() -> Path:
    """Resolve built-in preset directory (read-only seed source)."""
    candidates = _builtin_seed_candidates()

    # Prefer directory that actually contains preset JSON files.
    for p in candidates:
        try:
            if p.exists() and p.is_dir() and any(p.glob("*.json")):
                return p
        except Exception:
            continue

    for p in candidates:
        if p.exists():
            return p

    return candidates[-1]


def _resolve_templates_dir() -> Path:
    """Resolve user-writable templates directory next to the executable.

    In frozen (packaged) mode: <exe_dir>/templates
    In source mode: prefer a standalone <project_root>/templates directory so
    built-in presets remain an immutable seed source.
    """
    env_dir = os.environ.get("DOCX_FORMATTER_TEMPLATES_DIR", "").strip()
    if env_dir:
        return Path(env_dir).expanduser()

    exe_templates = Path(sys.executable).resolve().parent / "templates"
    if getattr(sys, "frozen", False):
        return exe_templates

    # Non-frozen fallback: if a standalone templates directory exists,
    # prefer it over writing into source presets.
    cwd_templates = Path.cwd().resolve() / "templates"
    project_templates = Path(__file__).resolve().parents[2] / "templates"
    for candidate in (exe_templates, cwd_templates, project_templates):
        if candidate.exists() and candidate.is_dir():
            return candidate

    # 开发模式默认落到项目根目录下独立 templates，避免直接改写内置预设。
    return project_templates


def _is_valid_scene_json(path: Path) -> bool:
    try:
        with open(path, "r", encoding="utf-8") as f:
            json.load(f)
        return True
    except Exception:
        return False


def _load_scene_json_payload(path: Path) -> dict | None:
    try:
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        return payload if isinstance(payload, dict) else None
    except Exception:
        return None


def _default_scene_payload_needs_repair(payload: dict | None) -> bool:
    """Return True when a default preset payload looks semantically broken."""
    if not isinstance(payload, dict):
        return True

    if str(payload.get("name", "") or "").strip() != "默认格式":
        return True

    heading_numbering = payload.get("heading_numbering", {})
    if not isinstance(heading_numbering, dict):
        return True
    levels = heading_numbering.get("levels", {})
    if not isinstance(levels, dict) or not levels:
        return True

    styles = payload.get("styles", {})
    if not isinstance(styles, dict) or not styles:
        return True

    return False


def _backup_scene_file(path: Path) -> bool:
    backup = path.with_suffix(path.suffix + ".broken")
    idx = 1
    while backup.exists():
        backup = path.with_suffix(path.suffix + f".broken{idx}")
        idx += 1
    try:
        import shutil

        shutil.move(str(path), str(backup))
        return True
    except Exception:
        return False


def _choose_seed_dir(preferred: Path, templates_dir: Path) -> Path | None:
    """Pick a usable preset seed directory, avoiding empty/self-only folders."""
    try:
        preferred_has_json = preferred.exists() and preferred.is_dir() and any(preferred.glob("*.json"))
    except Exception:
        preferred_has_json = False
    if preferred_has_json:
        return preferred

    templates_resolved = ""
    try:
        templates_resolved = str(templates_dir.resolve())
    except Exception:
        templates_resolved = str(templates_dir)

    for candidate in _builtin_seed_candidates():
        try:
            if not candidate.exists() or not candidate.is_dir() or not any(candidate.glob("*.json")):
                continue
            candidate_resolved = str(candidate.resolve())
        except Exception:
            continue
        if candidate_resolved == templates_resolved:
            continue
        return candidate
    return preferred if preferred.exists() and preferred.is_dir() else None


def _ensure_templates(builtin_dir: Path, templates_dir: Path) -> None:
    """Copy built-in presets into the user-writable templates folder.

    Only copies files that do not already exist in templates_dir,
    so user modifications are preserved.
    """
    import shutil
    seed_dir = _choose_seed_dir(builtin_dir, templates_dir)
    if seed_dir is None or not seed_dir.exists():
        return
    templates_dir.mkdir(parents=True, exist_ok=True)
    for src_file in seed_dir.glob("*.json"):
        dst_file = templates_dir / src_file.name
        if not dst_file.exists():
            shutil.copy2(src_file, dst_file)
            continue

        try:
            same_file = src_file.resolve() == dst_file.resolve()
        except Exception:
            same_file = False
        if same_file:
            continue

        # Recover broken built-in preset cache in templates by re-copying
        # from packaged defaults while preserving a backup for troubleshooting.
        src_payload = _load_scene_json_payload(src_file)
        dst_payload = _load_scene_json_payload(dst_file)
        if src_payload is not None and dst_payload is None:
            if not _backup_scene_file(dst_file):
                continue
            shutil.copy2(src_file, dst_file)
            continue

        # Recover a semantically empty/corrupted default preset even when the
        # JSON itself still parses successfully.
        if (
            src_file.name.lower() == "default_format.json"
            and src_payload is not None
            and not _default_scene_payload_needs_repair(src_payload)
            and _default_scene_payload_needs_repair(dst_payload)
        ):
            if not _backup_scene_file(dst_file):
                continue
            shutil.copy2(src_file, dst_file)
    # Also copy subdirectories (e.g. formats/)
    for src_sub in seed_dir.iterdir():
        if src_sub.is_dir():
            dst_sub = templates_dir / src_sub.name
            if not dst_sub.exists():
                shutil.copytree(src_sub, dst_sub)


def _default_scene_candidate_paths() -> list[Path]:
    candidates: list[Path] = [
        PRESETS_DIR / "default_format.json",
        _BUILTIN_PRESETS_DIR / "default_format.json",
    ]
    for candidate_dir in _builtin_seed_candidates():
        candidates.append(candidate_dir / "default_format.json")

    deduped: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        try:
            key = str(candidate.resolve())
        except Exception:
            key = str(candidate)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate)
    return deduped


_BUILTIN_PRESETS_DIR = _resolve_presets_dir()
PRESETS_DIR = _resolve_templates_dir()

_SCENE_UPGRADE_TOP_LEVEL_LABELS = {
    "heading_numbering_v2": "新版标题编号定义",
    "toc": "目录模式配置",
    "md_cleanup": "Markdown 清理配置",
    "formula_convert": "公式转换配置",
    "formula_to_table": "公式转表格配置",
    "equation_table_format": "公式表格格式配置",
    "formula_style": "公式样式配置",
    "formula_table": "公式表格视觉配置",
}

# Ensure templates folder is populated on first launch
try:
    _ensure_templates(_BUILTIN_PRESETS_DIR, PRESETS_DIR)
except Exception:
    pass  # Don't crash on permission errors etc.


def _parse_margin(data: dict) -> MarginConfig:
    return MarginConfig(**{k: data[k] for k in MarginConfig.__dataclass_fields__ if k in data})


def _parse_style(data: dict) -> StyleConfig:
    return StyleConfig(**{k: data[k] for k in StyleConfig.__dataclass_fields__ if k in data})


def _parse_heading_level(data: dict) -> HeadingLevelConfig:
    return HeadingLevelConfig(**{k: data[k] for k in HeadingLevelConfig.__dataclass_fields__ if k in data})


def _parse_number_shell(data: dict) -> NumberShellConfig:
    return NumberShellConfig(**{k: data[k] for k in NumberShellConfig.__dataclass_fields__ if k in data})


def _parse_number_core_style(data: dict) -> NumberCoreStyleConfig:
    return NumberCoreStyleConfig(**{k: data[k] for k in NumberCoreStyleConfig.__dataclass_fields__ if k in data})


def _parse_number_chain_segment(data: dict) -> NumberChainSegmentConfig:
    return NumberChainSegmentConfig(
        **{k: data[k] for k in NumberChainSegmentConfig.__dataclass_fields__ if k in data}
    )


def _parse_number_chain(data: dict) -> NumberChainConfig:
    payload = data if isinstance(data, dict) else {}
    segments_raw = payload.get("segments", [])
    segments = []
    if isinstance(segments_raw, list):
        for item in segments_raw:
            if isinstance(item, dict):
                segments.append(_parse_number_chain_segment(item))
    merged = {k: payload[k] for k in NumberChainConfig.__dataclass_fields__ if k in payload and k != "segments"}
    merged["segments"] = segments
    return NumberChainConfig(**merged)


def _parse_number_preset(data: dict) -> NumberPresetConfig:
    return NumberPresetConfig(**{k: data[k] for k in NumberPresetConfig.__dataclass_fields__ if k in data})


def _parse_heading_level_binding(data: dict) -> HeadingLevelBindingConfig:
    return HeadingLevelBindingConfig(
        **{k: data[k] for k in HeadingLevelBindingConfig.__dataclass_fields__ if k in data}
    )


def _normalize_number_core_style_id(
    core_style_id: str,
    *,
    catalog: dict[str, NumberCoreStyleConfig] | None = None,
    fallback: str = "arabic",
) -> str:
    normalized = str(core_style_id or "").strip()
    if not normalized or normalized in REMOVED_NUMBER_CORE_STYLE_IDS:
        normalized = fallback
    if catalog is not None and normalized not in catalog:
        normalized = fallback if fallback in catalog else next(iter(catalog), fallback)
    return normalized or fallback


def _normalize_catalog_entry_id(
    entry_id: str,
    *,
    catalog: dict[str, object] | None = None,
    fallback: str,
) -> str:
    normalized = str(entry_id or "").strip()
    if catalog is None:
        return normalized or fallback
    if normalized in catalog:
        return normalized
    if fallback in catalog:
        return fallback
    return next(iter(catalog), fallback)


def _merge_catalog(default_catalog: dict, incoming: dict, parser):
    merged = {}
    for key, value in (default_catalog or {}).items():
        merged[str(key)] = copy.deepcopy(value)
    if not isinstance(incoming, dict):
        return merged
    for raw_key, raw_value in incoming.items():
        key = str(raw_key or "").strip()
        if not key:
            continue
        base = merged.get(key)
        payload = asdict(base) if base is not None else {}
        if isinstance(raw_value, dict):
            payload.update(raw_value)
        merged[key] = parser(payload)
    return merged


def _parse_heading_numbering_v2(data: dict) -> HeadingNumberingV2Config:
    default_cfg = HeadingNumberingV2Config()
    payload = data if isinstance(data, dict) else {}
    core_style_catalog = _merge_catalog(
        default_cfg.core_style_catalog,
        payload.get("core_style_catalog", {}),
        _parse_number_core_style,
    )
    for style_id in REMOVED_NUMBER_CORE_STYLE_IDS:
        core_style_catalog.pop(style_id, None)

    level_bindings = _merge_catalog(
        default_cfg.level_bindings,
        payload.get("level_bindings", {}),
        _parse_heading_level_binding,
    )
    for binding in level_bindings.values():
        binding.display_core_style = _normalize_number_core_style_id(
            binding.display_core_style,
            catalog=core_style_catalog,
        )
        binding.reference_core_style = _normalize_number_core_style_id(
            binding.reference_core_style,
            catalog=core_style_catalog,
            fallback=binding.display_core_style,
        )
    shell_catalog = _merge_catalog(
        default_cfg.shell_catalog,
        payload.get("shell_catalog", {}),
        _parse_number_shell,
    )
    chain_catalog = _merge_catalog(
        default_cfg.chain_catalog,
        payload.get("chain_catalog", {}),
        _parse_number_chain,
    )
    for level_name, binding in level_bindings.items():
        binding.display_shell = _normalize_catalog_entry_id(
            binding.display_shell,
            catalog=shell_catalog,
            fallback="plain",
        )
        binding.chain = _normalize_catalog_entry_id(
            binding.chain,
            catalog=chain_catalog,
            fallback=default_chain_id_for_level(level_name),
        )

    return HeadingNumberingV2Config(
        enabled=bool(payload.get("enabled", default_cfg.enabled)),
        shell_catalog=shell_catalog,
        core_style_catalog=core_style_catalog,
        chain_catalog=chain_catalog,
        preset_catalog=_merge_catalog(
            default_cfg.preset_catalog,
            payload.get("preset_catalog", {}),
            _parse_number_preset,
        ),
        level_bindings=level_bindings,
    )


def _parse_heading_model(data: dict) -> HeadingModelConfig:
    default = HeadingModelConfig()
    payload = data if isinstance(data, dict) else {}

    def _merge_dict(key: str, default_value: dict[str, str], *, lower_values: bool = False) -> dict[str, str]:
        incoming = payload.get(key, {})
        if not isinstance(incoming, dict):
            return dict(default_value)
        merged = dict(default_value)
        for raw_k, raw_v in incoming.items():
            k = str(raw_k or "").strip()
            v = str(raw_v or "").strip()
            if not k or not v:
                continue
            merged[k] = v.lower() if lower_values else v
        return merged

    def _merge_list(key: str, default_value: list[str], *, lower: bool = False) -> list[str]:
        if key not in payload:
            return list(default_value)
        incoming = payload.get(key)
        if not isinstance(incoming, list):
            return list(default_value)
        merged = []
        for raw in incoming:
            text = str(raw or "").strip()
            if not text:
                continue
            val = text.lower() if lower else text
            if val not in merged:
                merged.append(val)
        return merged

    non_numbered_style = str(
        payload.get("non_numbered_heading_style_name", default.non_numbered_heading_style_name) or ""
    ).strip() or default.non_numbered_heading_style_name

    return HeadingModelConfig(
        level_to_word_style=_merge_dict("level_to_word_style", default.level_to_word_style),
        level_to_style_key=_merge_dict("level_to_style_key", default.level_to_style_key),
        style_alias_to_level=_merge_dict(
            "style_alias_to_level", default.style_alias_to_level, lower_values=True
        ),
        section_title_style_map=_merge_dict("section_title_style_map", default.section_title_style_map),
        non_numbered_title_sections=_merge_list(
            "non_numbered_title_sections", default.non_numbered_title_sections, lower=True
        ),
        non_numbered_title_texts=_merge_list("non_numbered_title_texts", default.non_numbered_title_texts),
        front_matter_title_texts=_merge_list("front_matter_title_texts", default.front_matter_title_texts),
        post_section_types=_merge_list("post_section_types", default.post_section_types, lower=True),
        non_numbered_heading_style_name=non_numbered_style,
        header_front_text=_merge_dict("header_front_text", default.header_front_text),
        header_back_types=_merge_list("header_back_types", default.header_back_types, lower=True),
    )


def _normalize_pipeline_steps(raw_steps) -> list[str]:
    """Normalize pipeline list to stable, de-duplicated, non-empty step names."""
    if not isinstance(raw_steps, list):
        return []
    normalized: list[str] = []
    for item in raw_steps:
        name = str(item).strip()
        if name and name not in normalized:
            normalized.append(name)
    return normalized


def _is_auto_managed_critical_rules(
    configured_rules: list[str],
    pipeline_steps: list[str],
) -> bool:
    normalized_rules = _normalize_pipeline_steps(configured_rules)
    normalized_pipeline = _normalize_pipeline_steps(pipeline_steps)
    if not normalized_rules:
        return True
    if normalized_rules == normalized_pipeline:
        return True
    return normalized_rules == _normalize_pipeline_steps(getattr(SceneConfig(), "pipeline_critical_rules", []))


def _sync_whitespace_pipeline_switch(config: SceneConfig) -> None:
    """Keep whitespace_normalize enabled flag and pipeline step in sync."""
    pipeline = list(config.pipeline or [])
    ws_cfg = getattr(config, "whitespace_normalize", None)
    if ws_cfg is None:
        return

    enabled = bool(getattr(ws_cfg, "enabled", False))
    has_step = "whitespace_normalize" in pipeline

    if enabled and not has_step:
        if "md_cleanup" in pipeline:
            idx = pipeline.index("md_cleanup") + 1
        elif "style_manager" in pipeline:
            idx = pipeline.index("style_manager")
        else:
            idx = 0
        pipeline.insert(idx, "whitespace_normalize")
        has_step = True
    elif (not enabled) and has_step:
        pipeline = [step for step in pipeline if step != "whitespace_normalize"]
        has_step = False

    config.pipeline = pipeline
    ws_cfg.enabled = has_step


def _sync_formula_pipeline_switch(config: SceneConfig) -> None:
    """Keep formula feature enabled flags and pipeline steps in sync."""
    pipeline = list(config.pipeline or [])
    convert_cfg = getattr(config, "formula_convert", None)
    to_table_cfg = getattr(config, "formula_to_table", None)
    eq_table_cfg = getattr(config, "equation_table_format", None)
    style_cfg = getattr(config, "formula_style", None)

    if not all([convert_cfg, to_table_cfg, eq_table_cfg, style_cfg]):
        return

    formula_steps = [
        ("formula_convert", convert_cfg),
        ("formula_to_table", to_table_cfg),
        ("equation_table_format", eq_table_cfg),
        ("formula_style", style_cfg),
    ]

    pipeline = [step for step in pipeline if step not in {name for name, _ in formula_steps}]

    enabled_steps = [name for name, cfg in formula_steps if bool(getattr(cfg, "enabled", False))]
    if enabled_steps:
        if "table_format" in pipeline:
            idx = pipeline.index("table_format") + 1
        elif "section_format" in pipeline:
            idx = pipeline.index("section_format")
        else:
            idx = len(pipeline)
        for step in enabled_steps:
            pipeline.insert(idx, step)
            idx += 1

    config.pipeline = pipeline
    enabled_set = set(pipeline)
    for step, cfg in formula_steps:
        cfg.enabled = step in enabled_set


def _sync_citation_link_pipeline_switch(config: SceneConfig) -> None:
    """Keep citation_link enabled flag and pipeline step in sync."""
    pipeline = list(config.pipeline or [])
    citation_cfg = getattr(config, "citation_link", None)
    if citation_cfg is None:
        return

    enabled = bool(getattr(citation_cfg, "enabled", False))
    has_step = "citation_link" in pipeline

    if enabled and not has_step:
        idx = len(pipeline)
        for anchor in (
            "section_format",
            "formula_style",
            "equation_table_format",
            "formula_to_table",
            "formula_convert",
            "table_format",
        ):
            if anchor in pipeline:
                idx = pipeline.index(anchor) + 1
                break
        pipeline.insert(idx, "citation_link")
        has_step = True
    elif (not enabled) and has_step:
        pipeline = [step for step in pipeline if step != "citation_link"]
        has_step = False

    config.pipeline = pipeline
    citation_cfg.enabled = has_step


def _sync_pipeline_critical_rules(config: SceneConfig) -> None:
    """Auto-manage critical rules unless the scene explicitly customizes them."""
    source = str(getattr(config, "_pipeline_critical_rules_source", "auto") or "auto").strip().lower()
    if source == "payload_custom":
        config.pipeline_critical_rules = _normalize_pipeline_steps(config.pipeline_critical_rules)
        return
    config.pipeline_critical_rules = _normalize_pipeline_steps(config.pipeline)


def _deep_merge(base: dict, override: dict) -> dict:
    """深度合并，使用 override 中的值覆盖 base。"""
    result = base.copy()
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def _collect_scene_upgrade_notes(original_payload: dict, migrated_payload: dict) -> list[str]:
    if not isinstance(original_payload, dict):
        return []

    notes: list[str] = []

    if original_payload != migrated_payload:
        notes.append("已迁移旧版标题层级键/引用到当前 heading1-heading8 结构。")

    if "heading_numbering" in original_payload and "heading_numbering_v2" not in original_payload:
        notes.append("已从旧版标题编号定义派生新版 heading_numbering_v2 结构。")

    missing_top_labels = [
        label
        for key, label in _SCENE_UPGRADE_TOP_LEVEL_LABELS.items()
        if key not in original_payload
    ]
    if missing_top_labels:
        notes.append(f"已补全当前版本配置项: {', '.join(missing_top_labels)}。")

    format_scope = original_payload.get("format_scope")
    if isinstance(format_scope, dict) and "page_ranges_text" not in format_scope:
        notes.append("已补全格式范围页码表达字段 page_ranges_text。")

    deduped: list[str] = []
    for note in notes:
        text = str(note or "").strip()
        if text and text not in deduped:
            deduped.append(text)
    return deduped


# ── 旧版 key 迁移（chapter/level → heading1-4）──

def _migrate_legacy_keys(data: dict) -> dict:
    """将旧版 chapter/level 键名迁移为 heading1-4。

    处理两类旧配置：
    1. 语义名格式：chapter / level1 / level2 / level3
    2. 样式 key 格式：chapter_heading / heading1(=H2) / heading2(=H3) / heading3(=H4)

    两阶段避免冲突：先用临时 key 暂存有碰撞风险的条目，再统一替换。
    """
    if not isinstance(data, dict):
        return data

    # 快速检测是否需要迁移
    raw = json.dumps(data)
    old_markers = ("chapter_heading", "level1", "level2", "level3")
    # "chapter" 单独检测：必须是精确的 JSON key（"chapter"），
    # 不能匹配 "chapter_heading" 或 "chapter.seq" 等
    has_old = any(f'"{k}"' in raw for k in old_markers)
    if not has_old:
        # 检查 "chapter" 作为独立 key（排除 chapter_heading 等）
        import re as _re
        has_old = bool(_re.search(r'"chapter"(?:\s*:)', raw))
    if not has_old:
        return data

    def _migrate_dict(d):
        """递归迁移单个 dict 及其子 dict。"""
        if not isinstance(d, dict):
            return d

        # 先递归处理子节点
        processed = {}
        for k, v in d.items():
            if isinstance(v, dict):
                processed[k] = _migrate_dict(v)
            elif isinstance(v, list):
                processed[k] = [_migrate_dict(item) if isinstance(item, dict) else item for item in v]
            else:
                processed[k] = v

        # 然后重命名当前层级的 key
        has_chapter = "chapter" in processed
        has_chapter_heading = "chapter_heading" in processed
        has_level1 = "level1" in processed
        has_any_old = has_chapter or has_chapter_heading or has_level1 or "level2" in processed or "level3" in processed

        if not has_any_old:
            # 无旧 key，也不需要 shift heading1/2/3
            # 还需检查 string value 中的旧引用
            result = {}
            for k, v in processed.items():
                if isinstance(v, str):
                    v = _migrate_str_value(k, v)
                result[k] = v
            return result

        # 有旧 key → 需要重命名
        # 使用临时 key 避免碰撞
        temp = {}
        for k, v in processed.items():
            if isinstance(v, str):
                v = _migrate_str_value(k, v)
            if k == "chapter" or k == "chapter_heading":
                temp.setdefault("__H1__", v)  # 合并到 heading1
            elif k == "level1":
                temp["__H2__"] = v
            elif k == "level2":
                temp["__H3__"] = v
            elif k == "level3":
                temp["__H4__"] = v
            elif has_any_old and k == "heading1":
                # 旧体系中 heading1 = Heading 2，需要 shift
                temp["__H2__"] = v
            elif has_any_old and k == "heading2":
                temp["__H3__"] = v
            elif has_any_old and k == "heading3":
                temp["__H4__"] = v
            else:
                temp[k] = v

        # 替换临时 key 为最终名
        result = {}
        for k, v in temp.items():
            if k == "__H1__":
                result["heading1"] = v
            elif k == "__H2__":
                result["heading2"] = v
            elif k == "__H3__":
                result["heading3"] = v
            elif k == "__H4__":
                result["heading4"] = v
            else:
                result[k] = v
        return result

    _STR_VALUE_MAP = {
        "chapter_heading": "heading1",
        "chapter": "heading1",
        "level1": "heading2",
        "level2": "heading3",
        "level3": "heading4",
    }

    def _migrate_str_value(key: str, v: str) -> str:
        if str(key or "").strip().lower() == "source":
            return v
        return _STR_VALUE_MAP.get(v, v)

    return _migrate_dict(data)


_HEADING_TEMPLATE_TOKEN_RE = re.compile(r"\{(?P<name>[a-zA-Z][a-zA-Z0-9_]*)\}")
_HEADING_LEVEL_TOKEN_RE = re.compile(r"^level([1-8])$")


def _legacy_default_template(level_name: str, format_name: str) -> str:
    fmt = str(format_name or "").strip().lower()
    if fmt == "chinese_chapter":
        return "第{current}章"
    if fmt == "chinese_section":
        return "第{current}节"
    if fmt == "chinese_ordinal":
        return "{current}、"
    if fmt == "chinese_ordinal_paren":
        return "（{current}）"
    if fmt == "arabic_dotted":
        level_idx = heading_level_index(level_name)
        if level_idx <= 1:
            return "{current}"
        return ".".join([f"{{level{i}}}" for i in range(1, level_idx)] + ["{current}"])
    return "{current}"


def _legacy_format_to_core_style(format_name: str) -> str:
    fmt = str(format_name or "").strip().lower()
    if fmt in {"chinese_chapter", "chinese_section", "chinese_ordinal", "chinese_ordinal_paren"}:
        return "cn_lower"
    if fmt == "arabic_dotted":
        return "arabic"
    if fmt in REMOVED_NUMBER_CORE_STYLE_IDS:
        return "arabic"
    if fmt in {
        "arabic",
        "arabic_pad2",
        "cn_lower",
        "cn_upper",
        "roman_upper",
        "circled",
        "circled_paren",
    }:
        return fmt
    return "arabic"


def _canonicalize_legacy_template(level_name: str, format_name: str, template: str) -> tuple[str, bool]:
    raw_template = str(template or "")
    if not raw_template:
        raw_template = _legacy_default_template(level_name, format_name)
    elif (
        raw_template == "{current}"
        and str(format_name or "").strip().lower()
        in {"chinese_chapter", "chinese_section", "chinese_ordinal", "chinese_ordinal_paren", "arabic_dotted"}
    ):
        raw_template = _legacy_default_template(level_name, format_name)

    level_idx = heading_level_index(level_name)
    used_legacy_parent = False

    def _replace(match: re.Match[str]) -> str:
        nonlocal used_legacy_parent

        name = str(match.group("name") or "").strip()
        if name in {"current", "n", "cn"}:
            return "{current}"
        if name == "parent":
            used_legacy_parent = True
            if level_idx <= 1:
                return "{current}"
            return ".".join(f"{{level{i}}}" for i in range(1, level_idx))
        if _HEADING_LEVEL_TOKEN_RE.match(name):
            return "{" + name + "}"
        return match.group(0)

    return _HEADING_TEMPLATE_TOKEN_RE.sub(_replace, raw_template), used_legacy_parent


def _legacy_level_semantically_matches_projection(
    level_name: str,
    existing_level: HeadingLevelConfig,
    projected_level: HeadingLevelConfig,
) -> bool:
    if (
        str(existing_level.format or "").strip().lower()
        != str(projected_level.format or "").strip().lower()
    ):
        return False

    if str(existing_level.effective_separator or "") != str(projected_level.effective_separator or ""):
        return False

    existing_template, existing_used_legacy_parent = _canonicalize_legacy_template(
        level_name,
        existing_level.format,
        existing_level.template,
    )
    projected_template, projected_used_legacy_parent = _canonicalize_legacy_template(
        level_name,
        projected_level.format,
        projected_level.template,
    )
    return (
        existing_template == projected_template
        and existing_used_legacy_parent == projected_used_legacy_parent
    )


def _project_legacy_heading_levels_from_v2(
    v2_config: HeadingNumberingV2Config,
    *,
    existing_levels: dict[str, HeadingLevelConfig] | None = None,
) -> dict[str, HeadingLevelConfig]:
    projected_levels = legacy_levels_from_v2(v2_config, existing_levels=existing_levels)
    if not isinstance(existing_levels, dict) or not existing_levels:
        return projected_levels

    merged_levels: dict[str, HeadingLevelConfig] = {}
    for level_name, projected_level in projected_levels.items():
        existing_level = existing_levels.get(level_name)
        if (
            isinstance(existing_level, HeadingLevelConfig)
            and _legacy_level_semantically_matches_projection(
                level_name,
                existing_level,
                projected_level,
            )
        ):
            merged_levels[level_name] = HeadingLevelConfig(**vars(existing_level))
            continue
        merged_levels[level_name] = projected_level
    return merged_levels


def _template_placeholder_to_source(name: str) -> str | None:
    token_name = str(name or "").strip()
    if token_name == "current":
        return "current"
    if _HEADING_LEVEL_TOKEN_RE.match(token_name):
        return token_name
    return None


def _shell_signature(shell: NumberShellConfig) -> tuple[str, str]:
    return (str(shell.prefix or ""), str(shell.suffix or ""))


def _chain_segment_signature(segment: NumberChainSegmentConfig) -> tuple[str, str, str]:
    return (
        str(segment.type or ""),
        str(segment.source or ""),
        str(segment.text or ""),
    )


def _chain_signature(segments: list[NumberChainSegmentConfig]) -> tuple[tuple[str, str, str], ...]:
    return tuple(_chain_segment_signature(segment) for segment in segments)


def _allocate_custom_catalog_id(prefix: str, level_name: str, existing: dict[str, object]) -> str:
    base = f"{prefix}_{str(level_name or '').strip().lower() or 'heading'}"
    candidate = base
    idx = 1
    while candidate in existing:
        candidate = f"{base}_{idx}"
        idx += 1
    return candidate


def _resolve_shell_id(
    shell_catalog: dict[str, NumberShellConfig],
    *,
    level_name: str,
    prefix: str,
    suffix: str,
) -> str:
    target_signature = (str(prefix or ""), str(suffix or ""))
    for shell_id, shell in shell_catalog.items():
        if _shell_signature(shell) == target_signature:
            return shell_id

    shell_id = _allocate_custom_catalog_id("custom_shell", level_name, shell_catalog)
    shell_catalog[shell_id] = NumberShellConfig(
        label=f"{target_signature[0]}{{}}{target_signature[1]}",
        prefix=target_signature[0],
        suffix=target_signature[1],
    )
    return shell_id


def _chain_source_label(source: str) -> str:
    if source == "current":
        return "当前"
    if source.startswith("level"):
        try:
            idx = int(source[5:])
        except ValueError:
            return source
        labels = {
            1: "一级",
            2: "二级",
            3: "三级",
            4: "四级",
            5: "五级",
            6: "六级",
            7: "七级",
            8: "八级",
        }
        return labels.get(idx, source)
    return source


def _build_chain_label(segments: list[NumberChainSegmentConfig]) -> str:
    parts: list[str] = []
    for segment in segments:
        if segment.type == "literal":
            parts.append(str(segment.text or ""))
            continue
        parts.append(_chain_source_label(segment.source))
    label = "".join(parts).strip()
    return label or "自定义"


def _resolve_chain_id(
    chain_catalog: dict[str, NumberChainConfig],
    *,
    level_name: str,
    segments: list[NumberChainSegmentConfig],
) -> str:
    target_signature = _chain_signature(segments)
    for chain_id, chain in chain_catalog.items():
        if _chain_signature(chain.segments) == target_signature:
            return chain_id

    chain_id = _allocate_custom_catalog_id("custom_chain", level_name, chain_catalog)
    chain_catalog[chain_id] = NumberChainConfig(
        label=_build_chain_label(segments),
        segments=copy.deepcopy(segments),
    )
    return chain_id


def _template_to_shell_and_chain(
    level_name: str,
    format_name: str,
    template: str,
    *,
    shell_catalog: dict[str, NumberShellConfig],
    chain_catalog: dict[str, NumberChainConfig],
) -> tuple[str, str, bool]:
    normalized_template, used_legacy_parent = _canonicalize_legacy_template(
        level_name, format_name, template
    )
    matches: list[tuple[re.Match[str], str]] = []
    for match in _HEADING_TEMPLATE_TOKEN_RE.finditer(normalized_template):
        source = _template_placeholder_to_source(match.group("name"))
        if source is None:
            continue
        matches.append((match, source))

    if not matches:
        shell_id = _resolve_shell_id(
            shell_catalog,
            level_name=level_name,
            prefix=normalized_template,
            suffix="",
        )
        return shell_id, "current_only", used_legacy_parent

    first_match = matches[0][0]
    last_match = matches[-1][0]
    prefix = normalized_template[:first_match.start()]
    suffix = normalized_template[last_match.end():]

    segments: list[NumberChainSegmentConfig] = []
    cursor = first_match.start()
    for match, source in matches:
        literal = normalized_template[cursor:match.start()]
        if literal:
            segments.append(NumberChainSegmentConfig(type="literal", text=literal))
        segments.append(NumberChainSegmentConfig(type="value", source=source))
        cursor = match.end()

    tail_literal = normalized_template[cursor:last_match.end()]
    if tail_literal:
        segments.append(NumberChainSegmentConfig(type="literal", text=tail_literal))

    shell_id = _resolve_shell_id(
        shell_catalog,
        level_name=level_name,
        prefix=prefix,
        suffix=suffix,
    )
    chain_id = _resolve_chain_id(
        chain_catalog,
        level_name=level_name,
        segments=segments or [NumberChainSegmentConfig(type="value", source="current")],
    )
    return shell_id, chain_id, used_legacy_parent


def _derive_heading_numbering_v2_from_legacy(
    legacy_config: HeadingNumberingConfig,
    *,
    enabled: bool = True,
) -> HeadingNumberingV2Config:
    derived = HeadingNumberingV2Config(enabled=enabled)
    active_levels = {
        level_name: level_cfg
        for level_name, level_cfg in (legacy_config.levels or {}).items()
        if level_name in heading_level_keys()
    }
    if not active_levels:
        return derived

    for level_name in heading_level_keys():
        binding = derived.level_bindings.get(level_name)
        if binding is not None:
            binding.enabled = False

    decimal_reference_levels: set[str] = set()

    for level_name in heading_level_keys():
        level_cfg = active_levels.get(level_name)
        if level_cfg is None:
            continue

        display_core_style = _legacy_format_to_core_style(level_cfg.format)
        display_shell, chain_id, used_legacy_parent = _template_to_shell_and_chain(
            level_name,
            level_cfg.format,
            level_cfg.template,
            shell_catalog=derived.shell_catalog,
            chain_catalog=derived.chain_catalog,
        )

        binding = derived.level_bindings.get(level_name)
        if binding is None:
            binding = HeadingLevelBindingConfig()

        binding.enabled = True
        binding.display_shell = display_shell
        binding.display_core_style = display_core_style
        binding.reference_core_style = display_core_style
        binding.chain = chain_id
        binding.title_separator = level_cfg.effective_separator
        binding.ooxml_separator_mode = "inline"
        binding.ooxml_suff = "nothing"
        derived.level_bindings[level_name] = binding

        if used_legacy_parent:
            for parent_idx in range(1, heading_level_index(level_name)):
                decimal_reference_levels.add(f"heading{parent_idx}")

    for level_name in decimal_reference_levels:
        binding = derived.level_bindings.get(level_name)
        if binding is not None and binding.enabled:
            binding.reference_core_style = "arabic"

    return derived


def _build_scene_config(
    data: dict,
    *,
    base_dir: Path | None = None,
    track_upgrade: bool = False,
) -> SceneConfig:
    """Build SceneConfig from in-memory dict."""
    if not isinstance(data, dict):
        raise ValueError("scene data must be a dict")

    payload = copy.deepcopy(data)
    base = Path(base_dir) if base_dir is not None else PRESETS_DIR
    base_resolved = base.resolve()

    format_file = payload.pop("format_file", None)
    if format_file:
        fmt_ref = Path(str(format_file))
        if fmt_ref.is_absolute():
            raise ValueError("format_file must be a relative path within preset directory")
        fmt_path = (base_resolved / fmt_ref).resolve()
        try:
            fmt_path.relative_to(base_resolved)
        except ValueError as exc:
            raise ValueError("format_file points outside preset directory") from exc
        if fmt_path.suffix.lower() != ".json":
            raise ValueError("format_file must be a .json file")
        with open(fmt_path, "r", encoding="utf-8") as f:
            fmt_data = json.load(f)
        payload = _deep_merge(fmt_data, payload)

    original_payload = copy.deepcopy(payload)
    payload = _migrate_legacy_keys(payload)
    upgrade_notes = (
        _collect_scene_upgrade_notes(original_payload, payload)
        if track_upgrade else []
    )

    config = SceneConfig()
    config.version = payload.get("version", "1.0")
    config.name = payload.get("name", "")
    config.description = payload.get("description", "")
    config.format_signature = str(payload.get("format_signature", "") or "")
    config.category = payload.get("category", config.category)
    config.category_label = payload.get("category_label", config.category_label)
    if isinstance(payload.get("capabilities"), dict):
        merged_caps = dict(config.capabilities)
        for k, v in payload["capabilities"].items():
            merged_caps[str(k)] = bool(v)
        config.capabilities = merged_caps
    if isinstance(payload.get("available_sections"), list):
        uniq_sections = []
        for sec in payload["available_sections"]:
            sec_name = str(sec).strip()
            if sec_name == "cover":
                continue
            if sec_name and sec_name not in uniq_sections:
                uniq_sections.append(sec_name)
        if uniq_sections:
            config.available_sections = uniq_sections

    layout_mode = str(payload.get("normal_table_layout_mode", config.normal_table_layout_mode)).strip().lower()
    if layout_mode in {"smart", "compact", "full"}:
        config.normal_table_layout_mode = layout_mode
    smart_levels = payload.get("normal_table_smart_levels", config.normal_table_smart_levels)
    try:
        smart_levels = int(smart_levels)
    except (TypeError, ValueError):
        smart_levels = config.normal_table_smart_levels
    if smart_levels in {3, 4, 5, 6}:
        config.normal_table_smart_levels = smart_levels
    border_mode = str(payload.get("normal_table_border_mode", config.normal_table_border_mode)).strip().lower()
    if border_mode in {"full_grid", "three_line", "keep"}:
        config.normal_table_border_mode = border_mode
    # 线宽（磅）
    for attr in ("table_border_width_pt", "three_line_header_width_pt", "three_line_bottom_width_pt"):
        raw = payload.get(attr)
        if raw is not None:
            try:
                setattr(config, attr, max(0.25, min(float(raw), 3.0)))
            except (TypeError, ValueError):
                pass
    line_spacing_mode = str(
        payload.get("normal_table_line_spacing_mode", config.normal_table_line_spacing_mode)
    ).strip().lower()
    if line_spacing_mode in {"single", "one_half", "double"}:
        config.normal_table_line_spacing_mode = line_spacing_mode
    config.normal_table_repeat_header = bool(
        payload.get("normal_table_repeat_header", config.normal_table_repeat_header)
    )

    if "page_setup" in payload:
        ps = payload["page_setup"]
        config.page_setup = PageSetupConfig(
            paper_size=ps.get("paper_size", "A4"),
            margin=_parse_margin(ps.get("margin", {})),
            gutter_cm=ps.get("gutter_cm", 0),
            header_distance_cm=ps.get("header_distance_cm", 3.0),
            footer_distance_cm=ps.get("footer_distance_cm", 3.0),
        )

    if "heading_numbering" in payload:
        hn = payload["heading_numbering"]
        schemes = {}
        for scheme_id, scheme_levels in hn.get("schemes", {}).items():
            schemes[scheme_id] = {
                ln: _parse_heading_level(ld)
                for ln, ld in scheme_levels.items()
            }
        levels = {}
        for level_name, level_data in hn.get("levels", {}).items():
            levels[level_name] = _parse_heading_level(level_data)
        enf = hn.get("enforcement", {})
        rg = hn.get("risk_guard", {})
        config.heading_numbering = HeadingNumberingConfig(
            mode=hn.get("mode", "B"),
            scheme=hn.get("scheme", "1"),
            levels=levels,
            schemes=schemes,
            enforcement=EnforcementConfig(**{
                k: enf[k] for k in EnforcementConfig.__dataclass_fields__ if k in enf
            }),
            risk_guard=HeadingRiskGuardConfig(**{
                k: rg[k] for k in HeadingRiskGuardConfig.__dataclass_fields__ if k in rg
            }),
        )
        if not config.heading_numbering.levels:
            config.heading_numbering.apply_scheme()
    if "heading_numbering_v2" in payload:
        config.heading_numbering_v2 = _parse_heading_numbering_v2(payload["heading_numbering_v2"])
        config.heading_numbering.levels = _project_legacy_heading_levels_from_v2(
            config.heading_numbering_v2,
            existing_levels=config.heading_numbering.levels,
        )
        config._heading_numbering_v2_source = "payload"  # type: ignore[attr-defined]
    elif "heading_numbering" in payload:
        config.heading_numbering_v2 = _derive_heading_numbering_v2_from_legacy(
            config.heading_numbering,
            enabled=bool(config.capabilities.get("heading_numbering", True)),
        )
        config._heading_numbering_v2_source = "derived"  # type: ignore[attr-defined]
    else:
        config._heading_numbering_v2_source = "default"  # type: ignore[attr-defined]

    if "heading_model" in payload:
        config.heading_model = _parse_heading_model(payload["heading_model"])

    if "toc" in payload:
        toc_raw = payload["toc"]
        if not isinstance(toc_raw, dict):
            toc_raw = {}
        default_toc = TocConfig()
        toc_mode = str(toc_raw.get("mode", default_toc.mode)).strip().lower()
        if toc_mode in {"plain", "legacy"}:
            toc_mode = "plain"
        elif toc_mode in {"word_native", "native", "auto", "field"}:
            toc_mode = "word_native"
        else:
            toc_mode = default_toc.mode
        config.toc = TocConfig(mode=toc_mode)

    if "caption" in payload:
        cap = payload["caption"]
        config.caption = CaptionConfig(**{
            k: cap[k] for k in CaptionConfig.__dataclass_fields__ if k in cap
        })

    if "chem_typography" in payload:
        chem = payload["chem_typography"]
        default_cfg = ChemTypographyConfig()
        default_scopes = default_cfg.scopes
        scopes = chem.get("scopes", {})
        if isinstance(scopes, dict):
            merged_scopes = {**default_scopes, **scopes}
        else:
            merged_scopes = dict(default_scopes)

        allow_tokens = chem.get("allow_tokens", default_cfg.allow_tokens)
        if not isinstance(allow_tokens, list):
            allow_tokens = list(default_cfg.allow_tokens)
        else:
            allow_tokens = [str(v) for v in allow_tokens if str(v).strip()]

        allow_patterns = chem.get("allow_patterns", default_cfg.allow_patterns)
        if not isinstance(allow_patterns, list):
            allow_patterns = list(default_cfg.allow_patterns)
        else:
            allow_patterns = [str(v) for v in allow_patterns if str(v).strip()]

        ignore_tokens = chem.get("ignore_tokens", default_cfg.ignore_tokens)
        if not isinstance(ignore_tokens, list):
            ignore_tokens = list(default_cfg.ignore_tokens)
        else:
            ignore_tokens = [str(v) for v in ignore_tokens if str(v).strip()]

        ignore_patterns = chem.get("ignore_patterns", default_cfg.ignore_patterns)
        if not isinstance(ignore_patterns, list):
            ignore_patterns = list(default_cfg.ignore_patterns)
        else:
            ignore_patterns = [str(v) for v in ignore_patterns if str(v).strip()]

        overrides_raw = chem.get("manual_overrides", default_cfg.manual_overrides)
        manual_overrides: dict[str, str] = {}
        if isinstance(overrides_raw, dict):
            for k, v in overrides_raw.items():
                key = str(k) if k is not None else ""
                val = str(v) if v is not None else ""
                if key and val:
                    manual_overrides[key] = val
        config.chem_typography = ChemTypographyConfig(
            enabled=bool(chem.get("enabled", True)),
            scopes={k: bool(v) for k, v in merged_scopes.items()},
            allow_tokens=allow_tokens,
            allow_patterns=allow_patterns,
            ignore_tokens=ignore_tokens,
            ignore_patterns=ignore_patterns,
            manual_overrides=manual_overrides,
        )

    if "md_cleanup" in payload:
        raw = payload["md_cleanup"]
        if not isinstance(raw, dict):
            raw = {}
        default_cfg = MdCleanupConfig()
        valid_ordered_styles = {
            "mixed",
            "decimal_dot",
            "decimal_paren_right",
            "decimal_cn_dun",
            "decimal_full_paren",
        }
        valid_unordered_styles = {
            "word_default",
            "bullet_dot",
            "bullet_circle",
            "bullet_square",
            "bullet_dash",
        }
        list_marker_separator = str(
            raw.get("list_marker_separator", default_cfg.list_marker_separator)
        ).strip().lower()
        if list_marker_separator not in {"tab", "half_space", "full_space"}:
            list_marker_separator = default_cfg.list_marker_separator
        ordered_list_style = str(
            raw.get("ordered_list_style", default_cfg.ordered_list_style)
        ).strip().lower()
        if ordered_list_style not in valid_ordered_styles:
            ordered_list_style = default_cfg.ordered_list_style
        unordered_list_style = str(
            raw.get("unordered_list_style", default_cfg.unordered_list_style)
        ).strip().lower()
        if unordered_list_style not in valid_unordered_styles:
            unordered_list_style = default_cfg.unordered_list_style
        config.md_cleanup = MdCleanupConfig(
            enabled=bool(raw.get("enabled", default_cfg.enabled)),
            preserve_existing_word_lists=bool(
                raw.get(
                    "preserve_existing_word_lists",
                    default_cfg.preserve_existing_word_lists,
                )
            ),
            formula_copy_noise_cleanup=bool(
                raw.get(
                    "formula_copy_noise_cleanup",
                    default_cfg.formula_copy_noise_cleanup,
                )
            ),
            suppress_formula_fake_lists=bool(
                raw.get(
                    "suppress_formula_fake_lists",
                    default_cfg.suppress_formula_fake_lists,
                )
            ),
            list_marker_separator=list_marker_separator,
            ordered_list_style=ordered_list_style,
            unordered_list_style=unordered_list_style,
        )

    if "whitespace_normalize" in payload:
        ws = payload["whitespace_normalize"]
        if not isinstance(ws, dict):
            ws = {}
        default_ws = WhitespaceNormalizeConfig()
        try:
            confidence = int(ws.get("context_min_confidence", default_ws.context_min_confidence))
        except (TypeError, ValueError):
            confidence = default_ws.context_min_confidence
        confidence = max(1, min(confidence, 4))
        config.whitespace_normalize = WhitespaceNormalizeConfig(
            enabled=bool(ws.get("enabled", default_ws.enabled)),
            normalize_space_variants=bool(
                ws.get("normalize_space_variants", default_ws.normalize_space_variants)
            ),
            convert_tabs=bool(ws.get("convert_tabs", default_ws.convert_tabs)),
            remove_zero_width=bool(ws.get("remove_zero_width", default_ws.remove_zero_width)),
            collapse_multiple_spaces=bool(
                ws.get("collapse_multiple_spaces", default_ws.collapse_multiple_spaces)
            ),
            trim_paragraph_edges=bool(
                ws.get("trim_paragraph_edges", default_ws.trim_paragraph_edges)
            ),
            smart_full_half_convert=bool(
                ws.get("smart_full_half_convert", default_ws.smart_full_half_convert)
            ),
            punctuation_by_context=bool(
                ws.get("punctuation_by_context", default_ws.punctuation_by_context)
            ),
            bracket_by_inner_language=bool(
                ws.get("bracket_by_inner_language", default_ws.bracket_by_inner_language)
            ),
            fullwidth_alnum_to_halfwidth=bool(
                ws.get("fullwidth_alnum_to_halfwidth", default_ws.fullwidth_alnum_to_halfwidth)
            ),
            quote_by_context=bool(
                ws.get("quote_by_context", default_ws.quote_by_context)
            ),
            protect_reference_numbering=bool(
                ws.get("protect_reference_numbering", default_ws.protect_reference_numbering)
            ),
            context_min_confidence=confidence,
        )

    if "citation_link" in payload:
        cite = payload["citation_link"]
        default_cite = CitationLinkConfig()
        if not isinstance(cite, dict):
            cite = {}
        config.citation_link = CitationLinkConfig(
            enabled=bool(cite.get("enabled", default_cite.enabled)),
            auto_number_reference_entries=bool(
                cite.get(
                    "auto_number_reference_entries",
                    default_cite.auto_number_reference_entries,
                )
            ),
            superscript_outer_page_numbers=bool(
                cite.get(
                    "superscript_outer_page_numbers",
                    default_cite.superscript_outer_page_numbers,
                )
            ),
        )

    if "formula_convert" in payload:
        raw = payload["formula_convert"]
        if not isinstance(raw, dict):
            raw = {}
        default_cfg = FormulaConvertConfig()
        output_mode = str(raw.get("output_mode", default_cfg.output_mode)).strip().lower()
        if output_mode not in {"word_native", "latex"}:
            output_mode = default_cfg.output_mode
        low_policy = str(raw.get("low_confidence_policy", default_cfg.low_confidence_policy)).strip().lower()
        if low_policy != "skip_and_mark":
            low_policy = default_cfg.low_confidence_policy
        fallback_timeout = raw.get("office_fallback_timeout_sec", default_cfg.office_fallback_timeout_sec)
        try:
            fallback_timeout = int(fallback_timeout)
        except (TypeError, ValueError):
            fallback_timeout = default_cfg.office_fallback_timeout_sec
        fallback_timeout = max(10, min(fallback_timeout, 600))
        config.formula_convert = FormulaConvertConfig(
            enabled=bool(raw.get("enabled", default_cfg.enabled)),
            output_mode=output_mode,
            low_confidence_policy=low_policy,
            office_fallback_enabled=bool(raw.get("office_fallback_enabled", default_cfg.office_fallback_enabled)),
            office_fallback_timeout_sec=fallback_timeout,
        )

    if "formula_to_table" in payload:
        raw = payload["formula_to_table"]
        if not isinstance(raw, dict):
            raw = {}
        default_cfg = FormulaToTableConfig()
        config.formula_to_table = FormulaToTableConfig(
            enabled=bool(raw.get("enabled", default_cfg.enabled)),
            block_only=bool(raw.get("block_only", default_cfg.block_only)),
        )

    if "equation_table_format" in payload:
        raw = payload["equation_table_format"]
        if not isinstance(raw, dict):
            raw = {}
        default_cfg = EquationTableFormatConfig()
        numbering_format = str(
            raw.get("numbering_format", default_cfg.numbering_format)
        ).strip().lower()
        if numbering_format not in {"seq", "chapter-seq", "chapter.seq"}:
            numbering_format = default_cfg.numbering_format
        config.equation_table_format = EquationTableFormatConfig(
            enabled=bool(raw.get("enabled", default_cfg.enabled)),
            numbering_format=numbering_format,
        )

    if "formula_style" in payload:
        raw = payload["formula_style"]
        if not isinstance(raw, dict):
            raw = {}
        default_cfg = FormulaStyleConfig()
        config.formula_style = FormulaStyleConfig(
            enabled=bool(raw.get("enabled", default_cfg.enabled)),
            unify_font=bool(raw.get("unify_font", default_cfg.unify_font)),
            unify_size=bool(raw.get("unify_size", default_cfg.unify_size)),
            unify_spacing=bool(raw.get("unify_spacing", default_cfg.unify_spacing)),
        )

    if "formula_table" in payload:
        raw = payload["formula_table"]
        if not isinstance(raw, dict):
            raw = {}
        default_cfg = FormulaTableConfig()
        config.formula_table = FormulaTableConfig(
            formula_font_name=str(
                raw.get("formula_font_name", default_cfg.formula_font_name)
            ).strip() or default_cfg.formula_font_name,
            formula_font_size_pt=float(
                raw.get("formula_font_size_pt", default_cfg.formula_font_size_pt)
            ),
            formula_font_size_display=str(
                raw.get(
                    "formula_font_size_display",
                    default_cfg.formula_font_size_display,
                )
            ).strip(),
            formula_line_spacing=float(
                raw.get("formula_line_spacing", default_cfg.formula_line_spacing)
            ),
            formula_space_before_pt=float(
                raw.get("formula_space_before_pt", default_cfg.formula_space_before_pt)
            ),
            formula_space_after_pt=float(
                raw.get("formula_space_after_pt", default_cfg.formula_space_after_pt)
            ),
            block_alignment=str(
                raw.get("block_alignment", default_cfg.block_alignment)
            ).strip().lower() or default_cfg.block_alignment,
            table_alignment=str(
                raw.get("table_alignment", default_cfg.table_alignment)
            ).strip().lower() or default_cfg.table_alignment,
            formula_cell_alignment=str(
                raw.get("formula_cell_alignment", default_cfg.formula_cell_alignment)
            ).strip().lower() or default_cfg.formula_cell_alignment,
            number_alignment=str(
                raw.get("number_alignment", default_cfg.number_alignment)
            ).strip().lower() or default_cfg.number_alignment,
            number_font_name=str(
                raw.get("number_font_name", default_cfg.number_font_name)
            ).strip() or default_cfg.number_font_name,
            number_font_size_pt=float(
                raw.get("number_font_size_pt", default_cfg.number_font_size_pt)
            ),
            number_font_size_display=str(
                raw.get(
                    "number_font_size_display",
                    default_cfg.number_font_size_display,
                )
            ).strip(),
            auto_shrink_number_column=bool(
                raw.get(
                    "auto_shrink_number_column",
                    default_cfg.auto_shrink_number_column,
                )
            ),
        )

    if "format_scope" in payload:
        fs = payload["format_scope"]
        default_sections = FormatScopeConfig().sections
        sections = {
            str(key): bool(value)
            for key, value in (fs.get("sections", {}) or {}).items()
            if str(key).strip() != "cover"
        }
        merged = {**default_sections, **sections}
        config.format_scope = FormatScopeConfig(
            mode=fs.get("mode", "auto"),
            page_ranges_text=str(fs.get("page_ranges_text", "") or ""),
            body_start_index=fs.get("body_start_index"),
            body_start_page=fs.get("body_start_page"),
            body_start_keyword=fs.get("body_start_keyword", ""),
            sections=merged,
        )

    if "styles" in payload:
        for style_name, style_data in payload["styles"].items():
            config.styles[style_name] = _parse_style(style_data)

    _STANDARD_STYLES = [
        "normal",
        "heading1", "heading2", "heading3", "heading4", "heading5", "heading6", "heading7", "heading8",
        "abstract_title_cn", "abstract_title_en", "abstract_body", "abstract_body_en",
        "toc_title", "toc_chapter", "toc_level1", "toc_level2",
        "references_body", "acknowledgment_body", "appendix_body", "resume_body", "symbol_table_body",
        "code_block", "figure_caption", "table_caption", "header_cn", "header_en", "page_number",
    ]

    _backfilled = set()
    for skey in _STANDARD_STYLES:
        if skey not in config.styles:
            config.styles[skey] = StyleConfig()
            _backfilled.add(skey)

    # Reorder starting with the standard order, then append any custom styles
    _ordered = {}
    for skey in _STANDARD_STYLES:
        if skey in config.styles:
            _ordered[skey] = config.styles.get(skey)
    for k, v in config.styles.items():
        if k not in _STANDARD_STYLES:
            _ordered[k] = v

    config.styles = _ordered
    for style_cfg in config.styles.values():
        if isinstance(style_cfg, StyleConfig):
            sync_style_config_indent_fields(style_cfg)
    config._backfilled_styles = _backfilled  # type: ignore[attr-defined]

    if "output" in payload:
        o = payload["output"]
        config.output = OutputConfig(**{
            k: o[k] for k in OutputConfig.__dataclass_fields__ if k in o
        })

    if "pipeline" in payload:
        pipeline = _normalize_pipeline_steps(payload["pipeline"])
        config.pipeline = pipeline
    # Backward compatibility: if legacy scenes only store optional steps in pipeline,
    # infer corresponding enabled flags before strict pipeline<->flag synchronization.
    for cfg_key in (
        "formula_convert",
        "formula_to_table",
        "equation_table_format",
        "formula_style",
        "citation_link",
    ):
        if cfg_key in payload:
            continue
        cfg_obj = getattr(config, cfg_key, None)
        if cfg_obj is None:
            continue
        if cfg_key in (config.pipeline or []):
            setattr(cfg_obj, "enabled", True)
    if "pipeline_strict_mode" in payload:
        config.pipeline_strict_mode = bool(payload["pipeline_strict_mode"])
    critical_rules_payload = []
    if "pipeline_critical_rules" in payload and isinstance(payload["pipeline_critical_rules"], list):
        critical_rules_payload = _normalize_pipeline_steps(payload["pipeline_critical_rules"])
        if critical_rules_payload:
            config.pipeline_critical_rules = critical_rules_payload
    payload_pipeline = _normalize_pipeline_steps(payload.get("pipeline", []))
    config._pipeline_critical_rules_source = (
        "auto"
        if _is_auto_managed_critical_rules(critical_rules_payload, payload_pipeline)
        else "payload_custom"
    )  # type: ignore[attr-defined]

    for bool_key in ("update_header", "update_page_number", "update_header_line"):
        if bool_key in payload:
            setattr(config, bool_key, bool(payload[bool_key]))
    _sync_whitespace_pipeline_switch(config)
    _sync_formula_pipeline_switch(config)
    _sync_citation_link_pipeline_switch(config)
    _sync_pipeline_critical_rules(config)
    config._scene_upgrade_notes = upgrade_notes  # type: ignore[attr-defined]
    config._scene_upgrade_applied = bool(upgrade_notes)  # type: ignore[attr-defined]
    return config


def load_scene(path: Path) -> SceneConfig:
    """从 JSON 文件加载场景配置。"""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return _build_scene_config(data, base_dir=Path(path).parent, track_upgrade=True)


def load_scene_from_data(data: dict, *, base_dir: Path | None = None) -> SceneConfig:
    """Load SceneConfig from an in-memory dict."""
    return _build_scene_config(data, base_dir=base_dir, track_upgrade=False)


def get_scene_upgrade_notes(config: SceneConfig) -> list[str]:
    notes = getattr(config, "_scene_upgrade_notes", None)
    if not isinstance(notes, list):
        return []
    result: list[str] = []
    for note in notes:
        text = str(note or "").strip()
        if text:
            result.append(text)
    return result


def list_presets() -> list[tuple[str, Path, str, str]]:
    """列出所有内置预设，返回 [(名称, 路径, 分类id, 分类名), ...]。"""
    results = []
    for p in PRESETS_DIR.glob("*.json"):
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            results.append((
                data.get("name", p.stem),
                p,
                data.get("category", "general"),
                data.get("category_label", "通用文档"),
            ))
        except Exception:
            continue
    results.sort(key=lambda x: (x[3], x[0]))
    return results


def load_preset(name: str) -> SceneConfig | None:
    """Load built-in preset by display name or file stem."""
    for preset_name, path, _, _ in list_presets():
        if preset_name == name or path.stem == name:
            return load_scene(path)
    return None


def default_scene_path() -> Path:
    """Return the canonical path for the default scene preset."""
    for candidate in _default_scene_candidate_paths():
        try:
            if candidate.exists() and candidate.is_file():
                return candidate
        except Exception:
            continue
    raise FileNotFoundError("默认场景模板 default_format.json 不存在")


def load_default_scene() -> SceneConfig:
    """Load a fresh SceneConfig from the default preset."""
    return load_scene(default_scene_path())


def is_protected_scene_path(path: Path) -> bool:
    """Return True when the path points to a protected built-in scene."""
    candidate = Path(path)
    try:
        candidate_resolved = candidate.resolve()
    except Exception:
        candidate_resolved = candidate

    for protected in _default_scene_candidate_paths():
        try:
            protected_resolved = protected.resolve()
        except Exception:
            protected_resolved = protected
        if candidate_resolved == protected_resolved:
            return True
    return False


def _atomic_write_json(path: Path, data: dict) -> None:
    """Write JSON via a sibling temp file and atomically replace the target."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)

    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=target.parent,
            prefix=f".{target.name}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temp_path = Path(handle.name)
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.flush()
            try:
                os.fsync(handle.fileno())
            except OSError:
                pass
        os.replace(temp_path, target)
    except Exception:
        if temp_path is not None:
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except Exception:
                pass
        raise


def save_scene(config: SceneConfig, path: Path) -> None:
    """将 SceneConfig 保存到 JSON 文件（增量保存）。

    保留 config 中的元数据字段（name, category 等）。
    """
    _sync_pipeline_critical_rules(config)
    for style_cfg in getattr(config, "styles", {}).values():
        if isinstance(style_cfg, StyleConfig):
            sync_style_config_indent_fields(style_cfg)
    data = asdict(config)
    if isinstance(data.get("available_sections"), list):
        data["available_sections"] = [
            sec for sec in data["available_sections"]
            if str(sec).strip() != "cover"
        ]
    format_scope = data.get("format_scope")
    if isinstance(format_scope, dict):
        sections = format_scope.get("sections")
        if isinstance(sections, dict):
            sections.pop("cover", None)
    # 清理 None 和内部字段
    for key in list(data.keys()):
        if data[key] is None:
            del data[key]
    _atomic_write_json(path, data)


def _safe_filename(name: str) -> str:
    """将场景名称转为安全文件名（不含扩展名）。

    替换文件系统非法字符为 '_'，去除首尾空白和点号。
    """
    import re
    # 替换 Windows / Unix 文件名非法字符
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name.strip())
    # 去除首尾的 '.' 和空白
    safe = safe.strip('. ')
    # 合并连续下划线
    safe = re.sub(r'_+', '_', safe)
    return safe or "unnamed"


def rename_scene(path: Path, new_name: str) -> Path:
    """修改场景文件中的 name 字段，并将文件重命名为与 name 一致的文件名。

    返回重命名后的新路径。如果目标文件名已存在则自动加数字后缀。
    """
    path = Path(path)
    if is_protected_scene_path(path):
        raise PermissionError("默认场景模板 default_format.json 不能重命名。")

    new_name = new_name.strip()
    if not new_name:
        raise ValueError("场景名称不能为空")

    # 更新 JSON 内 name 字段
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    data["name"] = new_name

    # 计算新文件名
    safe_stem = _safe_filename(new_name)
    parent = path.parent
    new_path = parent / f"{safe_stem}.json"

    # 如果新路径和旧路径相同（仅大小写差异等），直接原子写回
    if new_path.resolve() == path.resolve():
        _atomic_write_json(path, data)
        return path

    # 如果目标文件名已存在，添加数字后缀
    if new_path.exists():
        idx = 1
        while True:
            candidate = parent / f"{safe_stem}_{idx}.json"
            if not candidate.exists():
                new_path = candidate
                break
            idx += 1

    # 先原子写入新文件，再删除旧文件；写入失败时保留旧文件不动
    _atomic_write_json(new_path, data)
    if path.exists() and path.resolve() != new_path.resolve():
        path.unlink()

    return new_path


def delete_scene(path: Path) -> None:
    """删除指定的场景 JSON 文件。"""
    path = Path(path)
    if is_protected_scene_path(path):
        raise PermissionError("默认场景模板 default_format.json 不能删除。")
    if path.exists():
        path.unlink()
