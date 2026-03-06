"""鍦烘櫙鍔犺浇/淇濆瓨/瀵煎叆瀵煎嚭"""

import copy
import json
import os
import sys
from pathlib import Path
from src.scene.schema import (
    SceneConfig, PageSetupConfig, MarginConfig,
    HeadingNumberingConfig, HeadingLevelConfig,
    HeadingModelConfig,
    EnforcementConfig, HeadingRiskGuardConfig,
    StyleConfig, OutputConfig,
    FormatScopeConfig, CaptionConfig, ChemTypographyConfig,
    WhitespaceNormalizeConfig, CitationLinkConfig,
)


def _resolve_presets_dir() -> Path:
    """Resolve built-in preset directory (read-only seed source)."""
    if getattr(sys, "frozen", False):
        # PyInstaller onefile/onedir extraction root
        meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        exe_dir = Path(sys.executable).resolve().parent
        candidates = [
            meipass / "src" / "scene" / "presets",
            meipass / "templates",
            exe_dir / "src" / "scene" / "presets",
            exe_dir / "presets",
            exe_dir / "templates",
        ]
        for p in candidates:
            if p.exists():
                return p

    source_dir = Path(__file__).resolve().parent / "presets"
    if source_dir.exists():
        return source_dir

    return source_dir


def _resolve_templates_dir() -> Path:
    """Resolve user-writable templates directory next to the executable.

    In frozen (packaged) mode: <exe_dir>/templates
    In source mode: <project_root>/src/scene/presets (fallback to built-in)
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

    # 开发模式默认使用 presets 目录
    return _resolve_presets_dir()


def _ensure_templates(builtin_dir: Path, templates_dir: Path) -> None:
    """Copy built-in presets into the user-writable templates folder.

    Only copies files that do not already exist in templates_dir,
    so user modifications are preserved.
    """
    import shutil
    if not builtin_dir.exists():
        return
    templates_dir.mkdir(parents=True, exist_ok=True)
    for src_file in builtin_dir.glob("*.json"):
        dst_file = templates_dir / src_file.name
        if not dst_file.exists():
            shutil.copy2(src_file, dst_file)
    # Also copy subdirectories (e.g. formats/)
    for src_sub in builtin_dir.iterdir():
        if src_sub.is_dir():
            dst_sub = templates_dir / src_sub.name
            if not dst_sub.exists():
                shutil.copytree(src_sub, dst_sub)


_BUILTIN_PRESETS_DIR = _resolve_presets_dir()
PRESETS_DIR = _resolve_templates_dir()

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


def _deep_merge(base: dict, override: dict) -> dict:
    """娣卞害鍚堝苟锛宱verride 涓殑鍊艰鐩?base"""
    result = base.copy()
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


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
                    v = _migrate_str_value(v)
                result[k] = v
            return result

        # 有旧 key → 需要重命名
        # 使用临时 key 避免碰撞
        temp = {}
        for k, v in processed.items():
            if isinstance(v, str):
                v = _migrate_str_value(v)
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

    def _migrate_str_value(v: str) -> str:
        return _STR_VALUE_MAP.get(v, v)

    return _migrate_dict(data)


def _build_scene_config(data: dict, *, base_dir: Path | None = None) -> SceneConfig:
    """Build SceneConfig from in-memory dict."""
    if not isinstance(data, dict):
        raise ValueError("scene data must be a dict")

    payload = copy.deepcopy(data)
    payload = _migrate_legacy_keys(payload)
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

    config = SceneConfig()
    config.version = payload.get("version", "1.0")
    config.name = payload.get("name", "")
    config.description = payload.get("description", "")
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

    if "heading_model" in payload:
        config.heading_model = _parse_heading_model(payload["heading_model"])

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

    if "whitespace_normalize" in payload:
        ws = payload["whitespace_normalize"]
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

    if "format_scope" in payload:
        fs = payload["format_scope"]
        default_sections = FormatScopeConfig().sections
        sections = fs.get("sections", {})
        merged = {**default_sections, **sections}
        config.format_scope = FormatScopeConfig(
            mode=fs.get("mode", "auto"),
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
    config._backfilled_styles = _backfilled  # type: ignore[attr-defined]

    if "output" in payload:
        o = payload["output"]
        config.output = OutputConfig(**{
            k: o[k] for k in OutputConfig.__dataclass_fields__ if k in o
        })

    if "pipeline" in payload:
        config.pipeline = payload["pipeline"]
    if "pipeline_strict_mode" in payload:
        config.pipeline_strict_mode = bool(payload["pipeline_strict_mode"])
    if "pipeline_critical_rules" in payload and isinstance(payload["pipeline_critical_rules"], list):
        configured = []
        for name in payload["pipeline_critical_rules"]:
            text = str(name).strip()
            if text:
                configured.append(text)
        if configured:
            config.pipeline_critical_rules = configured

    for bool_key in ("update_header", "update_page_number", "update_header_line"):
        if bool_key in payload:
            setattr(config, bool_key, bool(payload[bool_key]))

    return config


def load_scene(path: Path) -> SceneConfig:
    """浠?JSON 鏂囦欢鍔犺浇鍦烘櫙閰嶇疆"""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return _build_scene_config(data, base_dir=Path(path).parent)


def load_scene_from_data(data: dict, *, base_dir: Path | None = None) -> SceneConfig:
    """Load SceneConfig from an in-memory dict."""
    return _build_scene_config(data, base_dir=base_dir)


def list_presets() -> list[tuple[str, Path, str, str]]:
    """鍒楀嚭鎵€鏈夊唴缃璁撅紝杩斿洖 [(鍚嶇О, 璺緞, 鍒嗙被id, 鍒嗙被鍚?, ...]"""
    results = []
    for p in PRESETS_DIR.glob("*.json"):
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            results.append((
                data.get("name", p.stem),
                p,
                data.get("category", "general"),
                data.get("category_label", "閫氱敤鏂囨。"),
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


def save_scene(config: SceneConfig, path: Path) -> None:
    """将 SceneConfig 保存到 JSON 文件（增量保存）。

    保留 config 中的元数据字段（name, category 等）。
    """
    from dataclasses import asdict
    data = asdict(config)
    # 清理 None 和内部字段
    for key in list(data.keys()):
        if data[key] is None:
            del data[key]
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


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

    # 如果新路径和旧路径相同（仅大小写差异等），直接写回
    if new_path.resolve() == path.resolve():
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
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

    # 写入新文件，删除旧文件
    with open(new_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    if path.exists() and path.resolve() != new_path.resolve():
        path.unlink()

    return new_path


def delete_scene(path: Path) -> None:
    """删除指定的场景 JSON 文件。"""
    path = Path(path)
    if path.exists():
        path.unlink()
