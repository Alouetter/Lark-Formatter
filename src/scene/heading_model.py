"""Heading model helpers shared by rules to avoid scattered hard-coding."""

from __future__ import annotations

import re

from src.scene.schema import HeadingModelConfig, SceneConfig


def _norm_no_space(text: str) -> str:
    return re.sub(r"\s+", "", text or "").strip()


def _as_str_dict(data, *, lower_values: bool = False) -> dict[str, str]:
    if not isinstance(data, dict):
        return {}
    result: dict[str, str] = {}
    for raw_k, raw_v in data.items():
        key = str(raw_k or "").strip()
        val = str(raw_v or "").strip()
        if not key or not val:
            continue
        result[key] = val.lower() if lower_values else val
    return result


def _as_str_list(data, *, lower: bool = False) -> list[str]:
    if not isinstance(data, list):
        return []
    result: list[str] = []
    for raw in data:
        text = str(raw or "").strip()
        if not text:
            continue
        item = text.lower() if lower else text
        if item not in result:
            result.append(item)
    return result


def _model(config: SceneConfig) -> HeadingModelConfig:
    hm = getattr(config, "heading_model", None)
    if isinstance(hm, HeadingModelConfig):
        return hm
    return HeadingModelConfig()


def get_level_to_word_style(config: SceneConfig) -> dict[str, str]:
    default = HeadingModelConfig().level_to_word_style
    merged = dict(default)
    merged.update(_as_str_dict(_model(config).level_to_word_style))
    return merged


def get_level_to_style_key(config: SceneConfig) -> dict[str, str]:
    default = HeadingModelConfig().level_to_style_key
    merged = dict(default)
    merged.update(_as_str_dict(_model(config).level_to_style_key))
    return merged


def get_style_alias_to_level(config: SceneConfig) -> dict[str, str]:
    default = HeadingModelConfig().style_alias_to_level
    merged = dict(default)
    merged.update(_as_str_dict(_model(config).style_alias_to_level, lower_values=True))
    return merged


def get_section_title_style_map(config: SceneConfig) -> dict[str, str]:
    default = HeadingModelConfig().section_title_style_map
    merged = dict(default)
    merged.update(_as_str_dict(_model(config).section_title_style_map))
    return merged


def get_non_numbered_title_sections(config: SceneConfig) -> set[str]:
    raw = _model(config).non_numbered_title_sections
    if isinstance(raw, list):
        return set(_as_str_list(raw, lower=True))
    return set(_as_str_list(HeadingModelConfig().non_numbered_title_sections, lower=True))


def get_non_numbered_title_norms(config: SceneConfig) -> set[str]:
    raw = _model(config).non_numbered_title_texts
    if isinstance(raw, list):
        values = _as_str_list(raw)
    else:
        values = HeadingModelConfig().non_numbered_title_texts
    return {_norm_no_space(v) for v in values if _norm_no_space(v)}


def get_front_matter_title_norms(config: SceneConfig) -> set[str]:
    raw = _model(config).front_matter_title_texts
    if isinstance(raw, list):
        values = _as_str_list(raw)
    else:
        values = HeadingModelConfig().front_matter_title_texts
    return {_norm_no_space(v).lower() for v in values if _norm_no_space(v)}


def get_post_section_types(config: SceneConfig) -> set[str]:
    raw = _model(config).post_section_types
    if isinstance(raw, list):
        return set(_as_str_list(raw, lower=True))
    return set(_as_str_list(HeadingModelConfig().post_section_types, lower=True))


def get_non_numbered_heading_style_name(config: SceneConfig) -> str:
    configured = str(_model(config).non_numbered_heading_style_name or "").strip()
    if configured:
        return configured
    return HeadingModelConfig().non_numbered_heading_style_name


def get_header_front_text(config: SceneConfig) -> dict[str, str]:
    default = HeadingModelConfig().header_front_text
    merged = dict(default)
    merged.update(_as_str_dict(_model(config).header_front_text))
    return merged


def get_header_back_types(config: SceneConfig) -> set[str]:
    raw = _model(config).header_back_types
    if isinstance(raw, list):
        return set(_as_str_list(raw, lower=True))
    return set(_as_str_list(HeadingModelConfig().header_back_types, lower=True))


def get_style_name_to_level(config: SceneConfig) -> dict[str, str]:
    """Build style-name lookup table (lowercased) for heading detection."""
    result: dict[str, str] = {}
    for level, style_name in get_level_to_word_style(config).items():
        s = str(style_name or "").strip().lower()
        if s:
            result[s] = str(level).strip().lower()
    for style_name, level in get_style_alias_to_level(config).items():
        s = str(style_name or "").strip().lower()
        lv = str(level or "").strip().lower()
        if s and lv:
            result[s] = lv
    return result


def detect_level_by_style_name(config: SceneConfig, style_name: str) -> str | None:
    s = str(style_name or "").strip()
    if not s:
        return None
    lower = s.lower()
    mapping = get_style_name_to_level(config)
    if lower in mapping:
        return mapping[lower]
    # Prefer longer keys first so aliases like "heading 1 char" can override
    # generic prefixes like "heading 1".
    for name_key, level in sorted(mapping.items(), key=lambda item: len(item[0]), reverse=True):
        if lower.startswith(name_key):
            tail = lower[len(name_key):]
            # Guard against false match: "heading 10" should not map to "heading 1".
            if tail and name_key[-1:].isdigit() and tail[0].isdigit():
                continue
            return level
    return None
