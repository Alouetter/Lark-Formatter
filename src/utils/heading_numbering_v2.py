"""Helpers for projecting heading_numbering_v2 into runtime structures."""

from __future__ import annotations

import copy

from src.scene.schema import (
    HeadingLevelBindingConfig,
    HeadingLevelConfig,
    HeadingNumberingV2Config,
    heading_level_index,
    heading_level_keys,
)
from src.utils.chinese import int_to_chinese


_DEFAULT_V2_BINDINGS = HeadingNumberingV2Config().level_bindings
_CN_UPPER_MAP = str.maketrans(
    {
        "一": "壹",
        "二": "贰",
        "三": "叁",
        "四": "肆",
        "五": "伍",
        "六": "陆",
        "七": "柒",
        "八": "捌",
        "九": "玖",
        "十": "拾",
        "百": "佰",
        "千": "仟",
        "零": "零",
        "两": "贰",
    }
)
_CIRCLED_DIGITS = {
    idx: ch
    for idx, ch in enumerate("①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳", start=1)
}
_CIRCLED_PAREN_DIGITS = {
    idx: ch
    for idx, ch in enumerate("⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇", start=1)
}


def _default_binding(level_name: str) -> HeadingLevelBindingConfig:
    return copy.deepcopy(
        _DEFAULT_V2_BINDINGS.get(level_name, HeadingLevelBindingConfig())
    )


def merged_level_binding(
    level_name: str,
    level_bindings: dict[str, HeadingLevelBindingConfig] | None,
) -> HeadingLevelBindingConfig:
    binding = _default_binding(level_name)
    incoming = (level_bindings or {}).get(level_name)
    if incoming is None:
        return binding
    return copy.deepcopy(incoming)


def _to_roman(value: int) -> str:
    if value <= 0:
        return ""
    mapping = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ]
    remaining = value
    parts: list[str] = []
    for threshold, symbol in mapping:
        while remaining >= threshold:
            parts.append(symbol)
            remaining -= threshold
    return "".join(parts)


def format_core_style_value(core_style_id: str, value: int) -> str:
    normalized = str(core_style_id or "").strip().lower()
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = 0
    if number <= 0:
        return ""

    if normalized == "arabic":
        return str(number)
    if normalized == "arabic_pad2":
        return f"{number:02d}"
    if normalized == "cn_lower":
        return int_to_chinese(number)
    if normalized == "cn_upper":
        return int_to_chinese(number).translate(_CN_UPPER_MAP)
    if normalized == "roman_upper":
        return _to_roman(number)
    if normalized == "circled":
        return _CIRCLED_DIGITS.get(number, str(number))
    if normalized == "circled_paren":
        return _CIRCLED_PAREN_DIGITS.get(number, f"({number})")
    return str(number)


def explicit_chain_template_from_binding(
    level_name: str,
    binding: HeadingLevelBindingConfig,
    v2_config: HeadingNumberingV2Config,
) -> str:
    chain = (v2_config.chain_catalog or {}).get(str(binding.chain or "current_only"))
    segments = getattr(chain, "segments", None) or []
    parts: list[str] = []
    max_parent = heading_level_index(level_name) - 1
    for segment in segments:
        segment_type = str(getattr(segment, "type", "") or "")
        if segment_type == "literal":
            parts.append(str(getattr(segment, "text", "") or ""))
            continue

        source = str(getattr(segment, "source", "") or "")
        if source == "current":
            parts.append("{current}")
            continue
        if source.startswith("level"):
            try:
                level_idx = int(source[5:])
            except ValueError:
                continue
            if 1 <= level_idx <= max_parent:
                parts.append(f"{{{source}}}")

    return "".join(parts) or "{current}"


def uses_decimal_parent_chain(
    level_name: str,
    binding: HeadingLevelBindingConfig,
    v2_config: HeadingNumberingV2Config,
) -> bool:
    level_idx = heading_level_index(level_name)
    if level_idx <= 1:
        return False
    if str(binding.display_core_style or "arabic") != "arabic":
        return False
    expected = ".".join(
        [f"{{level{idx}}}" for idx in range(1, level_idx)] + ["{current}"]
    )
    return (
        explicit_chain_template_from_binding(level_name, binding, v2_config) == expected
    )


def legacy_format_from_binding(binding: HeadingLevelBindingConfig) -> str:
    shell_id = str(binding.display_shell or "plain")
    core_style_id = str(binding.display_core_style or "arabic")
    chain_id = str(binding.chain or "current_only")

    if chain_id == "current_only":
        if shell_id == "chapter_cn" and core_style_id == "cn_lower":
            return "chinese_chapter"
        if shell_id == "section_cn" and core_style_id == "cn_lower":
            return "chinese_section"
        if shell_id == "dunhao_cn" and core_style_id == "cn_lower":
            return "chinese_ordinal"
        if shell_id == "paren_cn" and core_style_id == "cn_lower":
            return "chinese_ordinal_paren"

    if chain_id != "current_only" and core_style_id == "arabic":
        return "arabic_dotted"

    return core_style_id or "arabic"


def legacy_chain_template_from_binding(
    level_name: str,
    binding: HeadingLevelBindingConfig,
    v2_config: HeadingNumberingV2Config,
) -> str:
    if uses_decimal_parent_chain(level_name, binding, v2_config):
        return "{parent}.{current}"
    return explicit_chain_template_from_binding(level_name, binding, v2_config)


def legacy_template_from_binding(
    level_name: str,
    binding: HeadingLevelBindingConfig,
    v2_config: HeadingNumberingV2Config,
) -> str:
    shell = (v2_config.shell_catalog or {}).get(str(binding.display_shell or "plain"))
    prefix = str(getattr(shell, "prefix", "") or "")
    suffix = str(getattr(shell, "suffix", "") or "")
    return (
        f"{prefix}"
        f"{legacy_chain_template_from_binding(level_name, binding, v2_config)}"
        f"{suffix}"
    )


def legacy_levels_from_v2(
    v2_config: HeadingNumberingV2Config,
    *,
    existing_levels: dict[str, HeadingLevelConfig] | None = None,
) -> dict[str, HeadingLevelConfig]:
    levels: dict[str, HeadingLevelConfig] = {}
    bindings = getattr(v2_config, "level_bindings", {}) or {}
    legacy_levels = existing_levels or {}
    for level_name in heading_level_keys():
        binding = bindings.get(level_name)
        if binding is None or not binding.enabled:
            continue
        existing = legacy_levels.get(level_name)
        if isinstance(existing, HeadingLevelConfig):
            projected = HeadingLevelConfig(**vars(existing))
        else:
            projected = HeadingLevelConfig()
        projected.format = legacy_format_from_binding(binding)
        projected.template = legacy_template_from_binding(level_name, binding, v2_config)
        projected.separator = str(binding.title_separator or "")
        projected.custom_separator = None
        levels[level_name] = projected
    return levels


def normalize_start_at(value: int) -> int:
    try:
        start_at = int(value)
    except (TypeError, ValueError):
        return 1
    return max(start_at, 1)


def lvl_restart_value(
    level_name: str,
    binding: HeadingLevelBindingConfig,
) -> int | None:
    level_idx = heading_level_index(level_name)
    if level_idx <= 1:
        return None

    restart_on = str(getattr(binding, "restart_on", "") or "").strip().lower()
    if not restart_on:
        return 0

    restart_idx = heading_level_index(restart_on)
    if restart_idx <= 0 or restart_idx >= level_idx:
        return None
    if restart_idx == level_idx - 1:
        return None
    return restart_idx


def advance_heading_counters(
    level_name: str,
    counters: dict[str, int],
    level_bindings: dict[str, HeadingLevelBindingConfig] | None,
) -> dict[str, int]:
    next_counters = {
        level: int((counters or {}).get(level, 0) or 0)
        for level in heading_level_keys()
    }
    current_idx = heading_level_index(level_name)

    for level in heading_level_keys():
        level_idx = heading_level_index(level)
        if level_idx <= current_idx:
            continue
        binding = merged_level_binding(level, level_bindings)
        restart_on = str(getattr(binding, "restart_on", "") or "").strip().lower()
        if not restart_on:
            continue
        restart_idx = heading_level_index(restart_on)
        if 1 <= restart_idx <= current_idx:
            next_counters[level] = 0

    binding = merged_level_binding(level_name, level_bindings)
    start_at = normalize_start_at(getattr(binding, "start_at", 1))
    current_value = int(next_counters.get(level_name, 0) or 0)
    next_counters[level_name] = start_at if current_value <= 0 else current_value + 1
    return next_counters


def included_toc_levels(
    level_bindings: dict[str, HeadingLevelBindingConfig] | None,
) -> list[str]:
    included: list[str] = []
    for level_name in heading_level_keys():
        binding = merged_level_binding(level_name, level_bindings)
        if bool(getattr(binding, "include_in_toc", True)):
            included.append(level_name)
    return included
