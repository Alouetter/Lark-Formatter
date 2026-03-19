"""Shared heading numbering template helpers."""

from __future__ import annotations

from dataclasses import dataclass
import re

from src.utils.chinese import int_to_chinese


_HEADING_LEVEL_RE = re.compile(r"^heading([1-8])$")
_TEMPLATE_TOKEN_RE = re.compile(r"\{(?P<name>[a-zA-Z][a-zA-Z0-9_]*)\}")
_LEVEL_TOKEN_RE = re.compile(r"^level([1-8])$")
_CURRENT_TOKEN_ALIASES = {"current", "n", "cn"}
_CHINESE_NUMBER_FORMATS = {
    "chinese_chapter",
    "chinese_section",
    "chinese_ordinal",
    "chinese_ordinal_paren",
}

SUPPORTED_TEMPLATE_TOKENS = tuple(
    ["{current}", "{parent}"] + [f"{{level{i}}}" for i in range(1, 9)]
)


@dataclass(frozen=True)
class HeadingTemplateToken:
    raw_name: str
    kind: str
    level_index: int | None = None


def heading_level_to_index(level_name: str) -> int | None:
    match = _HEADING_LEVEL_RE.match(str(level_name or "").strip().lower())
    if not match:
        return None
    return int(match.group(1))


def default_heading_numbering_template(level_name: str) -> str:
    level_index = heading_level_to_index(level_name) or 1
    if level_index <= 1:
        return "{current}"
    parent_tokens = [f"{{level{i}}}" for i in range(1, level_index)]
    return ".".join(parent_tokens + ["{current}"])


def _parse_token(raw_name: str) -> HeadingTemplateToken | None:
    name = str(raw_name or "").strip()
    if not name:
        return None
    if name in _CURRENT_TOKEN_ALIASES:
        return HeadingTemplateToken(raw_name=name, kind="current")
    if name == "parent":
        return HeadingTemplateToken(raw_name=name, kind="parent")
    match = _LEVEL_TOKEN_RE.match(name)
    if match:
        return HeadingTemplateToken(
            raw_name=name,
            kind="level",
            level_index=int(match.group(1)),
        )
    return None


def iter_heading_template_tokens(template: str) -> list[HeadingTemplateToken]:
    tokens: list[HeadingTemplateToken] = []
    for match in _TEMPLATE_TOKEN_RE.finditer(str(template or "")):
        token = _parse_token(match.group("name"))
        if token is not None:
            tokens.append(token)
    return tokens


def validate_heading_numbering_template(level_name: str, template: str) -> list[str]:
    errors: list[str] = []
    template_text = str(template or "")
    current_level_index = heading_level_to_index(level_name)

    for match in _TEMPLATE_TOKEN_RE.finditer(template_text):
        raw_name = match.group("name")
        token = _parse_token(raw_name)
        if token is None:
            errors.append(f"unsupported placeholder {{{raw_name}}}")
            continue
        if token.kind == "parent" and current_level_index == 1:
            errors.append("{parent} is not valid for heading1")
            continue
        if (
            token.kind == "level"
            and current_level_index is not None
            and token.level_index is not None
            and token.level_index > current_level_index
        ):
            errors.append(
                f"{{level{token.level_index}}} cannot be used in {level_name}"
            )

    stripped = _TEMPLATE_TOKEN_RE.sub("", template_text)
    if "{" in stripped or "}" in stripped:
        errors.append("template contains unmatched braces")

    return errors


def template_uses_legacy_parent(level_name: str, template: str) -> bool:
    del level_name
    return any(token.kind == "parent" for token in iter_heading_template_tokens(template))


def build_ooxml_lvl_text(level_name: str, template: str) -> str:
    current_level_index = heading_level_to_index(level_name) or 1
    template_text = str(template or "")

    def _replace(match: re.Match[str]) -> str:
        token = _parse_token(match.group("name"))
        if token is None:
            return match.group(0)
        if token.kind == "parent":
            refs = [f"%{idx}" for idx in range(1, current_level_index)]
            return ".".join(refs)
        if token.kind == "current":
            return f"%{current_level_index}"
        if token.level_index is None:
            return match.group(0)
        return f"%{token.level_index}"

    return _TEMPLATE_TOKEN_RE.sub(_replace, template_text)


def _format_counter_value(format_name: str, value: int) -> str:
    if value <= 0:
        return ""
    if str(format_name or "").strip().lower() in _CHINESE_NUMBER_FORMATS:
        return int_to_chinese(value)
    return str(value)


def _level_format(level_name: str, levels_config: dict) -> str:
    level_cfg = (levels_config or {}).get(level_name)
    if level_cfg is None:
        return "arabic"
    return str(getattr(level_cfg, "format", "") or "arabic")


def _counter_value(counters: dict[str, int], level_name: str) -> int:
    raw = (counters or {}).get(level_name, 0)
    try:
        return int(raw)
    except (TypeError, ValueError):
        return 0


def render_heading_numbering(
    level_name: str,
    template: str,
    levels_config: dict,
    counters: dict[str, int],
) -> str:
    current_level_index = heading_level_to_index(level_name) or 1
    template_text = str(template or "")

    def _replace(match: re.Match[str]) -> str:
        token = _parse_token(match.group("name"))
        if token is None:
            return match.group(0)
        if token.kind == "parent":
            parts = []
            for idx in range(1, current_level_index):
                value = _counter_value(counters, f"heading{idx}")
                if value <= 0:
                    continue
                parts.append(str(value))
            return ".".join(parts)
        ref_level_name = level_name
        if token.kind == "level" and token.level_index is not None:
            ref_level_name = f"heading{token.level_index}"
        value = _counter_value(counters, ref_level_name)
        return _format_counter_value(
            _level_format(ref_level_name, levels_config),
            value,
        )

    return _TEMPLATE_TOKEN_RE.sub(_replace, template_text)
