"""DOCX preprocessing utilities for known compatibility issues."""

from __future__ import annotations

import os
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


_MAX_ATTR_VALUE_LEN = 200_000
_OVERSIZED_ATTR_RE = re.compile(
    rb"(\s[\w:.-]+)\s*=\s*(\"([^\"]*)\"|'([^']*)')",
    re.DOTALL,
)


def sanitize_docx(src_path: str, aggressive: bool = False) -> str:
    """Sanitize DOCX in-place and return the same path.

    Fixes:
    - invalid .rels references (Target="NULL")
    - oversized XML attribute values (aggressive mode)
    """
    src = Path(src_path)
    if not src.exists():
        raise FileNotFoundError(f"File not found: {src}")

    needs_basic_fix = _check_needs_fix(str(src))
    needs_aggressive_fix = aggressive or _check_needs_aggressive_fix(str(src))

    if not needs_basic_fix and not needs_aggressive_fix:
        return str(src)

    tmp_fd, tmp_path = tempfile.mkstemp(
        suffix=".docx",
        prefix=f".{src.stem}.sanitize.",
        dir=str(src.parent),
    )
    os.close(tmp_fd)

    try:
        if needs_aggressive_fix:
            _repair_docx(
                str(src),
                tmp_path,
                aggressive_xml_fix=True,
            )
        else:
            _repair_docx(str(src), tmp_path)
        os.replace(tmp_path, str(src))
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

    return str(src)


def docx_needs_sanitization(src_path: str, aggressive: bool = False) -> bool:
    """Return whether the DOCX would be changed by ``sanitize_docx``."""
    src = Path(src_path)
    if not src.exists():
        raise FileNotFoundError(f"File not found: {src}")
    return _check_needs_fix(str(src)) or aggressive or _check_needs_aggressive_fix(str(src))


def _check_needs_fix(path: str) -> bool:
    """Quick check for basic known issues."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            for name in zf.namelist():
                if not name.endswith(".rels"):
                    continue
                if b"NULL" in zf.read(name):
                    return True
    except zipfile.BadZipFile:
        return False
    return False


def _check_needs_aggressive_fix(path: str) -> bool:
    """Detect parser-blocking XML payloads (e.g. extremely long attr values)."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            for name in _iter_probe_xml_parts(zf):
                try:
                    etree.fromstring(zf.read(name))
                except etree.XMLSyntaxError as exc:
                    msg = str(exc).lower()
                    if "attvalue length too long" in msg:
                        return True
    except zipfile.BadZipFile:
        return False
    return False


def _iter_probe_xml_parts(zf: zipfile.ZipFile):
    names = set(zf.namelist())
    priority = [
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
    ]
    for name in priority:
        if name in names:
            yield name

    for name in sorted(names):
        if not name.endswith(".xml"):
            continue
        if name.startswith("word/header") or name.startswith("word/footer"):
            yield name


def _repair_docx(src_path: str, dst_path: str, *, aggressive_xml_fix: bool = False):
    """Repair DOCX and write to dst_path."""
    with zipfile.ZipFile(src_path, "r") as zin:
        with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename.endswith(".rels"):
                    data = _fix_rels(data)

                if aggressive_xml_fix and item.filename.endswith(".xml"):
                    data = _fix_oversized_xml_attrs(data)

                zout.writestr(item, data)


def _fix_rels(data: bytes) -> bytes:
    """Remove invalid Relationship entries with Target=.../NULL."""
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        return data

    ns = root.nsmap.get(None, "")
    removed = False

    for rel in root.findall(f"{{{ns}}}Relationship"):
        target = rel.get("Target", "")
        basename = target.rsplit("/", 1)[-1] if "/" in target else target
        if basename.upper() == "NULL":
            root.remove(rel)
            removed = True

    if removed:
        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
    return data


def _fix_oversized_xml_attrs(data: bytes, *, max_len: int = _MAX_ATTR_VALUE_LEN) -> bytes:
    """Clear oversized XML attribute values that break libxml2 parsing limits."""
    if b'="' not in data and b"='" not in data:
        return data

    changed = False

    def _repl(match: re.Match[bytes]) -> bytes:
        nonlocal changed
        value = match.group(3) if match.group(3) is not None else match.group(4)
        if value is None or len(value) <= max_len:
            return match.group(0)
        changed = True
        return match.group(1) + b'=""'

    repaired = _OVERSIZED_ATTR_RE.sub(_repl, data)
    return repaired if changed else data
