"""Optional MathType fallback using Word COM + OLE IDataObject extraction."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
from copy import deepcopy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from src.formula_core.convert import convert_mathml_to_omml

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_O_NS = "urn:schemas-microsoft-com:office:office"

_PYWIN32_MATHTYPE_CHILD = r"""
import json
import os
import ctypes
from pathlib import Path

import pythoncom
import win32com.client

DOC_PATH = (os.environ.get("DOCX_MATHTYPE_PATH") or "").strip()
if not DOC_PATH:
    raise RuntimeError("DOCX_MATHTYPE_PATH is empty")

user32 = ctypes.windll.user32

def _fmt_name(fmt_id: int) -> str:
    if fmt_id < 0xC000:
        return str(fmt_id)
    buf = ctypes.create_unicode_buffer(256)
    n = user32.GetClipboardFormatNameW(int(fmt_id), buf, len(buf))
    if n <= 0:
        return str(fmt_id)
    return buf.value[:n]

def _decode_medium_data(data):
    if isinstance(data, bytes):
        raw = data.rstrip(b"\x00")
        for enc in ("utf-8", "utf-16-le", "utf-16-be", "latin-1"):
            try:
                return raw.decode(enc).strip()
            except Exception:
                pass
        return raw.decode("utf-8", errors="ignore").strip()
    return str(data or "").replace("\x00", "").strip()

pythoncom.OleInitialize()
word = None
doc = None
result = {
    "ok": False,
    "detail": "",
    "found": 0,
    "extracted": 0,
    "mathml_items": [],
    "failures": [],
}
try:
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    try:
        word.Visible = False
    except Exception:
        pass
    try:
        word.DisplayAlerts = 0
    except Exception:
        pass

    doc = word.Documents.Open(str(Path(DOC_PATH).resolve()))
    mathml_format_ids = []
    for name in ("MathML Presentation", "MathML", "application/mathml+xml"):
        fmt_id = user32.RegisterClipboardFormatW(name)
        if fmt_id:
            mathml_format_ids.append((name, fmt_id))

    for idx in range(1, doc.InlineShapes.Count + 1):
        shape = doc.InlineShapes(idx)
        try:
            prog_id = str(shape.OLEFormat.ProgID or "")
        except Exception:
            continue
        prog_norm = prog_id.lower()
        if "dsmt" not in prog_norm and "mathtype" not in prog_norm:
            continue

        result["found"] += 1
        try:
            obj = shape.OLEFormat.Object
        except Exception as exc:
            result["failures"].append({"index": idx, "prog_id": prog_id, "reason": f"ole_object_unavailable:{exc}"})
            continue

        try:
            data_obj = obj._oleobj_.QueryInterface(pythoncom.IID_IDataObject)
        except Exception as exc:
            result["failures"].append({"index": idx, "prog_id": prog_id, "reason": f"idataobject_unavailable:{exc}"})
            continue

        got_text = ""
        got_name = ""
        available = []
        try:
            for fe in data_obj.EnumFormatEtc():
                fmt, _td, _aspect, _index, _tymed = fe
                available.append(_fmt_name(fmt))
        except Exception:
            pass

        for fmt_name, fmt_id in mathml_format_ids:
            fe = (fmt_id, None, pythoncom.DVASPECT_CONTENT, -1, pythoncom.TYMED_HGLOBAL)
            try:
                data_obj.QueryGetData(fe)
                medium = data_obj.GetData(fe)
                got_text = _decode_medium_data(medium.data)
                got_name = fmt_name
                if got_text:
                    break
            except Exception:
                continue

        if got_text:
            result["mathml_items"].append({"index": idx, "prog_id": prog_id, "format": got_name, "mathml": got_text})
            result["extracted"] += 1
        else:
            result["failures"].append({"index": idx, "prog_id": prog_id, "reason": "mathml_format_unavailable", "formats": available[:20]})

    result["ok"] = result["extracted"] > 0
    if result["ok"]:
        result["detail"] = f"extracted={result['extracted']}/{result['found']}"
    elif result["found"] == 0:
        result["detail"] = "no_mathtype_ole_found"
    else:
        result["detail"] = "no_mathml_extracted"
finally:
    if doc is not None:
        try:
            doc.Close(False)
        except Exception:
            pass
    if word is not None:
        try:
            word.Quit(False)
        except Exception:
            pass
    pythoncom.CoUninitialize()

print(json.dumps(result, ensure_ascii=False))
"""


def _w(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"


def _is_mathtype_object_node(obj_el) -> bool:
    try:
        raw = etree.tostring(obj_el, encoding="unicode").lower()
    except Exception:
        raw = ""
    return "equation.dsmt" in raw or "mathtype" in raw or "dsmt" in raw


def _paragraph_non_object_text(paragraph_el) -> str:
    return "".join(
        paragraph_el.xpath(
            (
                ".//*[namespace-uri()='%s' and local-name()='t' and "
                "not(ancestor::*[namespace-uri()='%s' and local-name()='object'])]/text()"
            ) % (_W_NS, _W_NS)
        )
    )


def count_mathtype_ole_objects(doc_path: str) -> int:
    path = Path(doc_path)
    if not path.exists():
        return 0
    try:
        with ZipFile(path) as zf:
            root = etree.fromstring(zf.read("word/document.xml"))
    except Exception:
        return 0
    count = 0
    for obj_el in root.xpath(".//*[namespace-uri()='%s' and local-name()='object']" % _W_NS):
        if _is_mathtype_object_node(obj_el):
            count += 1
    return count


def _replace_paragraph_with_omml(paragraph_el, omml_element) -> bool:
    ppr = paragraph_el.find(_w("pPr"))
    for child in list(paragraph_el):
        if child is not ppr:
            paragraph_el.remove(child)

    omml_copy = deepcopy(omml_element)
    local = omml_copy.tag.split("}")[-1] if "}" in omml_copy.tag else omml_copy.tag
    if local == "oMathPara":
        paragraph_el.append(omml_copy)
        return True

    run = etree.Element(_w("r"))
    run.append(omml_copy)
    paragraph_el.append(run)
    return True


def apply_mathml_fallback_payloads_to_docx(doc_path: str, mathml_payloads: list[dict]) -> dict[str, int]:
    path = Path(doc_path)
    if not path.exists():
        return {"found": 0, "replaced": 0, "skipped_mixed": 0, "unused_payloads": 0}

    with ZipFile(path, "r") as zin:
        entries = {info.filename: zin.read(info.filename) for info in zin.infolist()}

    if "word/document.xml" not in entries:
        return {"found": 0, "replaced": 0, "skipped_mixed": 0, "unused_payloads": len(mathml_payloads)}

    root = etree.fromstring(entries["word/document.xml"])
    paragraphs = root.xpath(".//*[namespace-uri()='%s' and local-name()='p']" % _W_NS)
    found = 0
    replaced = 0
    skipped_mixed = 0
    payload_index = 0

    for paragraph in paragraphs:
        obj_nodes = paragraph.xpath(".//*[namespace-uri()='%s' and local-name()='object']" % _W_NS)
        if not obj_nodes:
            continue
        if not any(_is_mathtype_object_node(obj_el) for obj_el in obj_nodes):
            continue
        found += 1
        if payload_index >= len(mathml_payloads):
            break
        payload = mathml_payloads[payload_index]
        payload_index += 1
        mathml = str(payload.get("mathml") or "").strip()
        if not mathml:
            continue
        block = not _paragraph_non_object_text(paragraph).strip()
        if not block:
            skipped_mixed += 1
            continue
        omml = convert_mathml_to_omml(mathml, block=True)
        if omml is None:
            continue
        if _replace_paragraph_with_omml(paragraph, omml):
            replaced += 1

    entries["word/document.xml"] = etree.tostring(root, encoding="utf-8", xml_declaration=True, standalone="yes")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx", dir=str(path.parent)) as tmp:
        tmp_path = Path(tmp.name)
    try:
        with ZipFile(tmp_path, "w", ZIP_DEFLATED) as zout:
            for name, data in entries.items():
                zout.writestr(name, data)
        shutil.move(str(tmp_path), str(path))
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass

    return {
        "found": found,
        "replaced": replaced,
        "skipped_mixed": skipped_mixed,
        "unused_payloads": max(0, len(mathml_payloads) - payload_index),
    }


def extract_mathtype_mathml_with_word(doc_path: str, timeout_sec: int = 30) -> tuple[bool, str, dict]:
    if os.name != "nt":
        return False, "Word COM fallback requires Windows.", {}
    if getattr(sys, "frozen", False):
        return False, "office fallback skipped in frozen executable.", {}
    exe_name = Path(sys.executable).name.lower()
    if "python" not in exe_name:
        return False, "office fallback requires Python interpreter runtime.", {}

    env = os.environ.copy()
    env["DOCX_MATHTYPE_PATH"] = str(Path(doc_path).resolve())
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    cmd = [sys.executable, "-c", _PYWIN32_MATHTYPE_CHILD]
    try:
        proc = subprocess.run(
            cmd,
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=max(10, int(timeout_sec)),
            check=False,
        )
    except subprocess.TimeoutExpired:
        return False, f"office fallback timed out after {timeout_sec}s", {}
    except Exception as exc:
        return False, f"office fallback failed to start: {exc}", {}

    payload_text = (proc.stdout or "").strip()
    if proc.returncode != 0:
        detail = (proc.stderr or "").strip() or payload_text or f"child exited with code {proc.returncode}"
        return False, detail, {}

    try:
        payload = json.loads(payload_text or "{}")
    except Exception as exc:
        return False, f"invalid office fallback payload: {exc}", {}

    ok = bool(payload.get("ok"))
    detail = str(payload.get("detail") or ("ok" if ok else "failed"))
    return ok, detail, payload


def apply_mathtype_office_fallback(doc_path: str, timeout_sec: int = 30) -> tuple[bool, str, dict]:
    found = count_mathtype_ole_objects(doc_path)
    if found <= 0:
        return False, "no_mathtype_ole_found", {"found": 0, "extracted": 0, "replaced": 0}

    ok, detail, payload = extract_mathtype_mathml_with_word(doc_path, timeout_sec=timeout_sec)
    if not ok:
        stats = {"found": found, "extracted": 0, "replaced": 0}
        if isinstance(payload, dict):
            stats.update({k: payload.get(k) for k in ("found", "extracted") if k in payload})
        return False, detail, stats

    mathml_payloads = list(payload.get("mathml_items") or [])
    replace_stats = apply_mathml_fallback_payloads_to_docx(doc_path, mathml_payloads)
    stats = {
        "found": found,
        "extracted": int(payload.get("extracted", 0) or 0),
        **replace_stats,
    }
    changed = int(replace_stats.get("replaced", 0) or 0) > 0
    summary = (
        f"extracted={stats['extracted']}/{stats['found']}, "
        f"replaced={stats['replaced']}, skipped_mixed={stats['skipped_mixed']}"
    )
    return changed, summary, stats
