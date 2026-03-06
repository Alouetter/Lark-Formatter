"""Refresh Word fields (especially TOC page numbers) via Word COM on Windows."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path


_PS_REFRESH_SCRIPT = r"""
$ErrorActionPreference = 'Stop'

$docPath = $env:DOCX_REFRESH_PATH
if (-not $docPath) {
    throw "DOCX_REFRESH_PATH is empty."
}
if (-not (Test-Path -LiteralPath $docPath)) {
    throw "File not found: $docPath"
}

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    # Open writable document and repaginate so TOC can compute real page numbers.
    $doc = $word.Documents.Open($docPath, $false, $false)
    $doc.Repaginate() | Out-Null

    foreach ($toc in @($doc.TablesOfContents)) {
        $toc.Update() | Out-Null
        $toc.UpdatePageNumbers() | Out-Null
    }

    # Pass-1: refresh TOC + reference numbering sequence fields first.
    foreach ($para in @($doc.Paragraphs)) {
        foreach ($fld in @($para.Range.Fields)) {
            try {
                $code = ($fld.Code.Text + '').ToUpperInvariant()
                if (
                    $code.Contains('PAGEREF') -or
                    $code.Contains('TOC') -or
                    ($code.Contains('SEQ') -and $code.Contains('REFENTRY'))
                ) {
                    $fld.Update() | Out-Null
                }
            } catch {
                # Some field codes may not be readable in malformed docs.
            }
        }
    }

    # Pass-2: refresh citation REF fields that target our reference bookmarks.
    foreach ($para in @($doc.Paragraphs)) {
        foreach ($fld in @($para.Range.Fields)) {
            try {
                $code = ($fld.Code.Text + '').ToUpperInvariant()
                if (
                    $code.Contains('REF ') -and
                    ($code.Contains('_REFNUM_') -or $code.Contains('_REFENTRY_'))
                ) {
                    $fld.Update() | Out-Null
                }
            } catch {
                # Some field codes may not be readable in malformed docs.
            }
        }
    }

    # One more repagination + TOC update for stability.
    $doc.Repaginate() | Out-Null
    foreach ($toc in @($doc.TablesOfContents)) {
        $toc.Update() | Out-Null
        $toc.UpdatePageNumbers() | Out-Null
    }

    $doc.Save()
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } catch {}
    }
    if ($word -ne $null) {
        try { $word.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""


_PYWIN32_REFRESH_CHILD = r"""
import os
import sys
from pathlib import Path

doc_path = (os.environ.get("DOCX_REFRESH_PATH") or "").strip()
if not doc_path:
    raise RuntimeError("DOCX_REFRESH_PATH is empty.")

import pythoncom
import win32com.client

word = None
doc = None
pythoncom.CoInitialize()
try:
    p = str(Path(doc_path).resolve())
    word = win32com.client.DispatchEx("Word.Application")
    try:
        word.Visible = False
    except Exception:
        pass
    try:
        word.DisplayAlerts = 0
    except Exception:
        pass

    doc = word.Documents.Open(
        p,
        False, False, False, "", "", False, "", "", 0, 0, False, True
    )
    doc.Repaginate()

    for toc in list(doc.TablesOfContents):
        try:
            toc.Update()
            toc.UpdatePageNumbers()
        except Exception:
            pass

    # Pass-1: refresh TOC + reference numbering sequence fields first.
    for para in list(doc.Paragraphs):
        try:
            fields = list(para.Range.Fields)
        except Exception:
            continue
        for fld in fields:
            try:
                code = (fld.Code.Text or "").upper()
            except Exception:
                code = ""
            if (
                ("PAGEREF" in code)
                or ("TOC" in code)
                or (("SEQ" in code) and ("REFENTRY" in code))
            ):
                try:
                    fld.Update()
                except Exception:
                    pass

    # Pass-2: refresh citation REF fields that target our reference bookmarks.
    for para in list(doc.Paragraphs):
        try:
            fields = list(para.Range.Fields)
        except Exception:
            continue
        for fld in fields:
            try:
                code = (fld.Code.Text or "").upper()
            except Exception:
                code = ""
            if ("REF " in code) and (("_REFNUM_" in code) or ("_REFENTRY_" in code)):
                try:
                    fld.Update()
                except Exception:
                    pass

    doc.Repaginate()
    for toc in list(doc.TablesOfContents):
        try:
            toc.Update()
            toc.UpdatePageNumbers()
        except Exception:
            pass

    doc.Save()
    print("ok(pywin32)")
finally:
    if doc is not None:
        try:
            doc.Close(False)
        except Exception:
            pass
    if word is not None:
        try:
            word.Quit()
        except Exception:
            pass
    pythoncom.CoUninitialize()
"""


def _refresh_via_pywin32(doc_path: Path, timeout_sec: int) -> tuple[bool, str]:
    """Run pywin32 refresh in a child process to avoid hanging UI thread."""
    if getattr(sys, "frozen", False):
        return False, "pywin32 refresh skipped in frozen executable."
    exe_name = Path(sys.executable).name.lower()
    if "python" not in exe_name:
        return False, "pywin32 skipped: current runtime is not a Python interpreter."

    env = os.environ.copy()
    env["DOCX_REFRESH_PATH"] = str(doc_path.resolve())
    cmd = [sys.executable, "-c", _PYWIN32_REFRESH_CHILD]
    try:
        proc = subprocess.run(
            cmd,
            env=env,
            capture_output=True,
            text=True,
            timeout=timeout_sec,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return False, f"pywin32 refresh timed out after {timeout_sec}s."
    except Exception as exc:
        return False, f"pywin32 refresh failed to start: {exc}"

    if proc.returncode == 0:
        detail = (proc.stdout or "").strip() or "ok(pywin32)"
        return True, detail

    stderr = (proc.stderr or "").strip()
    stdout = (proc.stdout or "").strip()
    detail = stderr if stderr else stdout
    if not detail:
        detail = f"pywin32 child exited with code {proc.returncode}"
    return False, detail


def _refresh_via_powershell(doc_path: Path, timeout_sec: int) -> tuple[bool, str]:
    """Fallback refresher via PowerShell COM script."""
    env = os.environ.copy()
    env["DOCX_REFRESH_PATH"] = str(doc_path.resolve())

    cmd = [
        "powershell",
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        _PS_REFRESH_SCRIPT,
    ]

    try:
        proc = subprocess.run(
            cmd,
            env=env,
            capture_output=True,
            text=True,
            timeout=timeout_sec,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return False, f"Word field refresh timed out after {timeout_sec}s."
    except Exception as exc:  # pragma: no cover - defensive branch
        return False, f"Word field refresh failed to start: {exc}"

    if proc.returncode == 0:
        return True, "ok(powershell)"

    stderr = (proc.stderr or "").strip()
    stdout = (proc.stdout or "").strip()
    detail = stderr if stderr else stdout
    if not detail:
        detail = f"powershell exited with code {proc.returncode}"
    return False, detail


def refresh_doc_fields_with_word(doc_path: str, timeout_sec: int = 10) -> tuple[bool, str]:
    """Try to refresh fields/page numbers in-place using local Microsoft Word.

    Returns:
        (True, "ok") on success, otherwise (False, reason).
    """
    if os.name != "nt":
        return False, "Word COM refresh requires Windows."

    p = Path(doc_path)
    if not p.exists():
        return False, f"Output file not found: {doc_path}"

    # 在临时副本上刷新，避免 COM 异常时锁住最终输出文件。
    tmp_dir = Path(tempfile.gettempdir()) / "lark_formatter_refresh"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    shadow = tmp_dir / f"{uuid.uuid4().hex}.docx"
    try:
        shutil.copy2(p, shadow)
    except Exception as exc:
        return False, f"Failed to prepare refresh shadow copy: {exc}"

    try:
        ok, detail = _refresh_via_pywin32(shadow, timeout_sec)
        if ok:
            try:
                shutil.copy2(shadow, p)
                return True, detail
            except Exception as exc:
                return False, f"Refreshed on shadow but failed to write back: {exc}"

        ok_ps, detail_ps = _refresh_via_powershell(shadow, timeout_sec)
        if ok_ps:
            try:
                shutil.copy2(shadow, p)
                return True, detail_ps
            except Exception as exc:
                return False, f"Refreshed on shadow but failed to write back: {exc}"

        return False, f"{detail}; fallback failed: {detail_ps}"
    finally:
        try:
            shadow.unlink(missing_ok=True)
        except Exception:
            pass
