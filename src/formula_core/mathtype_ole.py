"""MathType OLE/MTEF extraction helpers."""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
import math
import re
import unicodedata
from typing import Any

from lxml import etree

from .symbols import UNICODE_MATH_TO_LATEX, ensure_latex_command_word_boundaries

try:  # pragma: no cover
    import olefile
except Exception:  # pragma: no cover
    olefile = None

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_O_NS = "urn:schemas-microsoft-com:office:office"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_REC_LINE = 1
_REC_CHAR = 2
_REC_TMPL = 3
_REC_PILE = 4
_REC_MATRIX = 5
_REC_EMBELL = 6
_REC_RULER = 7
_REC_FONT_STYLE = 8
_REC_SIZE = 9
_REC_FULL = 10
_REC_SUB = 11
_REC_SUB2 = 12
_REC_SYM = 13
_REC_SUBSYM = 14
_REC_COLOR = 15
_REC_COLOR_DEF = 16
_REC_FONT_DEF = 17
_REC_EQN_PREFS = 18
_REC_ENCODING_DEF = 19

_OPT_NUDGE = 0x08
_OPT_CHAR_EMBELL = 0x01
_OPT_CHAR_FUNC_START = 0x02
_OPT_CHAR_POS8 = 0x04
_OPT_CHAR_POS16 = 0x10
_OPT_CHAR_NO_MTCODE = 0x20
_OPT_LINE_NULL = 0x01
_OPT_LINE_RULER = 0x02
_OPT_LINE_LSPACE = 0x04

_TM_ROOT = 10
_TM_FRACT = 11
_TM_UBAR = 12
_TM_OBAR = 13
_TM_ARROW = 14
_TM_INTEG = 15
_TM_SUM = 16
_TM_PROD = 17
_TM_COPROD = 18
_TM_UNION = 19
_TM_INTER = 20
_TM_INTOP = 21
_TM_SUMOP = 22
_TM_LIM = 23
_TM_HBRACE = 24
_TM_HBRACK = 25
_TM_LDIV = 26
_TM_SUB = 27
_TM_SUP = 28
_TM_SUBSUP = 29
_TM_DIRAC = 30
_TM_VEC = 31
_TM_TILDE = 32
_TM_HAT = 33
_TM_ARC = 34
_TM_BOX = 37

_FUNC_WORD_RE = re.compile(r"^[A-Za-z]+$")
_CTRL_WORD_RE = re.compile(r"^\\[A-Za-z]+$")
_BIG_OP = {
    _TM_SUM: r"\sum",
    _TM_PROD: r"\prod",
    _TM_COPROD: r"\coprod",
    _TM_UNION: r"\bigcup",
    _TM_INTER: r"\bigcap",
}
_UNICODE_LATEX = dict(UNICODE_MATH_TO_LATEX)
_UNICODE_LATEX_COMMANDS_RE = re.compile(
    "("
    + "|".join(
        sorted(
            {
                re.escape(value)
                for value in _UNICODE_LATEX.values()
                if value.startswith("\\") and value[1:].isalpha()
            },
            key=len,
            reverse=True,
        )
    )
    + r")(?=[A-Za-z0-9])"
)


@dataclass
class ParsedToken:
    kind: str
    text: str = ""
    sub: str = ""
    sup: str = ""
    precedes: bool = False
    needs_space_after: bool = False
    typesize: int | None = None


@dataclass
class MathTypeOleResult:
    source_type: str
    payload: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    confidence: float = 0.0
    object_only: bool = False


class _Reader:
    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0

    def eof(self) -> bool:
        return self.pos >= len(self.data)

    def read_byte(self) -> int:
        if self.eof():
            raise EOFError("unexpected end of MTEF stream")
        out = self.data[self.pos]
        self.pos += 1
        return out

    def read(self, size: int) -> bytes:
        if self.pos + size > len(self.data):
            raise EOFError("unexpected end of MTEF stream")
        out = self.data[self.pos:self.pos + size]
        self.pos += size
        return out

    def read_uint(self) -> int:
        first = self.read_byte()
        return first if first != 0xFF else int.from_bytes(self.read(2), "little")

    def read_sint(self) -> int:
        first = self.read_byte()
        return first - 128 if first != 0xFF else int.from_bytes(self.read(2), "little") - 32768

    def read_u16(self) -> int:
        return int.from_bytes(self.read(2), "little")

    def read_cstring(self) -> str:
        start = self.pos
        while self.pos < len(self.data) and self.data[self.pos] != 0:
            self.pos += 1
        raw = self.data[start:self.pos]
        if self.pos < len(self.data):
            self.pos += 1
        return raw.decode("latin-1", errors="ignore")


class _NibbleReader:
    def __init__(self, reader: _Reader):
        self.reader = reader
        self.pending_low: int | None = None

    def read(self) -> int:
        if self.pending_low is not None:
            out = self.pending_low
            self.pending_low = None
            return out
        byte = self.reader.read_byte()
        self.pending_low = byte & 0x0F
        return (byte >> 4) & 0x0F

    def align(self) -> None:
        self.pending_low = None


class _Parser:
    def __init__(self, data: bytes):
        self.r = _Reader(data)
        self.version = 0
        self.warnings: list[str] = []
        self.current_typesize: int | None = _REC_FULL

    def parse(self) -> str:
        self.version = self.r.read_byte()
        self.r.read(4)
        self.r.read_cstring()  # app key
        self.r.read_byte()  # equation options
        parts: list[str] = []
        while not self.r.eof():
            rec = self.r.read_byte()
            if rec == 0:
                continue
            tok = self._record(rec)
            if tok and tok.text:
                parts.append(tok.text)
        return self._normalize("".join(parts))

    def _record(self, rec: int) -> ParsedToken | None:
        if rec >= 100:
            self.r.read(self.r.read_uint())
            return None
        if rec == _REC_LINE:
            return self._line()
        if rec == _REC_CHAR:
            return self._char()
        if rec == _REC_TMPL:
            return self._tmpl()
        if rec == _REC_PILE:
            return self._pile()
        if rec == _REC_MATRIX:
            return self._matrix()
        if rec == _REC_RULER:
            self._skip_ruler()
            return None
        if rec in {_REC_FONT_STYLE, _REC_COLOR_DEF, _REC_FONT_DEF, _REC_EQN_PREFS, _REC_ENCODING_DEF, _REC_COLOR}:
            self._skip_def(rec)
            return None
        if rec in {_REC_SIZE, _REC_FULL, _REC_SUB, _REC_SUB2, _REC_SYM, _REC_SUBSYM}:
            self._apply_size_record(rec)
            return None
        if rec == _REC_EMBELL:
            self.r.read_byte()
            self.r.read_byte()
            return None
        raise ValueError(f"unsupported MTEF record type: {rec}")

    def _line(self) -> ParsedToken:
        options = self.r.read_byte()
        if options & _OPT_NUDGE:
            self._skip_nudge()
        if options & _OPT_LINE_LSPACE:
            self.r.read_u16()
        if options & _OPT_LINE_RULER:
            if self.r.read_byte() != _REC_RULER:
                raise ValueError("expected ruler record")
            self._skip_ruler()
        if options & _OPT_LINE_NULL:
            if not self.r.eof() and self.r.data[self.r.pos] == 0:
                self.r.read_byte()
            return ParsedToken(kind="text", text="")
        items: list[ParsedToken] = []
        while not self.r.eof():
            rec = self.r.read_byte()
            if rec == 0:
                break
            tok = self._record(rec)
            if tok is not None:
                items.append(tok)
        return ParsedToken(kind="text", text=self._compose_line(items))

    def _char(self) -> ParsedToken:
        options = self.r.read_byte()
        if options & _OPT_NUDGE:
            self._skip_nudge()
        typeface = self.r.read_sint()
        mtcode = None if options & _OPT_CHAR_NO_MTCODE else self.r.read_u16()
        pos8 = self.r.read_byte() if options & _OPT_CHAR_POS8 else None
        pos16 = self.r.read_u16() if options & _OPT_CHAR_POS16 else None
        if mtcode is not None:
            text = chr(mtcode)
        elif pos16 is not None:
            text = chr(pos16)
        elif pos8 is not None:
            text = chr(pos8)
        else:
            raise ValueError("unable to decode MathType char")
        if self._should_emit_upright_text(typeface, text):
            text = rf"\mathrm{{{text}}}"
        out = ParsedToken(
            kind="function" if (options & _OPT_CHAR_FUNC_START) or typeface == 2 else "text",
            text=text,
            typesize=self.current_typesize,
        )
        if options & _OPT_CHAR_EMBELL:
            while not self.r.eof():
                rec = self.r.read_byte()
                if rec == 0:
                    break
                if rec != _REC_EMBELL:
                    raise ValueError("expected embell record")
                emb = self._embell()
                out = ParsedToken(
                    kind="text",
                    text=self._apply_embell(out.text, emb),
                    typesize=out.typesize,
                )
        return out

    def _tmpl(self) -> ParsedToken:
        options = self.r.read_byte()
        if options & _OPT_NUDGE:
            self._skip_nudge()
        selector = self.r.read_byte()
        variation = self._variation()
        self.r.read_byte()  # template options
        if selector == _TM_ROOT:
            root = self._slot()
            body = self._slot()
            self._end()
            return ParsedToken(kind="text", text=rf"\sqrt[{root}]{{{body}}}" if variation == 1 and root else rf"\sqrt{{{body or root}}}")
        if selector == _TM_FRACT:
            num = self._slot()
            den = self._slot()
            self._end()
            return ParsedToken(kind="text", text=rf"\frac{{{num}}}{{{den}}}")
        if selector in {_TM_UBAR, _TM_OBAR}:
            body = self._slot()
            self._end()
            cmd = r"\underline" if selector == _TM_UBAR else r"\overline"
            return ParsedToken(kind="text", text=rf"{cmd}{{{body}}}")
        if selector in {_TM_SUB, _TM_SUP, _TM_SUBSUP}:
            sub = self._slot() if selector in {_TM_SUB, _TM_SUBSUP} else ""
            sup = self._slot() if selector in {_TM_SUP, _TM_SUBSUP} else ""
            return ParsedToken(kind="script", sub=sub, sup=sup, precedes=bool(variation & 0x0001))
        if selector in {_TM_INTEG, _TM_SUM, _TM_PROD, _TM_COPROD, _TM_UNION, _TM_INTER, _TM_INTOP, _TM_SUMOP}:
            body = self._slot()
            upper = self._slot() if variation & 0x0002 else ""
            lower = self._slot() if variation & 0x0001 else ""
            operator = self._slot()
            self._end()
            op = operator or (_BIG_OP.get(selector) if selector in _BIG_OP else (r"\int" if selector in {_TM_INTEG, _TM_INTOP} else r"\sum"))
            return ParsedToken(kind="text", text=op + (rf"_{{{lower}}}" if lower else "") + (rf"^{{{upper}}}" if upper else "") + ((" " + body) if body else ""))
        if selector == _TM_LIM:
            body = self._slot()
            lower = self._slot()
            upper = self._slot()
            self._end()
            return ParsedToken(kind="text", text=r"\lim" + (rf"_{{{lower}}}" if lower else "") + (rf"^{{{upper}}}" if upper else "") + ((" " + body) if body else ""))
        if selector in {_TM_HBRACE, _TM_HBRACK}:
            body = self._slot()
            note = self._slot()
            _ = self._slot()
            self._end()
            over = bool(variation & 0x0001)
            cmd = r"\overbrace" if over else r"\underbrace"
            return ParsedToken(kind="text", text=rf"{cmd}{{{body}}}" + ((rf"^{{{note}}}" if over else rf"_{{{note}}}") if note else ""))
        if selector == _TM_LDIV:
            dividend = self._slot()
            quotient = self._slot() if variation & 0x0001 else ""
            self._end()
            return ParsedToken(kind="text", text=rf"\frac{{{dividend}}}{{{quotient}}}" if quotient else dividend)
        if selector == _TM_DIRAC:
            left = self._slot()
            right = self._slot()
            if variation & 0x0001:
                self._slot()
            bar = self._slot()
            if variation & 0x0002:
                self._slot()
            self._end()
            mid = bar or "|"
            if left and right:
                return ParsedToken(kind="text", text=rf"\langle {left} {mid} {right} \rangle")
            return ParsedToken(kind="text", text=" ".join(v for v in [left, mid, right] if v))
        if selector in {_TM_ARROW, _TM_VEC, _TM_TILDE, _TM_HAT, _TM_BOX}:
            body = self._slot()
            extra = self._slot()
            self._end()
            if selector == _TM_ARROW:
                return ParsedToken(kind="text", text=extra or r"\rightarrow")
            if selector == _TM_VEC:
                return ParsedToken(kind="text", text=rf"\vec{{{body or extra}}}")
            if selector == _TM_TILDE:
                return ParsedToken(kind="text", text=rf"\tilde{{{body or extra}}}")
            if selector == _TM_HAT:
                return ParsedToken(kind="text", text=rf"\hat{{{body or extra}}}")
            return ParsedToken(kind="text", text=rf"\boxed{{{body or extra}}}")
        if selector == _TM_ARC:
            body = self._slot()
            self._end()
            return ParsedToken(kind="text", text=rf"\overset{{\frown}}{{{body}}}")
        slots: list[str] = []
        while not self.r.eof():
            rec = self.r.read_byte()
            if rec == 0:
                break
            tok = self._record(rec)
            if tok and tok.text:
                slots.append(tok.text)
        self.warnings.append(f"unsupported_mathtype_template:{selector}")
        return ParsedToken(kind="text", text=" ".join(slots))

    def _pile(self) -> ParsedToken:
        options = self.r.read_byte()
        if options & _OPT_NUDGE:
            self._skip_nudge()
        self.r.read(2)
        if options & _OPT_LINE_RULER:
            if self.r.read_byte() != _REC_RULER:
                raise ValueError("expected pile ruler")
            self._skip_ruler()
        lines: list[str] = []
        while not self.r.eof():
            rec = self.r.read_byte()
            if rec == 0:
                break
            tok = self._record(rec)
            if tok and tok.text:
                lines.append(tok.text)
        if len(lines) <= 1:
            return ParsedToken(kind="text", text=lines[0] if lines else "")
        return ParsedToken(kind="text", text=rf"\begin{{matrix}}{' \\\\ '.join(lines)}\end{{matrix}}")

    def _matrix(self) -> ParsedToken:
        options = self.r.read_byte()
        if options & _OPT_NUDGE:
            self._skip_nudge()
        self.r.read(3)
        rows = self.r.read_byte()
        cols = self.r.read_byte()
        self.r.read(math.ceil((rows + 1) / 4) + math.ceil((cols + 1) / 4))
        cells: list[str] = []
        while not self.r.eof():
            rec = self.r.read_byte()
            if rec == 0:
                break
            tok = self._record(rec)
            if tok and tok.text:
                cells.append(tok.text)
        if rows <= 0 or cols <= 0 or not cells:
            return ParsedToken(kind="text", text="")
        rows_out: list[str] = []
        idx = 0
        for _ in range(rows):
            row = cells[idx:idx + cols]
            idx += cols
            row.extend([""] * max(0, cols - len(row)))
            rows_out.append(" & ".join(row))
        return ParsedToken(kind="text", text=rf"\begin{{matrix}}{' \\\\ '.join(rows_out)}\end{{matrix}}")

    def _slot(self) -> str:
        while not self.r.eof():
            rec = self.r.read_byte()
            if rec == 0:
                return ""
            tok = self._record(rec)
            if tok is None:
                continue
            return str(tok.text or "")
        return ""

    def _end(self) -> None:
        if not self.r.eof() and self.r.read_byte() != 0:
            raise ValueError("missing template end")

    def _variation(self) -> int:
        first = self.r.read_byte()
        return (first & 0x7F) | (self.r.read_byte() << 8) if first & 0x80 else first

    def _embell(self) -> int:
        options = self.r.read_byte()
        if options & _OPT_NUDGE:
            self._skip_nudge()
        return self.r.read_byte()

    def _skip_nudge(self) -> None:
        dx = self.r.read_byte()
        dy = self.r.read_byte()
        if dx == 128 and dy == 128:
            self.r.read(4)

    def _skip_ruler(self) -> None:
        for _ in range(self.r.read_byte()):
            self.r.read_byte()
            self.r.read_u16()

    def _apply_size_record(self, rec: int) -> None:
        if rec in {_REC_FULL, _REC_SUB, _REC_SUB2, _REC_SYM, _REC_SUBSYM}:
            self.current_typesize = rec
            return
        marker = self.r.read_byte()
        if marker == 101:
            self.r.read(2)
            self.current_typesize = None
            return
        if marker == 100:
            logical_size = self.r.read_byte()
            delta_size = self.r.read_byte()
            self.current_typesize = logical_size if delta_size == 128 else None
            return
        delta_size = self.r.read_byte()
        self.current_typesize = marker if delta_size == 128 else None

    def _skip_def(self, rec: int) -> None:
        if rec == _REC_COLOR:
            self.r.read_uint()
            return
        if rec == _REC_FONT_STYLE:
            self.r.read_uint()
            self.r.read_byte()
            return
        if rec == _REC_COLOR_DEF:
            options = self.r.read_byte()
            self.r.read(8 if options & 0x01 else 6)
            if options & 0x04:
                self.r.read_cstring()
            return
        if rec == _REC_FONT_DEF:
            self.r.read_uint()
            self.r.read_cstring()
            return
        if rec == _REC_ENCODING_DEF:
            self.r.read_cstring()
            return
        if rec == _REC_EQN_PREFS:
            self.r.read_byte()
            self._skip_dim_array()
            self._skip_dim_array()
            for _ in range(self.r.read_byte()):
                if self.r.read_uint() != 0:
                    self.r.read_byte()

    def _skip_dim_array(self) -> None:
        nibs = _NibbleReader(self.r)
        for _ in range(self.r.read_byte()):
            nibs.read()
            while nibs.read() != 0x0F:
                pass
        nibs.align()

    def _compose_line(self, items: list[ParsedToken]) -> str:
        merged: list[ParsedToken] = []
        i = 0
        while i < len(items):
            item = items[i]
            if item.kind != "function":
                merged.append(item)
                i += 1
                continue
            word = [item.text]
            i += 1
            while (
                i < len(items)
                and items[i].kind == "function"
                and items[i].typesize == item.typesize
            ):
                word.append(items[i].text)
                i += 1
            joined = "".join(word)
            merged.append(
                ParsedToken(
                    kind="text",
                    text="\\" + joined if _FUNC_WORD_RE.fullmatch(joined) else joined,
                    needs_space_after=bool(_FUNC_WORD_RE.fullmatch(joined)),
                    typesize=item.typesize,
                )
            )
        out: list[ParsedToken] = []
        idx = 0
        while idx < len(merged):
            item = merged[idx]
            if item.kind != "script":
                out.append(item)
                idx += 1
                continue
            if item.precedes:
                if idx + 1 < len(merged) and merged[idx + 1].kind != "script":
                    base = merged[idx + 1]
                    base_text = base.text
                    advance = idx + 2
                    while advance < len(merged) and merged[advance].kind == "script" and not merged[advance].precedes:
                        base_text = self._apply_script_to_text(base_text, merged[advance])
                        advance += 1
                    out.append(
                        ParsedToken(
                            kind="text",
                            text=self._prescript_text(item, base_text),
                            needs_space_after=base.needs_space_after,
                            typesize=base.typesize,
                        )
                    )
                    idx = advance
                    continue
                out.append(ParsedToken(kind="text", text=self._script_text(item), typesize=item.typesize))
                idx += 1
                continue
            if out:
                base = out.pop()
                text = self._apply_script_to_text(base.text, item)
                out.append(
                    ParsedToken(
                        kind="text",
                        text=text,
                        needs_space_after=base.needs_space_after,
                        typesize=base.typesize,
                    )
                )
            else:
                out.append(ParsedToken(kind="text", text=self._script_text(item), typesize=item.typesize))
            idx += 1
        out = self._fold_implicit_size_scripts(out)
        parts: list[str] = []
        for idx, item in enumerate(out):
            if item.text:
                parts.append(item.text)
                if item.needs_space_after and idx + 1 < len(out):
                    parts.append(" ")
        return "".join(parts)

    @staticmethod
    def _wrap(text: str) -> str:
        value = str(text or "").strip()
        if not value:
            return "{}"
        if len(value) == 1 or value.startswith("\\") or _CTRL_WORD_RE.fullmatch(value) or (value.startswith("{") and value.endswith("}")):
            return value
        return "{" + value + "}"

    @staticmethod
    def _script_text(item: ParsedToken) -> str:
        return (rf"_{{{item.sub}}}" if item.sub else "") + (rf"^{{{item.sup}}}" if item.sup else "")

    def _apply_script_to_text(self, base_text: str, item: ParsedToken) -> str:
        return self._wrap(base_text) + self._script_text(item)

    @staticmethod
    def _prescript_text(item: ParsedToken, base_text: str) -> str:
        return rf"\prescript{{{item.sup}}}{{{item.sub}}}{{{base_text}}}"

    @staticmethod
    def _should_emit_upright_text(typeface: int, text: str) -> bool:
        if typeface != 1:
            return False
        value = str(text or "")
        if not value:
            return False
        return all(unicodedata.category(ch).startswith("L") for ch in value)

    def _fold_implicit_size_scripts(self, items: list[ParsedToken]) -> list[ParsedToken]:
        out: list[ParsedToken] = []
        idx = 0
        while idx < len(items):
            item = items[idx]
            if not self._is_implicit_subscript_token(item):
                out.append(item)
                idx += 1
                continue
            if not out or not self._can_host_implicit_subscript(out[-1]):
                out.append(item)
                idx += 1
                continue
            cluster = [item.text]
            advance = idx + 1
            while advance < len(items) and self._is_implicit_subscript_token(items[advance]):
                cluster.append(items[advance].text)
                advance += 1
            base = out.pop()
            out.append(
                ParsedToken(
                    kind="text",
                    text=self._wrap(base.text) + rf"_{{{''.join(cluster)}}}",
                    needs_space_after=base.needs_space_after,
                    typesize=base.typesize,
                )
            )
            idx = advance
        return out

    @staticmethod
    def _is_implicit_subscript_token(item: ParsedToken) -> bool:
        if item.kind != "text":
            return False
        if item.typesize not in {_REC_SUB, _REC_SUB2}:
            return False
        return _Parser._looks_like_inline_script_text(item.text)

    @staticmethod
    def _can_host_implicit_subscript(item: ParsedToken) -> bool:
        if item.kind != "text":
            return False
        return _Parser._looks_like_script_host_text(item.text)

    @staticmethod
    def _looks_like_inline_script_text(text: str) -> bool:
        value = str(text or "").strip()
        if not value:
            return False
        if value.startswith("\\") and value.endswith("}"):
            return True
        return all(
            ch.isalnum() or unicodedata.category(ch).startswith(("L", "N"))
            for ch in value
        )

    @staticmethod
    def _looks_like_script_host_text(text: str) -> bool:
        value = str(text or "").strip()
        if not value:
            return False
        if value.startswith("\\"):
            return True
        return any(
            ch.isalnum() or unicodedata.category(ch).startswith(("L", "N"))
            for ch in value
        )

    def _apply_embell(self, text: str, emb: int) -> str:
        return {
            5: f"{text}'",
            6: f"{text}''",
            8: rf"\tilde{{{text}}}",
            9: rf"\hat{{{text}}}",
            11: rf"\overrightarrow{{{text}}}",
            12: rf"\overleftarrow{{{text}}}",
            16: rf"\bar{{{text}}}",
            17: rf"\overline{{{text}}}",
            29: rf"\underline{{{text}}}",
        }.get(emb, text)

    @staticmethod
    def _normalize(text: str) -> str:
        value = re.sub(r"\s+", " ", str(text or "").strip())
        value = re.sub(r"\s+([,.;:!?])", r"\1", value)
        value = re.sub(r"([({\[])\s+", r"\1", value)
        value = re.sub(r"\s+([)}\]])", r"\1", value)
        value = _normalize_unicode_math_latex(value)
        return value.strip()


def _normalize_unicode_math_latex(text: str) -> str:
    value = str(text or "")
    for src, dst in _UNICODE_LATEX.items():
        value = value.replace(src, dst)
    value = ensure_latex_command_word_boundaries(value)
    value = re.sub(r"(\\times)(?=[A-Za-z0-9])", r"\1 ", value)
    value = _UNICODE_LATEX_COMMANDS_RE.sub(r"\1 ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def _classify_source(prog_id: str, raw_xml: str) -> str:
    value = str(prog_id or "").strip().lower()
    raw = str(raw_xml or "").lower()
    if "dsmt" in value or "mathtype" in value or "equation.dsmt" in raw:
        return "mathtype"
    if value.startswith("equation.3") or "microsoft equation" in raw:
        return "old_equation"
    return "ole_equation"


def _non_object_text(paragraph) -> str:
    return "".join(
        paragraph._p.xpath(
            (
                ".//*[namespace-uri()='%s' and local-name()='t' and "
                "not(ancestor::*[namespace-uri()='%s' and local-name()='object'])]/text()"
            ) % (_W_NS, _W_NS)
        )
    )


def _extract_equation_native(ole_blob: bytes) -> bytes | None:
    if olefile is None:
        raise RuntimeError("olefile_dependency_missing")
    with olefile.OleFileIO(BytesIO(ole_blob)) as ole:
        if not ole.exists("Equation Native"):
            return None
        return ole.openstream("Equation Native").read()


def decode_equation_native_to_latex(data: bytes) -> tuple[str, list[str]]:
    if len(data) >= 28 and int.from_bytes(data[:4], "little") == 28:
        data = data[28:]
    if not data:
        raise ValueError("empty_equation_native_stream")
    parser = _Parser(data)
    if data[0] != 5:
        raise ValueError(f"unsupported_mtef_version:{data[0]}")
    latex = parser.parse()
    if not latex:
        raise ValueError("empty_mtef_formula")
    return latex, list(dict.fromkeys(parser.warnings))


def extract_ole_formula_candidates(paragraph) -> list[MathTypeOleResult]:
    objects = paragraph._p.xpath(".//*[namespace-uri()='%s' and local-name()='object']" % _W_NS)
    if not objects:
        return []
    raw_xml = etree.tostring(paragraph._p, encoding="unicode")
    object_only = not _non_object_text(paragraph).strip()
    results: list[MathTypeOleResult] = []
    for obj in objects:
        ole_nodes = obj.xpath(".//*[namespace-uri()='%s' and local-name()='OLEObject']" % _O_NS)
        if not ole_nodes:
            continue
        ole = ole_nodes[0]
        prog_id = str(ole.get("ProgID") or "").strip()
        rel_id = ole.get(f"{{{_R_NS}}}id")
        source_type = _classify_source(prog_id, raw_xml)
        payload = {"text": "", "ole_prog_id": prog_id, "ole_relationship_id": rel_id or ""}
        if source_type != "mathtype":
            results.append(MathTypeOleResult(source_type=source_type, payload=payload, warnings=["ole_binary_unparsed"], confidence=0.58, object_only=object_only))
            continue
        if not rel_id or rel_id not in paragraph.part.rels:
            payload["diagnostic_code"] = "ole_relationship_missing"
            results.append(MathTypeOleResult(source_type=source_type, payload=payload, warnings=["ole_relationship_missing", "mathtype_binary_unparsed"], object_only=object_only))
            continue
        rel = paragraph.part.rels[rel_id]
        target_part = getattr(rel, "target_part", None)
        payload["ole_target_ref"] = str(getattr(rel, "target_ref", "") or "")
        payload["ole_partname"] = str(getattr(target_part, "partname", "") or "")
        blob = getattr(target_part, "blob", None)
        if not blob:
            payload["diagnostic_code"] = "ole_target_missing"
            results.append(MathTypeOleResult(source_type=source_type, payload=payload, warnings=["ole_target_missing", "mathtype_binary_unparsed"], object_only=object_only))
            continue
        try:
            native = _extract_equation_native(blob)
        except Exception as exc:
            payload["diagnostic_code"] = "mathtype_binary_unparsed"
            payload["diagnostic_detail"] = str(exc)
            results.append(MathTypeOleResult(source_type=source_type, payload=payload, warnings=["mathtype_binary_unparsed"], object_only=object_only))
            continue
        if not native:
            payload["diagnostic_code"] = "missing_equation_native_stream"
            results.append(MathTypeOleResult(source_type=source_type, payload=payload, warnings=["missing_equation_native_stream", "mathtype_binary_unparsed"], object_only=object_only))
            continue
        try:
            latex, extra = decode_equation_native_to_latex(native)
        except Exception as exc:
            detail = str(exc or "").strip()
            payload["diagnostic_code"] = detail if detail.startswith("unsupported_mtef_version") else "mathtype_decode_failed"
            if detail and payload["diagnostic_code"] == "mathtype_decode_failed":
                payload["diagnostic_detail"] = detail
            results.append(MathTypeOleResult(source_type=source_type, payload=payload, warnings=[payload["diagnostic_code"]], object_only=object_only))
            continue
        payload["latex"] = latex
        payload["normalized_latex"] = latex
        payload["mtef_version"] = 5
        results.append(MathTypeOleResult(source_type=source_type, payload=payload, warnings=["mathtype_mtef_decoded", *extra], confidence=0.86 if not extra else 0.78, object_only=object_only))
    return results
