"""Tests for compound formula conversion and bare-script normalization."""

from lxml import etree

from src.formula_core.ast import FormulaNode
from src.formula_core.convert import (
    ConversionOutcome,
    _build_compound_omml,
    _shape_from_latex,
    _tokenize_latex,
    convert_formula_node,
)
from src.formula_core.normalize import (
    _normalize_bare_scripts,
    text_formula_to_latex,
)

_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _math_texts(omml) -> list[str]:
    return [
        str(text or "")
        for text in omml.xpath(
            ".//*[namespace-uri()='%s' and local-name()='t']/text()" % _M_NS
        )
    ]


def _first_math_run_style(omml) -> tuple[list[str], list[str]]:
    first_run = omml.xpath(
        ".//*[namespace-uri()='%s' and local-name()='r'][1]" % _M_NS
    )[0]
    sty_vals = first_run.xpath(
        "./*[namespace-uri()='%s' and local-name()='rPr']"
        "/*[namespace-uri()='%s' and local-name()='sty']"
        "/@*[local-name()='val']"
        % (_M_NS, _M_NS)
    )
    scr_vals = first_run.xpath(
        "./*[namespace-uri()='%s' and local-name()='rPr']"
        "/*[namespace-uri()='%s' and local-name()='scr']"
        "/@*[local-name()='val']"
        % (_M_NS, _M_NS)
    )
    return list(sty_vals), list(scr_vals)


def _math_run_payloads(omml) -> list[tuple[str, list[str], list[str]]]:
    runs = omml.xpath(
        ".//*[namespace-uri()='%s' and local-name()='r']" % _M_NS
    )
    payloads: list[tuple[str, list[str], list[str]]] = []
    for run in runs:
        text = "".join(
            run.xpath(
                "./*[namespace-uri()='%s' and local-name()='t']/text()" % _M_NS
            )
        )
        sty_vals = run.xpath(
            "./*[namespace-uri()='%s' and local-name()='rPr']"
            "/*[namespace-uri()='%s' and local-name()='sty']"
            "/@*[local-name()='val']"
            % (_M_NS, _M_NS)
        )
        scr_vals = run.xpath(
            "./*[namespace-uri()='%s' and local-name()='rPr']"
            "/*[namespace-uri()='%s' and local-name()='scr']"
            "/@*[local-name()='val']"
            % (_M_NS, _M_NS)
        )
        payloads.append((text, list(sty_vals), list(scr_vals)))
    return payloads


# ── _normalize_bare_scripts ──


def test_bare_caret_normalized():
    assert _normalize_bare_scripts("a^2") == "a^{2}"


def test_bare_underscore_normalized():
    assert _normalize_bare_scripts("x_i") == "x_{i}"


def test_compound_bare_caret():
    assert _normalize_bare_scripts("a^2+b^2=c^2") == "a^{2}+b^{2}=c^{2}"


def test_already_braced_unchanged():
    assert _normalize_bare_scripts("a^{2}+b^{2}") == "a^{2}+b^{2}"


def test_mixed_bare_and_braced():
    assert _normalize_bare_scripts("a^2+b^{2}") == "a^{2}+b^{2}"


def test_bare_sub_compound():
    assert _normalize_bare_scripts("a_1+a_2") == "a_{1}+a_{2}"


def test_bare_sub_sup_on_same_base():
    result = _normalize_bare_scripts("x_i^2")
    assert result == "x_{i}^{2}"


def test_bare_caret_multichar_exponent_untouched():
    # ^10 has multi-digit exponent — regex does not split it, left unchanged.
    # This is safe: the compound tokenizer will handle it as text fallback.
    result = _normalize_bare_scripts("x^10")
    assert result == "x^10"


def test_emc2():
    assert _normalize_bare_scripts("E=mc^2") == "E=mc^{2}"


# ── text_formula_to_latex integration ──


def test_text_formula_normalizes_bare_scripts():
    latex, conf, _ = text_formula_to_latex("a^2+b^2=c^2")
    assert "^{2}" in latex
    assert "^2" not in latex.replace("^{2}", "")


def test_text_formula_emc2():
    latex, conf, _ = text_formula_to_latex("E=mc^2")
    assert "^{2}" in latex


def test_text_formula_recovers_escaped_fraction_chain():
    latex, _, warnings = text_formula_to_latex(
        "ab+cd=ad+bcbdESC_a9cae316rac{a}{b}+ESC_3a30d405rac{c}{d}=ESC_8f7ff319rac{ad+bc}{bd}ba\u200b+dc\u200b=bdad+bc\u200b"
    )
    assert latex == r"\frac{a}{b}+\frac{c}{d}=\frac{ad+bc}{bd}"
    assert "escape_placeholder_command_recovered" in warnings
    assert "escape_placeholder_formula_extracted" in warnings


def test_text_formula_recovers_escaped_single_fraction():
    latex, _, warnings = text_formula_to_latex(
        "11−xESC_57aa258drac{1}{1-x}1−x1\u200b"
    )
    assert latex == r"\frac{1}{1-x}"
    assert "escape_placeholder_formula_extracted" in warnings


def test_text_formula_recovers_escaped_sqrt():
    latex, _, warnings = text_formula_to_latex(
        "a2+b2ESC_04e74865qrt{a^2+b^2}a2+b2\u200b"
    )
    assert latex == r"\sqrt{a^{2}+b^{2}}"
    assert "escape_placeholder_formula_extracted" in warnings


def test_text_formula_extracts_rendered_short_command_segment():
    latex, _, warnings = text_formula_to_latex(
        "\u3000arcsin\u2061xESC_bd6e0b20rcsin xarcsinx"
    )
    assert latex == r"\arcsin x"
    assert "rendered_segment_extracted" in warnings


def test_text_formula_extracts_rendered_norm_segment():
    latex, _, warnings = text_formula_to_latex(
        "\u3000\u2225v\u2225ESC_6844cfe6vESC_000c1603\u2225v\u2225"
    )
    assert latex == r"\Vert v\Vert"
    assert "rendered_segment_extracted" in warnings


def test_text_formula_extracts_big_operator_context():
    latex, _, warnings = text_formula_to_latex(
        "\u3000\u220fp\u00a0prime11\u2212p\u2212sESC_1e6e9ccbrod{p__ESCc9acd060____ESC5ccc1688__ext{prime}}__ESC71508eae__rac{1}{1-p^{-s}}p\u00a0prime\u220f\u200b1\u2212p\u2212s1\u200b"
    )
    assert latex == r"\prod_{p \text{prime}} \frac{1}{1-p^{-s}}"
    assert "big_operator_context_extracted" in warnings


def test_text_formula_recovers_log_product_identity():
    latex, _, _ = text_formula_to_latex(
        "log⁡(ab)=log⁡a+log⁡bESC_4d57c617og(ab)=ESC_565e3a1eog a+ESC_dc98dc77og blog(ab)=loga+logb"
    )
    assert latex == r"\log(ab)=\log a+\log b"


def test_text_formula_recovers_log_quotient_identity():
    latex, _, _ = text_formula_to_latex(
        "log⁡ab=log⁡a−log⁡bESC_a7f766d9ogESC_b8b71f86rac{a}{b}=ESC_40d258ddog a-ESC_48caf391og blogba\u200b=loga−logb"
    )
    assert latex == r"\log\frac{a}{b}=\log a-\log b"


def test_text_formula_recovers_log_power_identity():
    latex, _, _ = text_formula_to_latex(
        "log⁡an=nlog⁡aESC_e22999a5og a^n=nESC_a571d3a5og alogan=nloga"
    )
    assert latex == r"\log a^{n}=n\log a"


def test_text_formula_recovers_negative_power_identity():
    latex, _, _ = text_formula_to_latex(
        "a−n=1ana^{-n}=ESC_8375f515rac{1}{a^n}a−n=an1\u200b"
    )
    assert latex == r"a^{-n}=\frac{1}{a^{n}}"


def test_text_formula_recovers_log_base_identity():
    latex, _, _ = text_formula_to_latex(
        "log⁡ba=ln⁡aln⁡bESC_152821b2ogb a=__ESC71d1b47crac{ESC777d775a__n a}{__ESC415037da__n b}logb\u200ba=lnblna"
    )
    assert latex == r"\log_{b} a=\frac{\ln a}{\ln b}"


def test_text_formula_recovers_euler_identity_with_pi():
    latex, _, _ = text_formula_to_latex("eiπ+1=0")
    assert latex == r"e^{i\pi}+1=0"


def test_text_formula_recovers_exponential_series_prefix():
    latex, _, _ = text_formula_to_latex(
        "ex=∑n=0∞xnn!e^x=ESC_c78eca2bum{n=0}^{__ESC349c9c3cnfty}ESC_95cf8848__rac{x^n}{n!}ex=∑n=0∞\u200bn!xn\u200b"
    )
    assert latex == r"e^{x}=\sum_{n=0}^{\infty}\frac{x^{n}}{n!}"


def test_text_formula_dedups_repeated_standalone_expression():
    latex, _, warnings = text_formula_to_latex("(x+y+z)2(x+y+z)2(x+y+z)2")
    assert latex == r"(x+y+z)^{2}"
    assert "formula_repeated_segment_extracted" in warnings
    assert "lost_superscript_recovered" in warnings


def test_text_formula_dedups_repeated_power_sum_fragment():
    latex, _, warnings = text_formula_to_latex("a2+b2a2+b2a2+b2")
    assert latex == r"a^{2}+b^{2}"
    assert "formula_repeated_segment_extracted" in warnings
    assert "lost_superscript_recovered" in warnings


def test_text_formula_recovers_rendered_simple_fraction():
    latex, _, warnings = text_formula_to_latex("11−x11−x1−x1")
    assert latex == r"\frac{1}{1-x}"
    assert "rendered_fraction_reconstructed" in warnings


def test_text_formula_recovers_rendered_fraction_identity():
    latex, _, warnings = text_formula_to_latex("ab+cd=ad+bcbdab+cd=ad+bcbdba+dc=bdad+bc")
    assert latex == r"\frac{a}{b}+\frac{c}{d}=\frac{ad+bc}{bd}"
    assert "rendered_fraction_identity_reconstructed" in warnings


def test_text_formula_recovers_rendered_quadratic_formula():
    latex, _, warnings = text_formula_to_latex(
        "x=−b±b2−4ac2ax=−b±b2−4ac2ax=2a−b±b2−4ac"
    )
    assert latex == r"x=\frac{-b\pm\sqrt{b^{2}-4ac}}{2a}"
    assert "quadratic_formula_copy_reconstructed" in warnings


# ── lost superscript recovery ──


def test_lost_superscript_basic():
    """(a-b)2=a2-2ab+b2 → (a-b)^{2}=a^{2}-2ab+b^{2}"""
    latex, conf, warnings = text_formula_to_latex("(a-b)2=a2-2ab+b2")
    assert "^{2}" in latex
    assert "lost_superscript_recovered" in warnings
    # Coefficient '2' in '-2ab' must NOT become superscript
    assert "-2ab" in latex or "-2ab" in latex.replace("^{2}", "ZZ")


def test_lost_superscript_unicode_minus():
    """Same formula with U+2212 MINUS SIGN instead of ASCII hyphen."""
    latex, conf, warnings = text_formula_to_latex("(a\u2212b)2=a2\u22122ab+b2")
    assert "^{2}" in latex
    assert "lost_superscript_recovered" in warnings


def test_lost_superscript_needs_equals():
    """Structured standalone power sums should also recover."""
    latex, _, warnings = text_formula_to_latex("a2+b2")
    assert latex == "a^{2}+b^{2}"
    assert "lost_superscript_recovered" in warnings


def test_lost_superscript_needs_multiple_hits():
    """Single hit like E=mc2 should NOT trigger (< 2 matches threshold)."""
    latex, _, warnings = text_formula_to_latex("E=mc2")
    assert "lost_superscript_recovered" not in warnings


def test_lost_superscript_standalone_paren_power():
    latex, _, warnings = text_formula_to_latex("(x+y+z)2")
    assert latex == "(x+y+z)^{2}"
    assert "lost_superscript_recovered" in warnings


def test_lost_superscript_polynomial_term_with_context():
    latex, _, warnings = text_formula_to_latex("ax2+bx+c=0")
    assert latex == "ax^{2}+bx+c=0"
    assert "lost_superscript_recovered" in warnings


def test_lost_superscript_function_value_with_context():
    latex, _, warnings = text_formula_to_latex("y=ax2+bx+c")
    assert latex == "y=ax^{2}+bx+c"
    assert "lost_superscript_recovered" in warnings


def test_lost_superscript_coefficient_preserved():
    """Digits that are coefficients (preceded by operator) must not change."""
    latex, _, warnings = text_formula_to_latex("(a-b)2=a2-2ab+b2")
    # '-2ab' should remain intact (2 is a coefficient, not superscript)
    assert "-2ab" in latex


def test_lost_superscript_paren_exponent():
    """)2 should become )^{2}."""
    latex, _, warnings = text_formula_to_latex("(x+y)2=(x+y)(x+y)=x2+2xy+y2")
    assert "lost_superscript_recovered" in warnings
    assert ")^{2}" in latex


# ── _shape_from_latex ──


def test_shape_single_sup_still_works():
    shape, payload = _shape_from_latex("a^{2}")
    assert shape == "superscript"
    assert payload["base"] == "a"
    assert payload["sup"] == "2"


def test_shape_single_sub_still_works():
    shape, payload = _shape_from_latex("x_{i}")
    assert shape == "subscript"
    assert payload["base"] == "x"
    assert payload["sub"] == "i"


def test_shape_single_sub_sup_still_works():
    shape, payload = _shape_from_latex("x_{i}^{2}")
    assert shape == "sub_sup"
    assert payload["base"] == "x"
    assert payload["sub"] == "i"
    assert payload["sup"] == "2"


def test_shape_fraction_still_works():
    shape, payload = _shape_from_latex("\\frac{a}{b}")
    assert shape == "fraction"


def test_shape_sqrt_still_works():
    shape, payload = _shape_from_latex("\\sqrt{x}")
    assert shape == "sqrt"


def test_shape_big_operator_still_works():
    shape, payload = _shape_from_latex("\\sum_{i=1}^{n} x")
    assert shape == "big_operator"


def test_shape_compound_detected():
    shape, payload = _shape_from_latex("a^{2}+b^{2}=c^{2}")
    assert shape == "compound"
    assert payload["expr"] == "a^{2}+b^{2}=c^{2}"


def test_shape_emc2_compound():
    shape, _ = _shape_from_latex("E=mc^{2}")
    assert shape == "compound"


def test_shape_greedy_regression():
    """a^2+b^2=c^2 must NOT greedily match as single superscript."""
    # After normalization by caller, input would be a^{2}+b^{2}=c^{2}
    shape, payload = _shape_from_latex("a^{2}+b^{2}=c^{2}")
    assert shape != "superscript"


def test_shape_fraction_equation_is_compound():
    shape, payload = _shape_from_latex(r"\frac{a}{b}+\frac{c}{d}=\frac{ad+bc}{bd}")
    assert shape == "compound"


def test_shape_function_equation_is_compound():
    shape, payload = _shape_from_latex(r"\tan x=\frac{\sin x}{\cos x}")
    assert shape == "compound"


def test_shape_simple_text():
    shape, _ = _shape_from_latex("x + y = z")
    assert shape == "text"


def test_shape_empty():
    shape, _ = _shape_from_latex("")
    assert shape == "empty"


# ── _tokenize_latex ──


def test_tokenize_compound_sup():
    elements = _tokenize_latex("a^{2}+b^{2}=c^{2}")
    # Expected: sSup(a,2), text(+), sSup(b,2), text(=), sSup(c,2)
    assert len(elements) == 5
    assert elements[0].tag == f"{{{_M_NS}}}sSup"
    assert elements[1].tag == f"{{{_M_NS}}}r"
    assert elements[2].tag == f"{{{_M_NS}}}sSup"
    assert elements[3].tag == f"{{{_M_NS}}}r"
    assert elements[4].tag == f"{{{_M_NS}}}sSup"


def test_tokenize_emc2():
    elements = _tokenize_latex("E=mc^{2}")
    # Expected: text(E=m), sSup(c,2)
    assert len(elements) == 2
    assert elements[0].tag == f"{{{_M_NS}}}r"  # text run "E=m"
    assert elements[1].tag == f"{{{_M_NS}}}sSup"


def test_tokenize_frac_plus_text():
    elements = _tokenize_latex("\\frac{a}{b}+c")
    assert len(elements) == 2
    assert elements[0].tag == f"{{{_M_NS}}}f"  # fraction
    assert elements[1].tag == f"{{{_M_NS}}}r"  # text "+c"


def test_tokenize_sub_sup_combined():
    elements = _tokenize_latex("x_{i}^{2}")
    assert len(elements) == 1
    assert elements[0].tag == f"{{{_M_NS}}}sSubSup"


def test_tokenize_prescript_compound():
    elements = _tokenize_latex(r"\prescript{1}{}{\mathrm{O}_{2}}")
    assert len(elements) == 1
    assert elements[0].tag == f"{{{_M_NS}}}sPre"
    xml = etree.tostring(elements[0], encoding="unicode")
    assert 'sty ns0:val="p"' in xml or 'm:sty m:val="p"' in xml


def test_tokenize_sqrt_in_compound():
    elements = _tokenize_latex("\\sqrt{x}+1")
    assert len(elements) == 2
    assert elements[0].tag == f"{{{_M_NS}}}rad"
    assert elements[1].tag == f"{{{_M_NS}}}r"


def test_tokenize_plain_text():
    elements = _tokenize_latex("abc + 123")
    assert len(elements) == 1
    assert elements[0].tag == f"{{{_M_NS}}}r"


def test_tokenize_big_operator():
    elements = _tokenize_latex("\\sum_{i=1}^{n}")
    assert len(elements) == 1
    assert elements[0].tag == f"{{{_M_NS}}}nary"


def test_tokenize_function():
    elements = _tokenize_latex("\\sin{x}")
    assert len(elements) == 1
    assert elements[0].tag == f"{{{_M_NS}}}func"


def test_tokenize_extended_function():
    elements = _tokenize_latex("\\arcsin x")
    assert len(elements) == 1
    assert elements[0].tag == f"{{{_M_NS}}}func"


def test_tokenize_symbol_commands():
    elements = _tokenize_latex("\\nabla \\cdot F")
    assert len(elements) >= 1
    assert all(el.tag == f"{{{_M_NS}}}r" for el in elements)


# ── _build_compound_omml ──


def test_compound_omml_structure():
    omml = _build_compound_omml("a^{2}+b^{2}=c^{2}", block=False)
    assert omml.tag == f"{{{_M_NS}}}oMath"
    children = list(omml)
    assert len(children) == 5


def test_compound_omml_block_wraps():
    omml = _build_compound_omml("a^{2}+b^{2}", block=True)
    assert omml.tag == f"{{{_M_NS}}}oMathPara"
    omath = omml[0]
    assert omath.tag == f"{{{_M_NS}}}oMath"


# ── convert_formula_node end-to-end ──


def test_convert_compound_expression_produces_valid_omml():
    node = FormulaNode(
        kind="equation",
        payload={"latex": "a^{2}+b^{2}=c^{2}"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.omml_element is not None
    assert outcome.reason == "latex_to_word_compound"
    xml = etree.tostring(outcome.omml_element, encoding="unicode")
    assert "sSup" in xml


def test_convert_prescript_expression_produces_structured_omml():
    node = FormulaNode(
        kind="equation",
        payload={"latex": r"\prescript{1}{}{\mathrm{O}_{2}}"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.omml_element is not None
    xml = etree.tostring(outcome.omml_element, encoding="unicode")
    assert "sPre" in xml
    assert "sSub" in xml
    assert "sty" in xml


def test_convert_bare_caret_normalized_then_compound():
    node = FormulaNode(
        kind="equation",
        payload={"text": "a^2+b^2=c^2"},
        source_type="plain_text",
        confidence=0.70,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert "compound" in outcome.reason or "superscript" in outcome.reason


def test_convert_single_frac_still_uses_fast_path():
    node = FormulaNode(
        kind="equation",
        payload={"latex": "\\frac{x}{y}"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_fraction"


def test_convert_single_sup_fast_path():
    node = FormulaNode(
        kind="equation",
        payload={"latex": "x^{2}"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_superscript"


def test_convert_emc2():
    node = FormulaNode(
        kind="equation",
        payload={"latex": "E=mc^{2}"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    xml = etree.tostring(outcome.omml_element, encoding="unicode")
    assert "sSup" in xml


def test_convert_word_native_unchanged():
    node = FormulaNode(
        kind="equation",
        payload={"linear_text": "a+b"},
        source_type="word_native",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "already_word_native"
    assert outcome.transformed is False


def test_convert_extended_function():
    node = FormulaNode(
        kind="equation",
        payload={"latex": r"\arcsin x"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_function"


def test_convert_symbol_compound_expression():
    node = FormulaNode(
        kind="equation",
        payload={"latex": r"\nabla \cdot F"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_compound"


def test_convert_fraction_equation_stays_compound():
    node = FormulaNode(
        kind="equation",
        payload={"latex": r"\frac{a}{b}+\frac{c}{d}=\frac{ad+bc}{bd}"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_compound"
    xml = etree.tostring(outcome.omml_element, encoding="unicode")
    assert xml.count("f>") >= 3
    assert "=" in xml


def test_convert_function_fraction_equation_is_structured():
    node = FormulaNode(
        kind="equation",
        payload={"latex": r"\log\frac{a}{b}=\log a-\log b"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_compound"
    xml = etree.tostring(outcome.omml_element, encoding="unicode")
    assert "\\frac" not in xml
    assert "func" in xml
    assert "<ns0:f" in xml or "<m:f" in xml


def test_existing_single_shapes_unaffected():
    """All existing single-shape expressions should still produce the same shape."""
    cases = [
        ("\\frac{1}{2}", "fraction"),
        ("\\sqrt{x}", "sqrt"),
        ("x^{2}", "superscript"),
        ("x_{i}", "subscript"),
        ("x_{i}^{2}", "sub_sup"),
        ("\\sum_{i=1}^{n} x", "big_operator"),
        ("\\sin x", "function"),
        ("\\left( x \\right)", "delimited"),
        ("x + y", "text"),
        ("", "empty"),
    ]
    for expr, expected_shape in cases:
        shape, _ = _shape_from_latex(expr)
        assert shape == expected_shape, (
            f"Failed for {expr!r}: got {shape}, expected {expected_shape}"
        )


# ── Bug fixes: double-backslash, plain text functions, sum regex, symbols ──


def test_double_backslash_frac():
    """\\\\frac{x^2 + a_i}{1+b} should normalize and produce fraction."""
    node = FormulaNode(
        kind="equation",
        payload={"latex": "\\\\frac{x^2 + a_i}{1+b}"},
        source_type="latex",
        confidence=0.85,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_fraction"


def test_double_backslash_sqrt_nested_frac():
    """\\\\sqrt{x + \\\\frac{1}{n}} should normalize to sqrt."""
    node = FormulaNode(
        kind="equation",
        payload={"latex": "\\\\sqrt{x + \\\\frac{1}{n}}"},
        source_type="latex",
        confidence=0.85,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_sqrt"


def test_double_backslash_left_right():
    """\\\\left(\\\\frac{a+b}{c+d}\\\\right) should produce delimited."""
    node = FormulaNode(
        kind="equation",
        payload={"latex": "\\\\left(\\\\frac{a+b}{c+d}\\\\right)"},
        source_type="latex",
        confidence=0.85,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_delimited"


def test_double_backslash_bmatrix():
    """\\\\begin{bmatrix}...\\\\end{bmatrix} should produce matrix."""
    node = FormulaNode(
        kind="equation",
        payload={"latex": "\\\\begin{bmatrix}a & b \\\\\\\\ c & d\\\\end{bmatrix}"},
        source_type="latex",
        confidence=0.85,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    assert outcome.reason == "latex_to_word_matrix"


def test_plain_text_sin_becomes_command():
    """Plain text 'sin x' should be normalized to '\\sin x'."""
    latex, conf, _ = text_formula_to_latex("sin x + cos x")
    assert "\\sin" in latex
    assert "\\cos" in latex


def test_plain_text_lim_sin_over_x():
    """lim x->0 (sin x)/x = 1 should produce \\lim and \\sin."""
    latex, conf, _ = text_formula_to_latex("lim x->0 (sin x)/x = 1")
    assert "\\lim" in latex
    assert "\\sin" in latex
    assert "\\to" in latex


def test_plain_text_sum_underscore_paren():
    """sum_(i=1)^n i^2 should produce \\sum_{i=1}^{n}."""
    latex, conf, _ = text_formula_to_latex("sum_(i=1)^n i^2")
    assert "\\sum_{i=1}^{n}" in latex


def test_ge_le_symbol_normalized():
    """≥ and ≤ should be converted to \\ge and \\le."""
    from src.formula_core.normalize import _normalize_arrow_and_symbols
    assert "\\ge" in _normalize_arrow_and_symbols("a ≥ b")
    assert "\\le" in _normalize_arrow_and_symbols("a ≤ b")
    assert "\\pm" in _normalize_arrow_and_symbols("a ± b")
    assert "\\ne" in _normalize_arrow_and_symbols("a ≠ b")
    assert "\\approx" in _normalize_arrow_and_symbols("a ≈ b")


def test_latexify_does_not_double_escape():
    """Already-escaped \\sin should not become \\\\sin."""
    from src.formula_core.normalize import _latexify_plain_function_names
    assert _latexify_plain_function_names("\\sin x") == "\\sin x"
    assert _latexify_plain_function_names("sin x") == "\\sin x"


def test_tokenize_symbol_command_superscript_uses_symbol_as_base():
    elements = _tokenize_latex(r"\pi^{2}")
    assert len(elements) == 1
    assert elements[0].tag == f"{{{_M_NS}}}sSup"


def test_tokenize_fraction_base_superscript_uses_previous_element_as_base():
    elements = _tokenize_latex(r"\frac{1}{2}^{n}")
    assert len(elements) == 1
    assert elements[0].tag == f"{{{_M_NS}}}sSup"
    xml = etree.tostring(elements[0], encoding="unicode")
    assert "f>" in xml


def test_convert_fraction_with_pi_power_keeps_superscript_structured():
    node = FormulaNode(
        kind="equation",
        payload={"latex": r"\frac{\pi^2}{6}"},
        source_type="latex",
        confidence=0.95,
    )
    outcome = convert_formula_node(node, "word_native", block=True)
    assert outcome.success is True
    xml = etree.tostring(outcome.omml_element, encoding="unicode")
    assert "^{2}" not in xml
    assert "sSup" in xml


def test_text_formula_recovers_escaped_math_alphabet_family_commands():
    latex, _, warnings = text_formula_to_latex(
        "ESC_11111111athcal{F}(x)+ESC_22222222athtt{x}_0"
    )
    assert latex == r"\mathcal{F}(x)+\mathtt{x}_{0}"
    assert "escape_placeholder_formula_extracted" in warnings


def test_text_formula_converts_variant_greek_unicode_chars():
    latex, _, _ = text_formula_to_latex("ϵ+ϕ=ϑ")
    assert latex == r"\varepsilon+\varphi=\vartheta"


def test_text_formula_inserts_boundaries_for_glued_unicode_symbol_commands():
    samples = {
        "ΔG": r"\Delta G",
        "αx": r"\alpha x",
        "∞x": r"\infty x",
        "ηx": r"\eta x",
        "ψx": r"\psi x",
        "Γn": r"\Gamma n",
    }
    for raw, expected in samples.items():
        latex, _, _ = text_formula_to_latex(raw)
        assert latex == expected


def test_text_formula_converts_glued_unicode_operator_commands():
    samples = {
        "√x": r"\sqrt{x}",
        "→AB": r"\to AB",
        "∂G/∂T": r"\partial G/\partial T",
    }
    for raw, expected in samples.items():
        latex, _, _ = text_formula_to_latex(raw)
        assert latex == expected


def test_convert_plain_text_unicode_partial_expression_is_not_literal():
    outcome = convert_formula_node(
        FormulaNode(
            kind="equation",
            payload={"text": "∂G/∂T"},
            source_type="plain_text",
            confidence=0.80,
        ),
        "word_native",
        block=False,
    )
    assert outcome.success is True
    assert outcome.reason != "fallback_literal_conversion"
    assert outcome.omml_element is not None
    assert "\\" not in "".join(_math_texts(outcome.omml_element))


def test_convert_semantic_function_names_do_not_truncate_hyperbolic_prefixes():
    samples = [
        (r"\sinh x", "sinh", "x"),
        (r"\cosh x", "cosh", "x"),
        (r"\tanh x", "tanh", "x"),
        (r"\csc x", "csc", "x"),
        (r"\cot x", "cot", "x"),
    ]

    for latex, expected_name, expected_arg in samples:
        outcome = convert_formula_node(
            FormulaNode(
                kind="equation",
                payload={"latex": latex},
                source_type="latex",
                confidence=0.95,
            ),
            "word_native",
            block=False,
        )
        assert outcome.success is True
        assert outcome.omml_element is not None
        name_text = outcome.omml_element.xpath(
            "string(.//*[namespace-uri()='%s' and local-name()='fName'])" % _M_NS
        ).strip()
        arg_text = outcome.omml_element.xpath(
            "string(.//*[namespace-uri()='%s' and local-name()='e'])" % _M_NS
        ).strip()
        assert name_text == expected_name
        assert arg_text == expected_arg


def test_convert_math_alphabet_and_marker_commands_without_mathml_fallback(monkeypatch):
    monkeypatch.setattr(
        "src.formula_core.convert._convert_latex_via_mathml",
        lambda expr, block: None,
    )

    samples = [
        (r"\mathbf A", ["b"], []),
        (r"\boldsymbol x", ["bi"], []),
        (r"\mathrm d", ["p"], []),
        (r"\mathit F", ["i"], []),
        (r"\mathbb{R}", ["p"], ["double-struck"]),
        (r"\mathcal{F}", ["p"], ["script"]),
        (r"\mathfrak{g}", ["p"], ["fraktur"]),
        (r"\mathscr{L}", ["p"], ["script"]),
        (r"\mathsf{A}", ["p"], ["sans-serif"]),
        (r"\mathtt{x}", ["p"], ["monospace"]),
    ]

    for latex, expected_sty, expected_scr in samples:
        outcome = convert_formula_node(
            FormulaNode(
                kind="equation",
                payload={"latex": latex},
                source_type="latex",
                confidence=0.95,
            ),
            "word_native",
            block=False,
        )
        assert outcome.success is True
        assert outcome.omml_element is not None
        assert "\\" not in "".join(_math_texts(outcome.omml_element))
        sty_vals, scr_vals = _first_math_run_style(outcome.omml_element)
        assert sty_vals == expected_sty
        assert scr_vals == expected_scr


def test_convert_complex_group_styled_commands_without_mathml_fallback(monkeypatch):
    monkeypatch.setattr(
        "src.formula_core.convert._convert_latex_via_mathml",
        lambda expr, block: None,
    )

    samples = [
        (
            r"\mathsf{A+B}",
            [
                ("A", ["p"], ["sans-serif"]),
                ("+", [], []),
                ("B", ["p"], ["sans-serif"]),
            ],
            None,
        ),
        (
            r"\mathbf{\alpha+\beta}",
            [
                ("α", ["b"], []),
                ("+", [], []),
                ("β", ["b"], []),
            ],
            None,
        ),
        (
            r"\mathbb{R^n}",
            [
                ("R", ["p"], ["double-struck"]),
                ("n", ["p"], ["double-struck"]),
            ],
            "sSup",
        ),
        (
            r"\mathcal{AB_i}",
            [
                ("A", ["p"], ["script"]),
                ("B", ["p"], ["script"]),
                ("i", ["p"], ["script"]),
            ],
            "sSub",
        ),
    ]

    for latex, expected_runs, expected_tag in samples:
        outcome = convert_formula_node(
            FormulaNode(
                kind="equation",
                payload={"latex": latex},
                source_type="latex",
                confidence=0.95,
            ),
            "word_native",
            block=False,
        )
        assert outcome.success is True
        assert outcome.omml_element is not None
        assert "\\" not in "".join(_math_texts(outcome.omml_element))
        if expected_tag is not None:
            xml = etree.tostring(outcome.omml_element, encoding="unicode")
            assert expected_tag in xml
        assert _math_run_payloads(outcome.omml_element) == expected_runs


def test_convert_accent_commands_without_mathml_fallback(monkeypatch):
    monkeypatch.setattr(
        "src.formula_core.convert._convert_latex_via_mathml",
        lambda expr, block: None,
    )

    lim_upp_samples = [
        (r"\vec x", "x", "→"),
        (r"\hat\alpha", "α", "^"),
        (r"\overrightarrow{AB}", "AB", "→"),
    ]

    for latex, expected_body, expected_lim in lim_upp_samples:
        outcome = convert_formula_node(
            FormulaNode(
                kind="equation",
                payload={"latex": latex},
                source_type="latex",
                confidence=0.95,
            ),
            "word_native",
            block=False,
        )
        assert outcome.success is True
        assert outcome.omml_element is not None
        lim_text = outcome.omml_element.xpath(
            "string(.//*[namespace-uri()='%s' and local-name()='lim'])" % _M_NS
        ).strip()
        body_text = outcome.omml_element.xpath(
            "string(.//*[namespace-uri()='%s' and local-name()='e'])" % _M_NS
        ).strip()
        assert lim_text == expected_lim
        assert body_text == expected_body

    bar_outcome = convert_formula_node(
        FormulaNode(
            kind="equation",
            payload={"latex": r"\bar{x}"},
            source_type="latex",
            confidence=0.95,
        ),
        "word_native",
        block=False,
    )
    assert bar_outcome.success is True
    assert bar_outcome.omml_element is not None
    bar_pos_vals = bar_outcome.omml_element.xpath(
        ".//*[namespace-uri()='%s' and local-name()='pos']/@*[local-name()='val']"
        % _M_NS
    )
    bar_body_text = bar_outcome.omml_element.xpath(
        "string(.//*[namespace-uri()='%s' and local-name()='e'])" % _M_NS
    ).strip()
    assert bar_pos_vals == ["top"]
    assert bar_body_text == "x"

    overline_outcome = convert_formula_node(
        FormulaNode(
            kind="equation",
            payload={"latex": r"\overline{AB}"},
            source_type="latex",
            confidence=0.95,
        ),
        "word_native",
        block=False,
    )
    assert overline_outcome.success is True
    assert overline_outcome.omml_element is not None
    chr_vals = overline_outcome.omml_element.xpath(
        ".//*[namespace-uri()='%s' and local-name()='chr']/@*[local-name()='val']"
        % _M_NS
    )
    overline_body_text = overline_outcome.omml_element.xpath(
        "string(.//*[namespace-uri()='%s' and local-name()='e'])" % _M_NS
    ).strip()
    assert chr_vals == ["―"]
    assert overline_body_text == "AB"


def test_convert_uppercase_greek_commands_preserve_case_without_mathml_fallback(monkeypatch):
    monkeypatch.setattr(
        "src.formula_core.convert._convert_latex_via_mathml",
        lambda expr, block: None,
    )

    samples = [
        (r"\Gamma(s)", "Γ"),
        (r"\Delta G", "Δ"),
        (r"\Phi(x)", "Φ"),
        (r"\Psi(x)", "Ψ"),
        (r"\Pi_n", "Π"),
        (r"\Sigma_{i=1}^n", "Σ"),
    ]

    for latex, expected_text in samples:
        outcome = convert_formula_node(
            FormulaNode(
                kind="equation",
                payload={"latex": latex},
                source_type="latex",
                confidence=0.95,
            ),
            "word_native",
            block=False,
        )
        assert outcome.success is True
        assert outcome.omml_element is not None
        texts = "".join(_math_texts(outcome.omml_element))
        assert texts.startswith(expected_text)
        assert "\\" not in texts


def test_convert_extended_big_operators_without_mathml_fallback(monkeypatch):
    monkeypatch.setattr(
        "src.formula_core.convert._convert_latex_via_mathml",
        lambda expr, block: None,
    )

    samples = [
        (r"\iint_{D} f", "∬"),
        (r"\iiint_{V} g", "∭"),
        (r"\oint_{C} h", "∮"),
    ]

    for latex, expected_symbol in samples:
        outcome = convert_formula_node(
            FormulaNode(
                kind="equation",
                payload={"latex": latex},
                source_type="latex",
                confidence=0.95,
            ),
            "word_native",
            block=False,
        )
        assert outcome.success is True
        assert outcome.omml_element is not None
        chr_vals = outcome.omml_element.xpath(
            ".//*[namespace-uri()='%s' and local-name()='chr']/@*[local-name()='val']"
            % _M_NS
        )
        assert chr_vals == [expected_symbol]
