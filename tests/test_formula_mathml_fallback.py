import pytest
from docx import Document
from lxml import etree

from src.engine.change_tracker import ChangeTracker
from src.engine.rules.formula_convert import FormulaConvertRule
from src.formula_core.ast import FormulaNode
from src.formula_core.convert import (
    _has_mathml_to_omml_support,
    convert_formula_node,
)
from src.scene.schema import SceneConfig

_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_COMPLEX_HYPERGEOMETRIC = r"""
{}_pF_q\!\left(
\begin{matrix}
a_1,\dots,a_p\\
b_1,\dots,b_q
\end{matrix}
\,\middle|\, z
\right)
=
\sum_{n=0}^{\infty}
\frac{
(a_1)_n\cdots(a_p)_n
}{
(b_1)_n\cdots(b_q)_n
}
\frac{z^n}{n!}
""".strip()

_COMPLEX_NESTED_DELIMITERS = r"""
\left\langle
0
\left|
\mathcal{T}
\left\{
\exp\!\left(
i\int d^4x\,J(x)\hat{\phi}(x)
\right)
\right\}
\right|
0
\right\rangle
=
\exp\!\left(
-\frac{1}{2}
\int d^4x\,d^4y\,
J(x)\,\Delta_F(x-y)\,J(y)
\right)
""".strip()


def _math_texts(omml_element) -> list[str]:
    return [
        str(text or "")
        for text in omml_element.xpath(
            ".//*[namespace-uri()='%s' and local-name()='t']/text()" % _M_NS
        )
    ]


pytestmark = pytest.mark.skipif(
    not _has_mathml_to_omml_support(),
    reason="MathML to OMML support not available",
)


def test_convert_complex_latex_uses_mathml_fallback_without_literal_commands():
    node = FormulaNode(
        kind="equation",
        payload={"latex": _COMPLEX_HYPERGEOMETRIC},
        source_type="latex",
        confidence=0.95,
    )

    outcome = convert_formula_node(node, "word_native", block=True)

    assert outcome.success is True
    assert outcome.reason == "latex_to_word_mathml"
    texts = _math_texts(outcome.omml_element)
    assert texts
    assert not any("\\" in text for text in texts)


def test_formula_convert_rule_handles_bare_complex_latex_paragraph():
    doc = Document()
    para = doc.add_paragraph(_COMPLEX_NESTED_DELIMITERS)

    cfg = SceneConfig()
    cfg.formula_convert.enabled = True

    tracker = ChangeTracker()
    context = {}
    FormulaConvertRule().apply(doc, cfg, tracker, context)

    assert "oMath" in para._p.xml
    assert "\\" not in para._p.xml

    stats = context.get("formula_runtime", {}).get("stats", {}).get("formula_convert", {})
    assert stats.get("converted") == 1


def test_multiline_display_blocks_prefer_mathml_primary_over_custom_compound_builder():
    expr = r"""
\left(
\prod_{j=1}^{m}
\int_{\gamma_j}
\frac{dz_j}{2\pi i}
\right)
\frac{
\prod_{1\le i<j\le m}(z_i-z_j)^2
}{
\prod_{j=1}^{m}
\left[
(z_j-a)^{\alpha_j+1}(z_j-b)^{\beta_j+1}
\right]
}
""".strip()
    node = FormulaNode(
        kind="equation",
        payload={"latex": expr, "latex_source": "multiline_display_block"},
        source_type="latex",
        confidence=0.95,
    )

    outcome = convert_formula_node(node, "word_native", block=True)

    assert outcome.success is True
    assert outcome.reason == "latex_to_word_mathml_primary"
    first_nary = outcome.omml_element.xpath(
        ".//*[namespace-uri()='%s' and local-name()='nary']" % _M_NS
    )[0]
    first_e = first_nary.xpath(
        "./*[namespace-uri()='%s' and local-name()='e']" % _M_NS
    )[0]
    nested_objects = [
        child for child in first_e
        if etree.QName(child).localname in {"nary", "f", "d", "r", "sSub", "sSup", "sSubSup"}
    ]
    assert nested_objects
    assert any(etree.QName(child).localname != "r" for child in nested_objects)
