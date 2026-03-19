from src.formula_core.repair import (
    extract_repeated_formula_candidate,
    repair_formula_text,
)


def test_repair_engine_extracts_repeated_formula_segment():
    outcome = repair_formula_text("a+b=ca+b=c")
    assert outcome.source == "repeated_segment"
    assert outcome.text == "a+b=c"
    assert "formula_repeated_segment_extracted" in outcome.warnings


def test_repair_engine_recovers_escaped_fraction_chain():
    outcome = repair_formula_text(
        "ab+cd=ad+bcbdESC_a9cae316rac{a}{b}+ESC_3a30d405rac{c}{d}=ESC_8f7ff319rac{ad+bc}{bd}ba\u200b+dc\u200b=bdad+bc\u200b"
    )
    assert outcome.source == "escaped_extracted"
    assert outcome.text == r"\frac{a}{b}+\frac{c}{d}=\frac{ad+bc}{bd}"
    assert "escape_placeholder_command_recovered" in outcome.warnings
    assert "escape_placeholder_formula_extracted" in outcome.warnings
    assert outcome.candidates


def test_repair_engine_recovers_escaped_single_fraction():
    outcome = repair_formula_text("11−xESC_57aa258drac{1}{1-x}1−x1\u200b")
    assert outcome.source == "escaped_extracted"
    assert outcome.text == r"\frac{1}{1-x}"


def test_repair_engine_recovers_escaped_math_alphabet_family_commands():
    outcome = repair_formula_text("ESC_11111111athcal{F}(x)+ESC_22222222athtt{x}_0")
    assert outcome.source == "escaped_extracted"
    assert outcome.text == r"\mathcal{F}(x)+\mathtt{x}_0"
    assert "escape_placeholder_command_recovered" in outcome.warnings


def test_repair_engine_recovers_escaped_sans_serif_and_script_commands():
    outcome = repair_formula_text("ESC_33333333athsf{A}+ESC_44444444athscr{L}")
    assert outcome.source == "escaped_extracted"
    assert outcome.text == r"\mathsf{A}+\mathscr{L}"
    assert "escape_placeholder_command_recovered" in outcome.warnings


def test_extract_repeated_formula_candidate_prefers_dominant_segment():
    best = extract_repeated_formula_candidate("x^2+1=0x^{2}+1=0")
    assert best in {"x^2+1=0", "x^{2}+1=0"}


def test_repair_engine_extracts_repeated_expression_segment():
    outcome = repair_formula_text("(x+y+z)2(x+y+z)2(x+y+z)2")
    assert outcome.source == "repeated_segment"
    assert outcome.text == "(x+y+z)2"
    assert "formula_repeated_segment_extracted" in outcome.warnings


def test_repair_engine_extracts_rendered_short_command_segment():
    outcome = repair_formula_text("\u3000arcsin\u2061xESC_bd6e0b20rcsin xarcsinx")
    assert outcome.text == r"\arcsin x"
    assert "rendered_segment_extracted" in outcome.warnings


def test_repair_engine_extracts_rendered_norm_segment():
    outcome = repair_formula_text("\u3000\u2225v\u2225ESC_6844cfe6vESC_000c1603\u2225v\u2225")
    assert outcome.text == r"\Vert v\Vert"
    assert "rendered_segment_extracted" in outcome.warnings


def test_repair_engine_extracts_glued_rendered_unicode_symbol_commands():
    samples = {
        "ΔG": r"\Delta G",
        "Γn": r"\Gamma n",
        "√x": r"\sqrt x",
        "∂x": r"\partial x",
        "∞x": r"\infty x",
        "→AB": r"\to AB",
    }
    for raw, expected in samples.items():
        outcome = repair_formula_text(raw)
        assert outcome.source == "rendered_segment"
        assert outcome.text == expected
        assert "rendered_segment_extracted" in outcome.warnings


def test_repair_engine_recovers_log_product_identity():
    outcome = repair_formula_text(
        "log⁡(ab)=log⁡a+log⁡bESC_4d57c617og(ab)=ESC_565e3a1eog a+ESC_dc98dc77og blog(ab)=loga+logb"
    )
    assert outcome.text == r"\log(ab)=\log a+\log b"


def test_repair_engine_recovers_log_quotient_identity():
    outcome = repair_formula_text(
        "log⁡ab=log⁡a−log⁡bESC_a7f766d9ogESC_b8b71f86rac{a}{b}=ESC_40d258ddog a-ESC_48caf391og blogba\u200b=loga−logb"
    )
    assert outcome.text == r"\log\frac{a}{b}=\log a-\log b"


def test_repair_engine_recovers_log_power_identity():
    outcome = repair_formula_text(
        "log⁡an=nlog⁡aESC_e22999a5og a^n=nESC_a571d3a5og alogan=nloga"
    )
    assert outcome.text == r"\log a^n=n\log a"


def test_repair_engine_recovers_negative_power_identity():
    outcome = repair_formula_text(
        "a−n=1ana^{-n}=ESC_8375f515rac{1}{a^n}a−n=an1\u200b"
    )
    assert outcome.text == r"a^{-n}=\frac{1}{a^n}"


def test_repair_engine_recovers_log_base_identity():
    outcome = repair_formula_text(
        "log⁡ba=ln⁡aln⁡bESC_152821b2ogb a=__ESC71d1b47crac{ESC777d775a__n a}{__ESC415037da__n b}logb\u200ba=lnblna"
    )
    assert outcome.text == r"\log_{b} a=\frac{\ln a}{\ln b}"



def test_repair_engine_recovers_rendered_simple_fraction():
    outcome = repair_formula_text("11\u2212x11\u2212x1\u2212x1")
    assert outcome.text == r"\frac{1}{1-x}"
    assert "rendered_fraction_reconstructed" in outcome.warnings


def test_repair_engine_recovers_rendered_fraction_identity():
    outcome = repair_formula_text("ab+cd=ad+bcbdab+cd=ad+bcbdba+dc=bdad+bc")
    assert outcome.text == r"\frac{a}{b}+\frac{c}{d}=\frac{ad+bc}{bd}"
    assert "rendered_fraction_identity_reconstructed" in outcome.warnings


def test_repair_engine_recovers_rendered_quadratic_formula():
    outcome = repair_formula_text(
        "x=\u2212b\u00b1b2\u22124ac2ax=\u2212b\u00b1b2\u22124ac2ax=2a\u2212b\u00b1b2\u22124ac"
    )
    assert outcome.text == r"x=\frac{-b\pm\sqrt{b^{2}-4ac}}{2a}"
    assert "quadratic_formula_copy_reconstructed" in outcome.warnings
