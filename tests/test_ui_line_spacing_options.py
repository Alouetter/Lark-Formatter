from src.ui.line_spacing_options import (
    LINE_SPACING_OPTIONS,
    line_spacing_is_editable,
    line_spacing_unit_label,
    normalize_line_spacing_type,
    resolve_line_spacing_value,
)


def test_line_spacing_options_have_expected_order():
    assert LINE_SPACING_OPTIONS == (
        ("exact", "固定值"),
        ("single", "单倍"),
        ("one_half", "1.5 倍"),
        ("double", "双倍"),
        ("multiple", "多倍"),
    )


def test_normalize_line_spacing_type_supports_labels():
    assert normalize_line_spacing_type("固定值") == "exact"
    assert normalize_line_spacing_type("单倍") == "single"
    assert normalize_line_spacing_type("多倍") == "multiple"


def test_line_spacing_unit_and_editable_behavior():
    assert line_spacing_unit_label("exact") == "磅"
    assert line_spacing_unit_label("single") == "倍"
    assert line_spacing_is_editable("exact")
    assert not line_spacing_is_editable("single")
    assert line_spacing_is_editable("multiple")


def test_resolve_line_spacing_value_handles_fixed_and_editable_types():
    assert resolve_line_spacing_value("single", 99) == 1.0
    assert resolve_line_spacing_value("one_half", 99) == 1.5
    assert resolve_line_spacing_value("double", 99) == 2.0
    assert resolve_line_spacing_value("exact", 16) == 16.0
    assert resolve_line_spacing_value("multiple", 1.8) == 1.8
