from src.ui.alignment_options import (
    ALIGNMENT_OPTIONS,
    alignment_display_label,
    normalize_alignment_value,
)


def test_alignment_options_have_expected_order_and_labels():
    assert ALIGNMENT_OPTIONS == (
        ("left", "左对齐"),
        ("right", "右对齐"),
        ("center", "居中对齐"),
        ("justify", "两端对齐"),
    )


def test_normalize_alignment_value_supports_internal_and_display_values():
    assert normalize_alignment_value("left") == "left"
    assert normalize_alignment_value("右对齐") == "right"
    assert normalize_alignment_value("居中对齐") == "center"
    assert normalize_alignment_value("垂直居中") == "center"
    assert normalize_alignment_value("两端对齐") == "justify"


def test_normalize_alignment_value_defaults_empty_to_justify():
    assert normalize_alignment_value("") == "justify"


def test_alignment_display_label_maps_internal_values():
    assert alignment_display_label("left") == "左对齐"
    assert alignment_display_label("center") == "居中对齐"
