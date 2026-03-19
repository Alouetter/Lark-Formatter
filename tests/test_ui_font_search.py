from src.ui.font_search import font_matches_query, normalize_font_search_text


def test_normalize_font_search_text_ignores_case_spaces_and_symbols():
    assert normalize_font_search_text(" Times New-Roman ") == "timesnewroman"


def test_font_matches_query_supports_contains_match():
    assert font_matches_query("Times New Roman", "times")
    assert font_matches_query("Times New Roman", "new roman")


def test_font_matches_query_supports_subsequence_match():
    assert font_matches_query("Times New Roman", "tnr")


def test_font_matches_query_handles_fullwidth_input():
    assert font_matches_query("Times New Roman", "Ｔｉｍｅｓ")


def test_font_matches_query_rejects_unrelated_fonts():
    assert not font_matches_query("宋体", "times")
