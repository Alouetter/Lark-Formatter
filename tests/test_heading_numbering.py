from src.engine.rules.heading_numbering import split_heading_text


def test_split_heading_text_chinese_chapter():
    num, sep, title = split_heading_text("第一章 绪论")
    assert num == "第一章"
    assert sep == " "
    assert title == "绪论"


def test_split_heading_text_arabic_multilevel():
    num, sep, title = split_heading_text("1.2.3\t测试标题")
    assert num == "1.2.3"
    assert sep == "\t"
    assert title == "测试标题"


def test_split_heading_text_keeps_non_heading_numbers():
    num, sep, title = split_heading_text("2024年工作总结")
    assert num == ""
    assert sep == ""
    assert title == "2024年工作总结"
