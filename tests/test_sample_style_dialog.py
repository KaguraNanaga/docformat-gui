"""Preview-card mapping tests for local sample-style learning."""

from sample_style_dialog import build_style_cards


def _settings():
    return {
        "page": {
            "top": 2.5, "bottom": 2.5, "left": 2.8, "right": 2.6,
            "width_mm": 210, "height_mm": 297, "orientation": "portrait",
        },
        "page_number": True,
        "title": {"font_cn": "方正小标宋简体", "size": 22, "bold": False,
                  "align": "center", "indent": 0, "line_spacing": 33},
        "heading1": {"font_cn": "黑体", "size": 16, "bold": False,
                     "align": "left", "indent": 32, "line_spacing": 28},
        "body": {"font_cn": "仿宋_GB2312", "size": 16, "bold": False,
                 "align": "justify", "indent": 32, "line_spacing": 28},
        "signature": {"font_cn": "仿宋_GB2312", "size": 16, "align": "right"},
    }


def test_style_cards_show_only_reusable_roles_and_page_summary():
    cards = build_style_cards(_settings())
    assert [row["key"] for row in cards["rows"]] == ["title", "heading1", "body"]
    body = next(row for row in cards["rows"] if row["key"] == "body")
    assert "首行缩进 2 字" in body["summary"]
    assert "行距 28 磅" in body["summary"]
    assert "A4" in cards["page_row"]["summary"]
    assert "竖向" in cards["page_row"]["summary"]


def test_missing_fonts_and_warnings_are_presented_without_changing_settings():
    cards = build_style_cards(
        _settings(), warnings=["多节文档只读取第一节", "  "], available_fonts={"黑体"},
    )
    by_key = {row["key"]: row for row in cards["rows"]}
    assert by_key["heading1"]["font_missing"] is False
    assert by_key["title"]["font_missing"] is True
    assert cards["warnings"] == ["多节文档只读取第一节"]
