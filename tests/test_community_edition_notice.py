"""Community-edition anti-reseller notice and official-channel safeguards."""

from pathlib import Path

import docformat_gui as gui


PROJECT_ROOT = Path(__file__).parent.parent


def test_notice_schedule_uses_the_declared_community_edition_launches():
    for start_count in (1, 3, 7, 10, 20):
        assert gui.should_show_community_notice(start_count, is_pro=False)

    for start_count in (2, 4, 6, 8, 11, 21):
        assert not gui.should_show_community_notice(start_count, is_pro=False)

    assert not gui.should_show_community_notice(1, is_pro=True)


def test_notice_ticker_uses_the_declared_free_and_pro_message():
    assert gui.COMMUNITY_NOTICE_TEXT == (
        "本免费开源社区版仅限个人和非商业用途免费使用；未经许可，不得销售、收费分发或用于其他商业目的。"
        "如付费购买到本社区版本，请要求退款并举报商家。"
        "Pro 版已发布，提供红头文件支持、错别字和病句检查、自定义 AI 接入、智能 Agent 助手、模板管理、PDF 工具、自动更新、优先问题修复和使用支持。"
    )


def test_startup_notice_lists_pro_features_only_once():
    full_message = f"{gui.COMMUNITY_NOTICE_TEXT}\n\n{gui.COMMUNITY_NOTICE_DETAILS_FOOTER}"
    assert full_message.count("错别字和病句检查") == 1
    assert "了解 Pro 版详情" in full_message


def test_pro_details_dialog_has_room_and_a_scrollable_content_area():
    assert gui.PRO_INFO_DIALOG_SIZE[0] >= 860
    assert gui.PRO_INFO_DIALOG_SIZE[1] >= 760

    source = (PROJECT_ROOT / "docformat_gui.py").read_text(encoding="utf-8")
    dialog_source = source[source.index("class ProInfoDialog"):source.index("class CommunityNoticeTicker")]
    assert "ttk.Scrollbar" in dialog_source
    assert "self.resizable(True, True)" in dialog_source


def test_pro_information_is_embedded_and_bundled_for_all_builds():
    pro_page = (PROJECT_ROOT / "PRO.md").read_text(encoding="utf-8")
    for text in (
        gui.XIANYU_STORE_NAME,
        gui.XIANYU_STORE_URL,
        "小红书店铺：筹备中",
        "抖音商城：筹备中",
        "免费开源社区版（GitHub）",
        "PolyForm Noncommercial License 1.0.0",
        "仅限个人和非商业用途免费使用",
        "错别字和病句检查",
        "自定义 AI 接入",
        "智能 Agent 助手",
    ):
        assert text in pro_page
        assert text in gui.EMBEDDED_PRO_INFO_TEXT
    assert (PROJECT_ROOT / "assets" / "xianyu_qr.png").is_file()

    build_script = (PROJECT_ROOT / "build.py").read_text(encoding="utf-8")
    assert build_script.count("assets/xianyu_qr.png") == 3


def test_ui_keeps_the_notice_and_pro_details_available_without_a_browser():
    source = (PROJECT_ROOT / "docformat_gui.py").read_text(encoding="utf-8")
    for text in (
        "继续使用免费开源社区版",
        "了解 Pro 版",
        "打开闲鱼店铺",
        "此版本为免费开源社区版",
        "CommunityNoticeTicker",
        "ProInfoDialog",
        "AboutDialog",
    ):
        assert text in source
