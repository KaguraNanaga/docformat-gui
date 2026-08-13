"""Community-edition notice and repository-content safeguards."""

from pathlib import Path

import docformat_gui as gui


PROJECT_ROOT = Path(__file__).parent.parent


def test_notice_schedule_uses_the_declared_community_edition_launches():
    for start_count in (1, 3, 7, 10, 20):
        assert gui.should_show_community_notice(start_count)

    for start_count in (2, 4, 6, 8, 11, 21):
        assert not gui.should_show_community_notice(start_count)


def test_notice_only_contains_the_free_community_disclosure():
    assert gui.COMMUNITY_NOTICE_TEXT == (
        "本免费开源版仅限个人和非商业用途免费使用；未经许可，不得销售、收费分发或用于其他商业目的。"
        "如付费购买到本免费开源版本，请要求退款并举报商家。"
    )


def test_repository_does_not_ship_commercial_promotion_surfaces():
    assert not (PROJECT_ROOT / ("P" + "RO.md")).exists()
    assert not (PROJECT_ROOT / "assets" / ("xian" + "yu_qr.png")).exists()

    forbidden = (
        "P" + "ro 版",
        "P" + "RO.md",
        "闲" + "鱼",
        "goo" + "fish.com",
        "xian" + "yu_qr",
    )
    for relative_path in (
        "docformat_gui.py",
        "README.md",
        "build.py",
        ".github/workflows/build.yml",
    ):
        source = (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")
        for text in forbidden:
            assert text not in source


def test_ui_keeps_the_community_notice_and_about_dialog():
    source = (PROJECT_ROOT / "docformat_gui.py").read_text(encoding="utf-8")
    for text in (
        "继续使用免费开源版",
        "此版本为免费开源版",
        "CommunityNoticeTicker",
        "AboutDialog",
    ):
        assert text in source
