"""Community-edition anti-reseller notice and official-channel safeguards."""

from pathlib import Path

import docformat_gui as gui


PROJECT_ROOT = Path(__file__).parent.parent


def test_notice_schedule_is_low_frequency_and_never_shows_for_pro():
    day = 24 * 60 * 60
    assert gui.should_show_community_notice(1, None, 100 * day, is_pro=False)
    assert not gui.should_show_community_notice(4, 99 * day, 100 * day, is_pro=False)
    assert gui.should_show_community_notice(4, 84 * day, 100 * day, is_pro=False)
    assert gui.should_show_community_notice(10, 0, 100 * day, is_pro=False)
    assert not gui.should_show_community_notice(11, 0, 100 * day, is_pro=False)
    assert gui.should_show_community_notice(30, 0, 100 * day, is_pro=False)
    assert not gui.should_show_community_notice(30, 0, 100 * day, is_pro=True)


def test_official_channel_is_documented_and_bundled_for_all_builds():
    pro_page = (PROJECT_ROOT / "PRO.md").read_text(encoding="utf-8")
    assert gui.XIANYU_STORE_NAME in pro_page
    assert gui.XIANYU_STORE_URL in pro_page
    assert (PROJECT_ROOT / "assets" / "xianyu_qr.png").is_file()

    build_script = (PROJECT_ROOT / "build.py").read_text(encoding="utf-8")
    assert build_script.count("assets/xianyu_qr.png") == 3


def test_ui_keeps_the_notice_and_banner_non_blocking():
    source = (PROJECT_ROOT / "docformat_gui.py").read_text(encoding="utf-8")
    for text in ("继续使用社区版", "了解Pro版", "打开闲鱼店铺", "免费开源社区版"):
        assert text in source
