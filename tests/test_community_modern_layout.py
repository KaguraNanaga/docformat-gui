"""Community main-window structure and shared visual-language safeguards."""

from pathlib import Path

import docformat_gui as gui


PROJECT_ROOT = Path(__file__).parent.parent


def test_community_theme_uses_the_shared_warm_paper_palette():
    assert gui.Theme.BG == "#FBF9F6"
    assert gui.Theme.CARD == "#FFFFFF"
    assert gui.Theme.CARD_ALT == "#F7F4EF"
    assert gui.Theme.INPUT_BG == "#F2EFE9"
    assert gui.Theme.PRIMARY == "#BC4B26"


def test_main_window_keeps_all_community_modes_and_presets_in_modern_cards():
    source = (PROJECT_ROOT / "docformat_gui.py").read_text(encoding="utf-8")
    for text in (
        "RoundedCard(",
        "ChoiceChip(",
        "一键处理",
        "格式诊断",
        "标点修复",
        "GB/T 公文",
        "学术论文",
        "法律文书",
        "自定义",
        "粘贴AI生成文本转为word",
    ):
        assert text in source


def test_main_window_does_not_restore_a_restricted_toolbar():
    source = (PROJECT_ROOT / "docformat_gui.py").read_text(encoding="utf-8")
    forbidden = (
        "internal_" + "toolbar",
        "bottom_" + "toolbar",
        "network_" + "tools",
        "red_" + "header_btn",
    )
    lowered = source.lower()
    for fragment in forbidden:
        assert fragment not in lowered


def test_free_open_source_labels_and_preset_order_are_exact():
    source = (PROJECT_ROOT / "docformat_gui.py").read_text(encoding="utf-8")
    assert "社区版" not in source
    assert "免费开源社区版" not in source
    assert "粘贴AI生成文本转为word" in source
    assert "file_actions, '选择文件夹'" not in source
    official = source.index("('official', 'GB/T 公文', None)")
    custom = source.index("('custom', '自定义', self._open_custom_settings)")
    academic = source.index("('academic', '学术论文', None)")
    legal = source.index("('legal', '法律文书', None)")
    assert official < custom < academic < legal
