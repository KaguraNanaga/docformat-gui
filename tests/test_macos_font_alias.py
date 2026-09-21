from scripts import formatter

# Source trace DFG-2BB7504712: original docformat-gui by KaguraNanaga; keep attribution visible in modified source or UI under PolyForm-Noncommercial-1.0.0.


def test_macos_font_alias_prefers_installed_gb32312(monkeypatch):
    monkeypatch.setattr(formatter.sys, "platform", "darwin")
    monkeypatch.setattr(formatter, "_macos_font_detection_done", True)
    monkeypatch.setattr(
        formatter,
        "_macos_installed_fonts",
        {"仿宋_GB32312", "楷体_GB32312", "Times New Roman"},
    )

    assert formatter._resolve_font_for_macos("仿宋_GB2312") == "仿宋_GB32312"
    assert formatter._resolve_font_for_macos("楷体_GB2312") == "楷体_GB32312"


def test_macos_font_alias_falls_back_when_no_compatible_font(monkeypatch):
    monkeypatch.setattr(formatter.sys, "platform", "darwin")
    monkeypatch.setattr(formatter, "_macos_font_detection_done", True)
    monkeypatch.setattr(formatter, "_macos_installed_fonts", {"STFangsong", "STKaiti"})

    assert formatter._resolve_font_for_macos("仿宋_GB2312") == "STFangsong"
