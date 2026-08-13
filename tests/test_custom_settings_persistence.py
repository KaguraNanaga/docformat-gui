from copy import deepcopy
from pathlib import Path

from docx import Document

import docformat_gui
from scripts import formatter
from scripts.formatter import PRESETS, format_document


def test_gui_saves_and_loads_active_custom_preset(tmp_path, monkeypatch):
    config_file = tmp_path / "custom_settings.json"
    monkeypatch.setattr(docformat_gui, "CONFIG_FILE", config_file)

    preset_a = deepcopy(docformat_gui.DEFAULT_CUSTOM_SETTINGS)
    preset_a.update({"id": "preset-a", "name": "A"})
    preset_a["page"]["left"] = 1.1

    preset_b = deepcopy(docformat_gui.DEFAULT_CUSTOM_SETTINGS)
    preset_b.update({"id": "preset-b", "name": "B"})
    preset_b["page"]["left"] = 2.2

    docformat_gui.save_custom_settings({
        "schema_version": docformat_gui.CONFIG_SCHEMA_VERSION,
        "active_preset_id": "preset-b",
        "presets": [preset_a, preset_b],
    })

    loaded = docformat_gui.load_custom_settings()
    active = docformat_gui.get_active_user_preset(loaded)

    assert loaded["active_preset_id"] == "preset-b"
    assert active["name"] == "B"
    assert active["page"]["left"] == 2.2


def test_formatter_reads_active_custom_preset_from_user_config(tmp_path, monkeypatch):
    appdata = tmp_path / "appdata"
    config_dir = appdata / "DocFormatter"
    config_dir.mkdir(parents=True)
    monkeypatch.setenv("APPDATA", str(appdata))
    monkeypatch.setattr(formatter.sys, "platform", "win32")

    preset = deepcopy(PRESETS["official"])
    preset.update({"id": "custom-active", "name": "用户保存的自定义格式"})
    preset["page"] = {"top": 1.1, "bottom": 1.2, "left": 1.3, "right": 1.4}

    monkeypatch.setattr(docformat_gui, "CONFIG_FILE", config_dir / "custom_settings.json")
    docformat_gui.save_custom_settings({
        "schema_version": docformat_gui.CONFIG_SCHEMA_VERSION,
        "active_preset_id": "custom-active",
        "presets": [preset],
    })

    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    doc = Document()
    doc.add_paragraph("这是正文内容。")
    doc.save(source)

    format_document(str(source), str(output), preset_name="custom")

    result = Document(output)
    section = result.sections[0]
    assert abs(section.top_margin.cm - 1.1) < 0.02
    assert abs(section.bottom_margin.cm - 1.2) < 0.02
    assert abs(section.left_margin.cm - 1.3) < 0.02
    assert abs(section.right_margin.cm - 1.4) < 0.02


def test_custom_paragraph_rules_survive_save_and_reload(tmp_path, monkeypatch):
    config_file = tmp_path / "custom_settings.json"
    monkeypatch.setattr(docformat_gui, "CONFIG_FILE", config_file)
    preset = deepcopy(docformat_gui.DEFAULT_CUSTOM_SETTINGS)
    preset.update({
        "id": "generic-local-rules",
        "name": "本地规则模板",
        "layout_mode": "generic",
        "paragraph_rule_version": 1,
        "paragraph_types": [{
            "id": "note",
            "name": "提示段",
            "format": {
                "font_cn": "黑体", "font_en": "Arial", "size": 18,
                "bold": True, "align": "center", "indent": 0,
                "line_spacing": 24, "space_before": 0, "space_after": 0,
            },
        }],
        "paragraph_rules": [{
            "id": "rule_note", "name": "提示前缀", "type_id": "note",
            "pattern": r"^提示：", "previous_type": "", "next_pattern": "",
            "next_next_pattern": "", "priority": 500, "enabled": True,
        }],
    })
    docformat_gui.save_custom_settings({
        "schema_version": docformat_gui.CONFIG_SCHEMA_VERSION,
        "active_preset_id": preset["id"],
        "presets": [preset],
    })

    loaded = docformat_gui.get_active_user_preset(docformat_gui.load_custom_settings())

    assert loaded["paragraph_types"][0]["id"] == "note"
    assert loaded["paragraph_rules"][0]["pattern"] == r"^提示："

    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    document = Document()
    document.add_paragraph("提示：请认真核对数据。")
    document.save(source)
    format_document(str(source), str(output), "custom", custom_settings=loaded)
    paragraph = Document(output).paragraphs[0]
    assert paragraph.alignment == 1
    assert paragraph.runs[0].bold is True
    assert paragraph.runs[0].font.size.pt == 18
