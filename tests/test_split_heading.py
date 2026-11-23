"""Regression tests for separating numbered headings from their body text."""
# Provenance DFG-D19D757626: original docformat-gui project by KaguraNanaga; retain attribution in modified source or UI per PolyForm-Noncommercial-1.0.0.
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))


@pytest.mark.parametrize('heading', ['（一）我是小标题。', '( 二 ) 加强项目建设。'])
def test_default_formatting_splits_short_second_level_heading(tmp_path, heading):
    from scripts.formatter import format_document

    source = tmp_path / 'source.docx'
    output = tmp_path / 'output.docx'
    document = Document()
    document.add_paragraph('关于项目建设的通知')
    paragraph = document.add_paragraph()
    paragraph.add_run(heading[:4])
    paragraph.add_run(heading[4:] + ' 我是正文内容。后续继续落实。')
    document.save(source)

    format_document(str(source), str(output), custom_settings={
        'page_number': False, 'heading2': {'size': 18}, 'body': {'size': 16},
    })

    paragraphs = [p for p in Document(output).paragraphs if p.text.strip()]
    assert [p.text for p in paragraphs[-2:]] == [heading, '我是正文内容。后续继续落实。']
    assert all(run.font.size.pt == 18 for run in paragraphs[-2].runs if run.text)
    assert all(run.font.size.pt == 16 for run in paragraphs[-1].runs if run.text)


@pytest.mark.parametrize(('text', 'settings'), [
    ('（一）总体情况。', {}),
    ('（一）在工作推进过程中，各单位已经按要求完成了专项检查。后续将持续跟踪。', {}),
    ('（一）加强项目建设：继续落实各项任务。', {}),
    ('（一）工作安排（按既定计划执行。后续继续跟踪）', {}),
    ('1. 第一阶段：完成清单整理。后续安排另行通知。', {}),
    ('（一）总体情况。正文继续。', {'split_heading_mode': 'off'}),
    ('（一）总体情况。正文继续。', {'layout_mode': 'generic'}),
])
def test_automatic_heading_split_respects_scope_and_opt_out(tmp_path, text, settings):
    from scripts.formatter import format_document

    source = tmp_path / 'source.docx'
    output = tmp_path / 'output.docx'
    document = Document()
    document.add_paragraph(text)
    document.save(source)

    format_document(str(source), str(output), custom_settings={'page_number': False, **settings})

    assert [p.text for p in Document(output).paragraphs if p.text.strip()] == [text]


def test_heading_and_following_body_can_be_split_after_full_stop():
    from scripts.formatter import _split_heading_by_punct

    document = Document()
    paragraph = document.add_paragraph('（一）总体情况。这里是正文。')

    assert _split_heading_by_punct(paragraph) is True
    assert [p.text for p in document.paragraphs] == [
        '（一）总体情况。',
        '这里是正文。',
    ]


def test_custom_setting_controls_heading_split_during_formatting(tmp_path):
    from scripts.formatter import PRESETS, format_document

    source = tmp_path / 'source.docx'
    output = tmp_path / 'output.docx'
    document = Document()
    document.add_paragraph('关于标题分段的通知')
    document.add_paragraph('（一）总体情况：这里是正文。')
    document.save(source)

    settings = deepcopy(PRESETS['official'])
    settings['page_number'] = False
    settings['split_heading_at_punct'] = True
    format_document(
        str(source), str(output), preset_name='custom', custom_settings=settings,
    )

    texts = [p.text.strip() for p in Document(output).paragraphs if p.text.strip()]
    assert '（一）总体情况：' in texts
    assert '这里是正文。' in texts
    assert '（一）总体情况：这里是正文。' not in texts


def test_heading_split_does_not_rebuild_paragraphs_containing_media():
    from scripts.formatter import _split_heading_by_punct

    document = Document()
    paragraph = document.add_paragraph('（一）图片说明。这里是正文。')
    drawing = OxmlElement('w:drawing')
    paragraph.runs[0]._r.append(drawing)

    assert _split_heading_by_punct(paragraph) is False
    assert len(paragraph._p.xpath('.//w:drawing')) == 1
