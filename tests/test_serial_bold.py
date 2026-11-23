"""Exercise the sequence-word bold options through DOCX formatting."""

from docx import Document
import pytest

from scripts.formatter import format_document


@pytest.mark.parametrize(('options', 'bold_text'), [
    ({}, '一是'),
    ({'bold_serial_sentence': True}, '一是提高3.5个百分点；二是跟踪进度。三要按期完成'),
    ({'bold_serial': False, 'bold_serial_sentence': True}, ''),
])
def test_serial_bold_scope_preserves_other_sentences(tmp_path, options, bold_text):
    source = tmp_path / 'source.docx'
    output = tmp_path / 'output.docx'
    document = Document()
    document.add_paragraph('关于工作安排的通知')
    document.add_paragraph('一、工作安排')
    body = document.add_paragraph()
    for text in ('一', '是提高3.5个百分点；二', '是跟踪进度。另行通知。', '三要按期完成'):
        body.add_run(text)
    original_text = body.text
    document.save(source)

    format_document(str(source), str(output), custom_settings={'page_number': False, **options})

    body = Document(output).paragraphs[-1]
    assert body.text == original_text
    assert ''.join(run.text for run in body.runs if run.bold) == bold_text
