from pathlib import Path

from modern_ui import ModernButton


class _FakeVariable:
    def __init__(self):
        self.value = ""

    def set(self, value):
        self.value = str(value)

    def get(self):
        return self.value


class _FakeField:
    def __init__(self):
        self.override = None

    def set_display_override(self, text=None):
        self.override = text


class _FakeLogPanel:
    def __init__(self):
        self.messages = []

    def log(self, message, tag):
        self.messages.append((message, tag))


class _FakeResultPanel:
    def __init__(self):
        self.reset_count = 0

    def reset(self):
        self.reset_count += 1


def _make_selection_app():
    from docformat_gui import DocFormatApp

    app = object.__new__(DocFormatApp)
    app.input_files = []
    app.input_file = _FakeVariable()
    app.output_file = _FakeVariable()
    app.input_field = _FakeField()
    app.output_field = _FakeField()
    app.log_panel = _FakeLogPanel()
    app.result_panel = _FakeResultPanel()
    app.batch_selections = []
    app._update_batch_summary = lambda filenames: app.batch_selections.append(list(filenames))
    return app


def test_modern_button_accepts_legacy_text_updates():
    """Processing state must not forward ``text`` to the underlying Canvas."""
    source = Path("modern_ui.py").read_text(encoding="utf-8")
    configure_body = source.split("class ModernButton", 1)[1].split("class ChoiceChip", 1)[0]

    assert 'text = kwargs.pop("text", None)' in configure_body
    assert 'self.label.configure(text=str(text))' in configure_body


def test_single_file_selection_sets_a_ready_to_run_output_path(tmp_path):
    from docformat_gui import DocFormatApp

    source = tmp_path / "示例.docx"
    source.touch()
    app = _make_selection_app()

    DocFormatApp._add_files_to_list(app, [str(source)])

    assert app.input_files == [str(source)]
    assert app.input_file.get() == str(source)
    assert app.output_file.get() == str(tmp_path / "示例_processed.docx")
    assert app.input_field.override is None
    assert app.output_field.override is None
    assert app.batch_selections == [[str(source)]]


def test_multi_file_selection_keeps_all_paths_and_displays_a_batch_summary(tmp_path):
    from docformat_gui import DocFormatApp

    sources = [tmp_path / f"文档-{index}.docx" for index in range(1, 4)]
    for source in sources:
        source.touch()
    paths = [str(source) for source in sources]
    app = _make_selection_app()

    DocFormatApp._on_file_selected(app, paths)

    assert app.input_files == paths
    assert app.input_field.override == "已选择 3 个文件"
    assert app.output_file.get() == str(tmp_path)
    assert app.output_field.override == f"输出目录：{tmp_path}"
    assert app.batch_selections == [paths]
