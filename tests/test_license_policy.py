"""License migration safeguards for v1.8.8.3 and later releases."""

from pathlib import Path


PROJECT_ROOT = Path(__file__).parent.parent


def _read(relative_path):
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def test_polyform_license_has_project_notices_and_no_foreign_notice():
    license_text = _read("LICENSE")

    assert license_text.startswith("# PolyForm Noncommercial License 1.0.0")
    assert "<https://polyformproject.org/licenses/noncommercial/1.0.0>" in license_text
    assert "Required Notice: Copyright 2025-2026 KaguraNanaga." in license_text
    assert "released in version v1.8.8.3 and later" in license_text
    assert "Xueyou Luo" not in license_text
    assert "CAD files" not in license_text


def test_license_history_keeps_the_mit_boundary_explicit():
    history = _read("LICENSE-HISTORY.md")

    assert "v1.8.8.2" in history
    assert "MIT License" in history
    assert "v1.8.8.3" in history
    assert "PolyForm Noncommercial License 1.0.0" in history


def test_application_and_build_versions_match_the_current_release():
    gui_source = _read("docformat_gui.py")
    build_source = _read("build.py")

    assert "__version__ = '1.8.8.4'" in gui_source
    assert 'VERSION = "1.8.8.4"' in build_source
    assert "LICENSE_NAME = 'PolyForm Noncommercial License 1.0.0'" in gui_source
    assert "LICENSE_URL = 'https://polyformproject.org/licenses/noncommercial/1.0.0'" in gui_source


def test_readme_uses_a_brief_release_note_and_accurate_license_section():
    readme = _read("README.md")
    readme_en = _read("README_EN.md")

    assert "修复了一些小 Bug 和小问题，优化使用体验" in readme
    assert "使用许可调整" not in readme
    assert "License-PolyForm%20Noncommercial%201.0.0" in readme
    assert "License-PolyForm%20Noncommercial%201.0.0" in readme_en
    assert "仅限个人和非商业用途" in readme
    assert "free for personal and noncommercial purposes" in readme_en


def test_packaged_apps_include_the_license_and_history_files():
    build_source = _read("build.py")
    workflow = _read(".github/workflows/build.yml")

    assert "--add-data=LICENSE;." in build_source
    assert build_source.count("--add-data=LICENSE:.") == 2
    assert "--add-data=LICENSE-HISTORY.md;." in build_source
    assert build_source.count("--add-data=LICENSE-HISTORY.md:.") == 2
    assert workflow.count('--add-data "LICENSE:."') == 2
    assert workflow.count('--add-data "LICENSE-HISTORY.md:."') == 2
