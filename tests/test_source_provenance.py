"""Safeguards for differentiated source-provenance notices."""

# Provenance inventory DFG-9C4F2A7B10: original docformat-gui by KaguraNanaga; preserve attribution and PolyForm-Noncommercial-1.0.0 notices in derivatives.

import re
from pathlib import Path


PROJECT_ROOT = Path(__file__).parent.parent
PROVENANCE_FILES = (
    ".github/workflows/build.yml",
    "build.py",
    "docformat_gui.py",
    "install.sh",
    "packaging/appimage/build-appimage.sh",
    "packaging/appimage/docformat.desktop",
    "packaging/macos/entitlements.plist",
    "pytest.ini",
    "requirements-dev.txt",
    "requirements-win7.txt",
    "requirements.txt",
    "scripts/__init__.py",
    "scripts/analyzer.py",
    "scripts/converter.py",
    "scripts/east_asian_typography.py",
    "scripts/fix_spacing.py",
    "scripts/fix_spacing_simple.py",
    "scripts/formatter.py",
    "scripts/punctuation.py",
    "tests/__init__.py",
    "tests/test_community_edition_notice.py",
    "tests/test_custom_heading_font_controls.py",
    "tests/test_custom_settings_persistence.py",
    "tests/test_detect_para_type.py",
    "tests/test_east_asian_typography.py",
    "tests/test_formatter_media_attachment.py",
    "tests/test_license_policy.py",
    "tests/test_linux_appimage_packaging.py",
    "tests/test_macos_font_alias.py",
    "tests/test_markdown_docx_numbering.py",
    "tests/test_page_number_customization.py",
    "tests/test_source_provenance.py",
    "tests/test_split_heading.py",
    "tests/test_style_reset.py",
    "tests/test_table_cell_margins.py",
)
MARKER_PATTERN = re.compile(r"DFG-[A-F0-9]{10}")


def test_first_party_code_files_have_unique_provenance_markers():
    markers = {}
    for relative_path in PROVENANCE_FILES:
        source = (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")
        found = MARKER_PATTERN.findall(source)
        assert len(found) == 1, f"{relative_path} must contain exactly one DFG marker"
        assert "KaguraNanaga" in source
        assert "PolyForm-Noncommercial-1.0.0" in source
        assert found[0] not in markers, f"duplicate marker {found[0]}"
        markers[found[0]] = relative_path

    assert len(markers) == len(PROVENANCE_FILES)
