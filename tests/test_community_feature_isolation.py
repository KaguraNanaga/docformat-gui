"""Prevent restricted feature surfaces from entering the community source tree."""

from pathlib import Path
import re


PROJECT_ROOT = Path(__file__).parent.parent
RUNTIME_FILES = (
    "docformat_gui.py",
    "modern_ui.py",
    "window_layout.py",
    "layout_rules_dialog.py",
    "sample_style_dialog.py",
    "scripts/formatter.py",
    "scripts/layout_analyzer.py",
    "scripts/layout_rules.py",
    "build.py",
)


def test_runtime_has_no_restricted_feature_modules_or_labels():
    forbidden_patterns = (
        r"\bA" + r"I\b",
        "人工智能",
        "红" + "头",
        "red" + "_header",
        "综" + "办",
        "飞" + "书",
        r"\bO" + r"CR\b",
        r"\bA" + r"gent\b",
        "内部" + "工具栏",
        "软件" + "更新",
        "update" + "_service",
        "license" + "_policy",
    )
    for relative_path in RUNTIME_FILES:
        source = (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")
        for pattern in forbidden_patterns:
            assert re.search(pattern, source, re.IGNORECASE) is None, (
                f"{relative_path} contains restricted pattern {pattern!r}"
            )


def test_runtime_imports_only_local_community_modules():
    source = "\n".join(
        (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")
        for relative_path in RUNTIME_FILES
    )
    restricted_module_fragments = (
        "layout_rule_" + "ai",
        "update_" + "service",
        "license_" + "policy",
        "red_" + "header",
        "internal_" + "tools",
        "ocr_" + "tools",
    )
    lowered = source.lower()
    for fragment in restricted_module_fragments:
        assert fragment not in lowered
