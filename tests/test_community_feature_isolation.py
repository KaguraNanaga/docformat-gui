"""Prevent restricted feature surfaces from entering the community source tree."""

import ast
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


def _without_about_copy(source):
    """Only static About copy may describe the separately offered products."""
    tree = ast.parse(source)

    class RemoveCopyLiterals(ast.NodeTransformer):
        def visit_Constant(self, node):
            return ast.Constant(value="") if isinstance(node.value, str) else node

    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == "AboutDialog":
            for child in ast.walk(node):
                if isinstance(child, ast.Assign) and any(
                    isinstance(target, ast.Name) and target.id == "about_text"
                    for target in child.targets
                ):
                    child.value = RemoveCopyLiterals().visit(child.value)
    return ast.dump(tree)


def test_about_copy_exception_preserves_executable_references():
    source = 'class AboutDialog:\n    about_text = f"AI {red_header.build()}"\n'
    scanned = _without_about_copy(source)
    assert "AI" not in scanned
    assert "red_header" in scanned
    assert "build" in scanned


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
        if relative_path == "docformat_gui.py":
            source = _without_about_copy(source)
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
