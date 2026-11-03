"""Deterministic custom paragraph-rule tests."""

from copy import deepcopy

import pytest

from scripts.formatter import PRESETS
from scripts.layout_rules import (
    ParagraphRuleError,
    compile_rule_set,
    ensure_paragraph_rule_defaults,
    match_custom_type,
    validate_paragraph_rules,
)


def _settings():
    settings = deepcopy(PRESETS["official"])
    settings["layout_mode"] = "generic"
    return ensure_paragraph_rule_defaults(settings)


def test_default_rules_match_chapters_articles_and_context_sequences():
    _types, rules = compile_rule_set(_settings())
    assert match_custom_type("第一章 总则", rules=rules) == "chapter"
    assert match_custom_type("第十二条 适用范围", rules=rules) == "article"
    assert match_custom_type(
        "1", next_text="这是一段语录正文。", next_next_text="（2026年会议讲话）", rules=rules,
    ) == "quote_number"
    assert match_custom_type(
        "这是一段语录正文。", previous_type="quote_number",
        next_text="（2026年会议讲话）", rules=rules,
    ) == "quote_text"


def test_invalid_or_overly_broad_rules_are_rejected():
    settings = _settings()
    with pytest.raises(ParagraphRuleError):
        validate_paragraph_rules(
            [{"id": "all", "name": "全部", "type_id": "chapter", "pattern": ".*"}],
            settings["paragraph_types"],
        )
    with pytest.raises(ParagraphRuleError):
        validate_paragraph_rules(
            [{"id": "nested", "name": "危险重复", "type_id": "chapter", "pattern": "(a+)+$"}],
            settings["paragraph_types"],
        )
