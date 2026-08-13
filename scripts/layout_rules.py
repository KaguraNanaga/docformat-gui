"""Deterministic paragraph-type rules for general-purpose documents.

Only this local matcher decides how a paragraph is formatted. Rules are
deliberately small and auditable.
"""

from __future__ import annotations

from copy import deepcopy
import re


PARAGRAPH_RULE_VERSION = 1
CUSTOM_TYPE_PREFIX = "custom:"
MAX_PATTERN_LENGTH = 240
MAX_REGEX_AST_DEPTH = 40
MAX_REGEX_AST_NODES = 500
MAX_BOUNDED_REPEAT = 200
BASE_PARAGRAPH_TYPES = {
    "title", "heading1", "heading2", "heading3", "heading4", "body",
    "recipient", "signature", "date", "attachment", "closing", "source_note",
}
INLINE_FORMAT_KEYS = {
    "font_cn", "font_en", "size", "bold", "italic", "underline", "color",
}


class ParagraphRuleError(ValueError):
    """Raised when a paragraph type or rule is unsafe or malformed."""


def _safe_id(value, fallback="type_custom"):
    value = re.sub(r"[^a-z0-9_]+", "_", str(value or "").strip().lower())
    value = value.strip("_")
    if not value:
        value = re.sub(r"[^a-z0-9_]+", "_", str(fallback or "").strip().lower()).strip("_")
    if not value:
        value = "type_custom"
    if not value[0].isalpha():
        value = f"type_{value}"
    return value[:40]


def validate_manual_id(value, label="ID"):
    """Validate an ID typed by a user without silently rewriting it."""
    value = str(value or "").strip()
    if not re.fullmatch(r"[a-z][a-z0-9_]{0,39}", value):
        raise ParagraphRuleError(
            f"{label} 只能使用小写字母、数字、下划线，并且必须以字母开头。"
        )
    return value


def custom_para_type(type_id):
    return f"{CUSTOM_TYPE_PREFIX}{_safe_id(type_id)}"


def plain_type_id(para_type):
    value = str(para_type or "")
    return value[len(CUSTOM_TYPE_PREFIX):] if value.startswith(CUSTOM_TYPE_PREFIX) else value


def _type_format(settings, key, fallback="body"):
    value = settings.get(key)
    if not isinstance(value, dict):
        value = settings.get(fallback, {})
    return deepcopy(value if isinstance(value, dict) else {})


def built_in_paragraph_types(settings):
    body = _type_format(settings, "body")
    chapter = _type_format(settings, "heading1")
    article = _type_format(settings, "heading2")
    quote_number = deepcopy(body)
    quote_number.update({"bold": True, "align": "left", "indent": 0})
    quote_text = deepcopy(body)
    quote_source = deepcopy(body)
    quote_source.update({"align": "right", "indent": 0})
    return [
        {"id": "chapter", "name": "章标题", "format": chapter},
        {
            "id": "article", "name": "条款", "format": article,
            "prefix": {
                "pattern": rf"^第[{_CN_NUM}\d]+条",
                "format": {"bold": True},
            },
        },
        {"id": "quote_number", "name": "语录序号", "format": quote_number},
        {"id": "quote_text", "name": "语录正文", "format": quote_text},
        {"id": "quote_source", "name": "语录出处", "format": quote_source},
    ]


_CN_NUM = "一二三四五六七八九十百千零○〇"
_SOURCE_PATTERN = (
    r"^[\(（].{0,120}(?:\d{4}年|[" + _CN_NUM + r"]+年|"
    r"讲话|考察|会议|主持|座谈|发表|指出|强调).{0,120}[\)）]$"
)


def built_in_paragraph_rules():
    return [
        {
            "id": "rule_chapter", "name": "第几章", "type_id": "chapter",
            "pattern": rf"^第[{_CN_NUM}\d]+章[^\r\n]*$",
            "priority": 300, "enabled": True,
        },
        {
            "id": "rule_article", "name": "第几条", "type_id": "article",
            "pattern": rf"^第[{_CN_NUM}\d]+条[^\r\n]*$",
            "priority": 290, "enabled": True,
        },
        {
            "id": "rule_quote_number", "name": "语录独立序号", "type_id": "quote_number",
            "pattern": r"^\d{1,3}[\s　]*[\.．、]?$",
            "next_pattern": r"^\S.{3,}$", "next_next_pattern": _SOURCE_PATTERN,
            "priority": 280, "enabled": True,
        },
        {
            "id": "rule_quote_text", "name": "序号后的语录正文", "type_id": "quote_text",
            "pattern": r"^\S.{3,}$", "previous_type": "quote_number",
            "next_pattern": _SOURCE_PATTERN,
            "priority": 270, "enabled": True,
        },
        {
            "id": "rule_quote_source", "name": "语录出处", "type_id": "quote_source",
            "pattern": _SOURCE_PATTERN, "previous_type": "quote_text",
            "priority": 260, "enabled": True,
        },
    ]


def ensure_paragraph_rule_defaults(settings):
    """Seed one conservative rule pack exactly once for generic presets."""
    if not isinstance(settings, dict):
        return settings
    if str(settings.get("layout_mode") or "official").lower() != "generic":
        return settings
    if not settings.get("paragraph_rule_version"):
        settings.setdefault("paragraph_types", built_in_paragraph_types(settings))
        settings.setdefault("paragraph_rules", built_in_paragraph_rules())
        settings["paragraph_rule_version"] = PARAGRAPH_RULE_VERSION
    else:
        settings.setdefault("paragraph_types", [])
        settings.setdefault("paragraph_rules", [])
    return settings


def validate_paragraph_types(types):
    validated = []
    seen = set()
    for index, raw in enumerate(types or []):
        if not isinstance(raw, dict):
            continue
        type_id = _safe_id(raw.get("id"), f"type_{index + 1}")
        if type_id in BASE_PARAGRAPH_TYPES:
            raise ParagraphRuleError(f"段落类型 ID 不能使用系统保留名：{type_id}")
        if type_id in seen:
            raise ParagraphRuleError(f"段落类型 ID 重复：{type_id}")
        name = str(raw.get("name") or type_id).strip()[:40]
        fmt = raw.get("format")
        if not isinstance(fmt, dict):
            raise ParagraphRuleError(f"段落类型「{name}」缺少格式参数。")
        seen.add(type_id)
        item = {"id": type_id, "name": name, "format": deepcopy(fmt)}
        prefix = raw.get("prefix")
        if isinstance(prefix, dict):
            prefix_pattern = _validated_pattern(prefix.get("pattern"), name, optional=True)
            prefix_format = prefix.get("format")
            if prefix_pattern and isinstance(prefix_format, dict):
                clean_format = {
                    key: deepcopy(value)
                    for key, value in prefix_format.items()
                    if key in INLINE_FORMAT_KEYS
                }
                if clean_format:
                    item["prefix"] = {"pattern": prefix_pattern, "format": clean_format}
        validated.append(item)
    return validated


def _regex_parser():
    parser = getattr(re, "_parser", None)
    if parser is not None:
        return parser
    try:
        import sre_parse as parser  # type: ignore[deprecated]
    except (ImportError, AttributeError) as exc:
        raise ParagraphRuleError("当前运行环境无法执行正则安全校验。") from exc
    return parser


def _validate_regex_structure(pattern, label):
    """Reject backtracking-prone constructs using the parsed regex tree."""
    parser = _regex_parser()
    try:
        nodes = parser.parse(pattern, 0)
    except (re.error, ValueError, OverflowError, RuntimeError) as exc:
        raise ParagraphRuleError(f"规则「{label}」的表达式无效：{exc}") from exc

    repeat_ops = tuple(
        op for op in (
            getattr(parser, "MAX_REPEAT", None),
            getattr(parser, "MIN_REPEAT", None),
            getattr(parser, "POSSESSIVE_REPEAT", None),
        ) if op is not None
    )
    branch_op = getattr(parser, "BRANCH", None)
    groupref_ops = {
        op for op in (
            getattr(parser, "GROUPREF", None),
            getattr(parser, "GROUPREF_EXISTS", None),
        ) if op is not None
    }
    subpattern_op = getattr(parser, "SUBPATTERN", None)
    assertion_ops = {
        op for op in (
            getattr(parser, "ASSERT", None),
            getattr(parser, "ASSERT_NOT", None),
        ) if op is not None
    }
    atomic_group_op = getattr(parser, "ATOMIC_GROUP", None)
    max_repeat = getattr(parser, "MAXREPEAT", None)
    visited = 0

    def walk(current_nodes, *, in_repeat=False, depth=0):
        nonlocal visited
        if depth > MAX_REGEX_AST_DEPTH:
            raise ParagraphRuleError(f"规则「{label}」的表达式嵌套层级过深。")
        for op, argument in current_nodes:
            visited += 1
            if visited > MAX_REGEX_AST_NODES:
                raise ParagraphRuleError(f"规则「{label}」的表达式结构过于复杂。")
            if op in groupref_ops:
                raise ParagraphRuleError(f"规则「{label}」不能使用反向引用。")
            if op in repeat_ops:
                _minimum, maximum, child_nodes = argument
                if in_repeat:
                    raise ParagraphRuleError(
                        f"规则「{label}」包含容易造成卡顿的嵌套重复。"
                    )
                if max_repeat is not None and maximum != max_repeat and maximum > MAX_BOUNDED_REPEAT:
                    raise ParagraphRuleError(
                        f"规则「{label}」的重复次数不能超过 {MAX_BOUNDED_REPEAT}。"
                    )
                walk(child_nodes, in_repeat=True, depth=depth + 1)
            elif branch_op is not None and op == branch_op:
                if in_repeat:
                    raise ParagraphRuleError(
                        f"规则「{label}」不能对包含交替分支的表达式执行重复。"
                    )
                for branch in argument[1]:
                    walk(branch, in_repeat=in_repeat, depth=depth + 1)
            elif subpattern_op is not None and op == subpattern_op:
                walk(argument[3], in_repeat=in_repeat, depth=depth + 1)
            elif op in assertion_ops:
                walk(argument[1], in_repeat=in_repeat, depth=depth + 1)
            elif atomic_group_op is not None and op == atomic_group_op:
                walk(argument, in_repeat=in_repeat, depth=depth + 1)

    walk(nodes)


def _validated_pattern(value, label, *, optional=False):
    pattern = str(value or "").strip()
    if not pattern and optional:
        return ""
    if not pattern:
        raise ParagraphRuleError(f"规则「{label}」缺少匹配表达式。")
    if len(pattern) > MAX_PATTERN_LENGTH:
        raise ParagraphRuleError(f"规则「{label}」的表达式过长。")
    _validate_regex_structure(pattern, label)
    return pattern


def validate_paragraph_rules(rules, types):
    type_ids = {item["id"] for item in validate_paragraph_types(types)}
    validated = []
    seen = set()
    for index, raw in enumerate(rules or []):
        if not isinstance(raw, dict):
            continue
        name = str(raw.get("name") or f"规则{index + 1}").strip()[:60]
        rule_id = _safe_id(raw.get("id"), f"rule_{index + 1}")
        if rule_id in seen:
            raise ParagraphRuleError(f"规则 ID 重复：{rule_id}")
        type_id = _safe_id(raw.get("type_id"))
        if type_id not in type_ids:
            raise ParagraphRuleError(f"规则「{name}」引用了不存在的段落类型。")
        previous_type = str(raw.get("previous_type") or "").strip()
        previous_type = _safe_id(previous_type) if previous_type else ""
        if previous_type and previous_type not in type_ids | BASE_PARAGRAPH_TYPES:
            raise ParagraphRuleError(f"规则「{name}」的前置类型不存在。")
        pattern = _validated_pattern(raw.get("pattern"), name)
        next_pattern = _validated_pattern(raw.get("next_pattern"), name, optional=True)
        next_next_pattern = _validated_pattern(raw.get("next_next_pattern"), name, optional=True)
        if pattern.replace(" ", "") in {".*", "^.*$", ".+", "^.+$", "^\\S.*$"} and not (
            previous_type or next_pattern or next_next_pattern
        ):
            raise ParagraphRuleError(f"规则「{name}」过于宽泛，会匹配几乎所有段落。")
        try:
            priority = max(-1000, min(1000, int(raw.get("priority", 100))))
        except (TypeError, ValueError):
            priority = 100
        seen.add(rule_id)
        validated.append({
            "id": rule_id,
            "name": name,
            "type_id": type_id,
            "pattern": pattern,
            "previous_type": previous_type,
            "next_pattern": next_pattern,
            "next_next_pattern": next_next_pattern,
            "priority": priority,
            "enabled": bool(raw.get("enabled", True)),
        })
    return sorted(validated, key=lambda item: item["priority"], reverse=True)


def compile_rule_set(settings):
    """Return validated type mapping and priority-sorted rules."""
    ensure_paragraph_rule_defaults(settings)
    types = validate_paragraph_types(settings.get("paragraph_types", []))
    rules = validate_paragraph_rules(settings.get("paragraph_rules", []), types)
    return {item["id"]: item for item in types}, rules


def match_custom_type(text, *, previous_type="", next_text="", next_next_text="", rules=None):
    """Return a custom type id when the first validated rule matches."""
    text = str(text or "").strip()
    previous_type = plain_type_id(previous_type)
    for rule in rules or []:
        if not rule.get("enabled", True):
            continue
        if rule.get("previous_type") and rule["previous_type"] != previous_type:
            continue
        if not re.search(rule["pattern"], text):
            continue
        if rule.get("next_pattern") and not re.search(rule["next_pattern"], str(next_text or "").strip()):
            continue
        if rule.get("next_next_pattern") and not re.search(
            rule["next_next_pattern"], str(next_next_text or "").strip()
        ):
            continue
        return rule["type_id"]
    return ""


def merge_rule_pack(settings, rule_pack):
    """Merge an explicitly selected local rule pack without replacing other rules."""
    settings = deepcopy(settings if isinstance(settings, dict) else {})
    ensure_paragraph_rule_defaults(settings)
    existing_types = {item.get("id"): item for item in settings.get("paragraph_types", []) if isinstance(item, dict)}
    for item in rule_pack.get("types", []) if isinstance(rule_pack, dict) else []:
        existing_types[item.get("id")] = deepcopy(item)
    existing_rules = {item.get("id"): item for item in settings.get("paragraph_rules", []) if isinstance(item, dict)}
    for item in rule_pack.get("rules", []) if isinstance(rule_pack, dict) else []:
        existing_rules[item.get("id")] = deepcopy(item)
    settings["paragraph_types"] = validate_paragraph_types(existing_types.values())
    settings["paragraph_rules"] = validate_paragraph_rules(existing_rules.values(), settings["paragraph_types"])
    settings["paragraph_rule_version"] = PARAGRAPH_RULE_VERSION
    return settings
