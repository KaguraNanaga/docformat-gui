"""OOXML paragraph settings for Chinese line-breaking rules in Word/WPS."""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


_PPR_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap",
    "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
)
_PPR_ORDER_INDEX = {name: index for index, name in enumerate(_PPR_ORDER)}


def _insert_paragraph_property(p_pr, element, local_name):
    """Insert a property in the order expected by strict OOXML consumers."""
    property_index = _PPR_ORDER_INDEX[local_name]
    for index, child in enumerate(p_pr):
        child_name = child.tag.rsplit("}", 1)[-1]
        if _PPR_ORDER_INDEX.get(child_name, -1) > property_index:
            p_pr.insert(index, element)
            return
    p_pr.append(element)


def _set_paragraph_boolean_property(paragraph, local_name, value):
    p_pr = paragraph._p.get_or_add_pPr()
    tag = qn(f"w:{local_name}")
    element = p_pr.find(tag)
    if element is None:
        element = OxmlElement(f"w:{local_name}")
        _insert_paragraph_property(p_pr, element, local_name)
    desired_value = "1" if value else "0"
    if element.get(qn("w:val")) == desired_value:
        return False
    element.set(qn("w:val"), desired_value)
    return True


def apply_chinese_line_break_rules_to_paragraph(paragraph):
    """Enable Chinese kinsoku and disable hanging punctuation for one paragraph."""
    if not paragraph.text.strip():
        return False
    changed = _set_paragraph_boolean_property(paragraph, "kinsoku", True)
    return _set_paragraph_boolean_property(paragraph, "overflowPunct", False) or changed


def _iter_container_paragraphs(container, seen):
    for paragraph in container.paragraphs:
        paragraph_id = id(paragraph._p)
        if paragraph_id not in seen:
            seen.add(paragraph_id)
            yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_container_paragraphs(cell, seen)


def apply_chinese_line_break_rules(document):
    """Apply Chinese punctuation line-break protection across document stories.

    Body, nested tables, independent headers and footers are included. Existing
    nodes are updated in place, so repeated processing is idempotent.
    """
    seen = set()
    changed_count = sum(
        apply_chinese_line_break_rules_to_paragraph(paragraph)
        for paragraph in _iter_container_paragraphs(document, seen)
    )
    story_names = (
        "header", "first_page_header", "even_page_header",
        "footer", "first_page_footer", "even_page_footer",
    )
    for section in document.sections:
        for story_name in story_names:
            story = getattr(section, story_name)
            if story.is_linked_to_previous:
                continue
            changed_count += sum(
                apply_chinese_line_break_rules_to_paragraph(paragraph)
                for paragraph in _iter_container_paragraphs(story, seen)
            )
    return changed_count
