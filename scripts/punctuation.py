#!/usr/bin/env python3
"""
标点符号修复 v5
- 修复引号处理bug：使用明确的Unicode转义序列
- 正确处理省略号和句号
"""

import re
import sys
from docx import Document

# 中文标点（使用Unicode转义确保正确）
LEFT_DOUBLE_QUOTE = '\u201c'   # " 左双引号
RIGHT_DOUBLE_QUOTE = '\u201d'  # " 右双引号
LEFT_SINGLE_QUOTE = '\u2018'   # ' 左单引号
RIGHT_SINGLE_QUOTE = '\u2019'  # ' 右单引号

# 基本替换映射
REPLACEMENTS = {
    "(": "（",
    ")": "）",
    ":": "：",
    ";": "；",
    "?": "？",
    "!": "！",
}

# 占位符前缀（使用不可能出现在正常文本中的字符序列）
_PLACEHOLDER_PREFIX = "\x02PROT"


def _protect_special_patterns(text):
    """提取并保护不应被替换的特殊模式，返回 (处理后文本, 保护列表)"""
    protected = []
    counter = [0]

    def _replace_with_placeholder(match):
        placeholder = f"{_PLACEHOLDER_PREFIX}{counter[0]}\x03"
        protected.append((placeholder, match.group()))
        counter[0] += 1
        return placeholder

    result = text

    # 保护 URL（http:// https:// ftp://）
    result = re.sub(r"(?:https?|ftp)://\S+", _replace_with_placeholder, result)

    # 保护邮箱地址
    result = re.sub(r"[\w.+-]+@[\w-]+\.[\w.-]+", _replace_with_placeholder, result)

    # 保护 Windows 文件路径 C:\ D:\
    result = re.sub(r"[A-Za-z]:\\", _replace_with_placeholder, result)

    # 保护标准编号 ISO 9001:2015 等（字母+数字+冒号+数字）
    result = re.sub(r"[A-Za-z]+[\s-]?\d+:\d{2,}", _replace_with_placeholder, result)

    # 保护时间格式 HH:MM 或 HH:MM:SS
    # 注意：不能用 \b，因为 Python3 中中文字符也属于 \w，
    # 导致"上午9:30"中 午 和 9 之间没有 \b 边界
    result = re.sub(r"(?<!\d)(\d{1,2}:\d{2}(?::\d{2})?)(?!\d)", _replace_with_placeholder, result)

    return result, protected


def _restore_protected(text, protected):
    """恢复被保护的内容"""
    result = text
    for placeholder, original in protected:
        result = result.replace(placeholder, original)
    return result


def has_chinese(text):
    """检查是否包含中文"""
    return bool(re.search(r"[\u4e00-\u9fff]", text))


def fix_text(text):
    """修复文本中的标点"""
    if not text:
        return text

    # 保护特殊模式
    result, protected = _protect_special_patterns(text)

    # ===== 第一步：处理省略号（必须在句号之前）=====
    result = re.sub(r"\.{2,}", "……", result)
    result = re.sub(r"。{2,}", "……", result)

    # ===== 第二步：处理破折号 =====
    result = re.sub(r"--+", "——", result)
    result = re.sub(r"—(?!—)", "——", result)

    # ===== 第三步：基本标点替换（只在有中文的文本中）=====
    if has_chinese(result):
        for en, cn in REPLACEMENTS.items():
            result = result.replace(en, cn)

    # ===== 第四步：逗号特殊处理 =====
    result = re.sub(r"([\u4e00-\u9fff]),", r"\1，", result)
    result = re.sub(r",([\u4e00-\u9fff])", r"，\1", result)

    # ===== 第五步：句号特殊处理 =====
    result = re.sub(r"([\u4e00-\u9fff])\.(\s|$)", r"\1。\2", result)

    # ===== 第六步：双引号处理 =====
    # 需要处理的双引号字符（使用Unicode确保正确）
    double_quote_chars = [
        '"',        # U+0022 ASCII直引号
        '\u201c',   # U+201C 左双引号 "
        '\u201d',   # U+201D 右双引号 "
        '\u201e',   # U+201E 双低引号 „
        '\u201f',   # U+201F 双高反引号 ‟
        '\u300c',   # U+300C 日文左引号 「
        '\u300d',   # U+300D 日文右引号 」
    ]

    # 统一替换为临时标记
    temp_result = result
    for q in double_quote_chars:
        temp_result = temp_result.replace(q, "\x00")

    # 配对处理
    if "\x00" in temp_result:
        chars = list(temp_result)
        quote_count = 0
        for i, c in enumerate(chars):
            if c == "\x00":
                # 偶数位置用左引号，奇数位置用右引号
                if quote_count % 2 == 0:
                    chars[i] = LEFT_DOUBLE_QUOTE  # "
                else:
                    chars[i] = RIGHT_DOUBLE_QUOTE  # "
                quote_count += 1
        result = "".join(chars)

    # ===== 第七步：单引号处理 =====
    single_quote_chars = [
        "'",        # U+0027 ASCII单引号
        '\u2018',   # U+2018 左单引号 '
        '\u2019',   # U+2019 右单引号 '
        '\u201a',   # U+201A 单低引号 ‚
        '\u201b',   # U+201B 单高反引号 ‛
    ]

    temp_result = result
    for q in single_quote_chars:
        temp_result = temp_result.replace(q, "\x01")

    if "\x01" in temp_result:
        chars = list(temp_result)
        quote_count = 0
        for i, c in enumerate(chars):
            if c == "\x01":
                # 偶数位置用左引号，奇数位置用右引号
                if quote_count % 2 == 0:
                    chars[i] = LEFT_SINGLE_QUOTE  # '
                else:
                    chars[i] = RIGHT_SINGLE_QUOTE  # '
                quote_count += 1
        result = "".join(chars)

    # 恢复被保护的内容
    result = _restore_protected(result, protected)
    return result


def _fix_simple_punctuation(text):
    """只做不涉及配对的简单标点替换，保留引号不动"""
    if not text:
        return text

    # 保护特殊模式
    result, protected = _protect_special_patterns(text)

    # 省略号（必须在句号之前）
    result = re.sub(r"\.{2,}", "……", result)
    result = re.sub(r"。{2,}", "……", result)

    # 破折号
    result = re.sub(r"--+", "——", result)
    result = re.sub(r"—(?!—)", "——", result)

    # 基本替换（只在有中文时）
    if has_chinese(result):
        for en, cn in REPLACEMENTS.items():
            result = result.replace(en, cn)

    # 逗号
    result = re.sub(r"([\u4e00-\u9fff]),", r"\1，", result)
    result = re.sub(r",([\u4e00-\u9fff])", r"，\1", result)

    # 句号
    result = re.sub(r"([\u4e00-\u9fff])\.(\s|$)", r"\1。\2", result)

    # 恢复被保护的内容
    result = _restore_protected(result, protected)
    return result


def _fix_quotes_whole_text(text):
    """对完整文本做引号配对替换"""
    result = text

    # 双引号
    double_quote_chars = ['"', "\u201c", "\u201d", "\u201e", "\u201f", "\u300c", "\u300d"]
    temp = result
    for q in double_quote_chars:
        temp = temp.replace(q, "\x00")

    if "\x00" in temp:
        chars = list(temp)
        quote_idx = 0
        for i, c in enumerate(chars):
            if c == "\x00":
                chars[i] = LEFT_DOUBLE_QUOTE if quote_idx % 2 == 0 else RIGHT_DOUBLE_QUOTE
                quote_idx += 1
        result = "".join(chars)

    # 单引号
    single_quote_chars = ["'", "\u2018", "\u2019", "\u201a", "\u201b"]
    temp = result
    for q in single_quote_chars:
        temp = temp.replace(q, "\x01")

    if "\x01" in temp:
        chars = list(temp)
        quote_idx = 0
        for i, c in enumerate(chars):
            if c == "\x01":
                chars[i] = LEFT_SINGLE_QUOTE if quote_idx % 2 == 0 else RIGHT_SINGLE_QUOTE
                quote_idx += 1
        result = "".join(chars)

    return result


def _redistribute_text_to_runs(runs, new_full_text):
    """将新的完整文本按原 run 的长度边界重新分配回各 run，保留格式"""
    run_lengths = [len(run.text) for run in runs]

    # 如果长度一致（只是字符替换，没有增删），直接按原长度切分
    total_original = sum(run_lengths)
    if len(new_full_text) == total_original:
        pos = 0
        for i, run in enumerate(runs):
            run.text = new_full_text[pos:pos + run_lengths[i]]
            pos += run_lengths[i]
    else:
        # 长度变了（极少情况），回退到旧方案：全塞第一个 run
        runs[0].text = new_full_text
        for run in runs[1:]:
            run.text = ""


def process_paragraph(para):
    """处理段落 - 简单替换按 run 做（保留格式），引号配对跨 run 做"""
    full_text = para.text
    if not full_text.strip():
        return False

    runs = para.runs
    if not runs:
        return False

    changed = False

    # 第一步：逐 run 做简单标点替换（保留所有格式）
    for run in runs:
        original = run.text
        fixed = _fix_simple_punctuation(original)
        if fixed != original:
            run.text = fixed
            changed = True

    # 第二步：引号配对需要整段处理
    full_after_simple = para.text
    full_after_quotes = _fix_quotes_whole_text(full_after_simple)

    if full_after_quotes != full_after_simple:
        _redistribute_text_to_runs(runs, full_after_quotes)
        changed = True

    return changed


def process_document(input_path, output_path):
    """处理文档"""
    print(f"Reading: {input_path}")
    doc = Document(input_path)

    changes = 0

    # 处理段落
    for i, para in enumerate(doc.paragraphs):
        if process_paragraph(para):
            changes += 1
            preview = para.text[:50] + "..." if len(para.text) > 50 else para.text
            print(f"  Para {i + 1}: {preview}")

    # 处理表格
    table_changes = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if process_paragraph(para):
                        table_changes += 1

    if table_changes > 0:
        print(f"  Tables: {table_changes} cells fixed")

    print()
    print(f"Total: {changes} paragraphs + {table_changes} table cells fixed")
    doc.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python punctuation.py input.docx output.docx")
        sys.exit(1)

    process_document(sys.argv[1], sys.argv[2])

# 测试用例（验证 fix_text）
# assert fix_text('会议时间:上午9:30至下午14:30,请准时参加.') == '会议时间：上午9:30至下午14:30，请准时参加。'
# assert fix_text('请于14:30前将材料发送至 report@gov.cn,逾期不候.') == '请于14:30前将材料发送至 report@gov.cn，逾期不候。'
# assert fix_text('详情请访问 https://www.example.com:8080/path 了解.') == '详情请访问 https://www.example.com:8080/path 了解。'
# assert fix_text('参照 ISO 9001:2015 执行.') == '参照 ISO 9001:2015 执行。'
# assert fix_text('每日9:00前完成巡检.') == '每日9:00前完成巡检。'
# assert '还需要改进' in fix_text('他说"这个方案"不错，但是"还需要改进')  # 奇数引号照常替换
