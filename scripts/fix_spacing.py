#!/usr/bin/env python3
"""行距统一工具 - 修复段落行距不一致"""

import sys
from docx import Document
from docx.shared import Pt


def fix_line_spacing(input_path, output_path):
    """统一段落行距"""
    print(f"Reading: {input_path}")
    doc = Document(input_path)

    # 公文标准行距：29.45pt固定值（约 374015 twips）
    target_spacing = 374015  # 29.45pt in twips

    changes = 0

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        pf = para.paragraph_format

        # 跳过标题（通常行距可能不同）
        # 标题特征：短、居中、有特定编号
        if len(text) < 20 or pf.alignment is not None:
            continue

        # 修改行距为固定值29.45pt
        if pf.line_spacing != target_spacing:
            pf.line_spacing = Pt(29.45)  # 设置为29.45pt固定值
            changes += 1

    print(f"Fixed line spacing for {changes} paragraphs")
    doc.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python fix_spacing.py input.docx output.docx")
        sys.exit(1)

    fix_line_spacing(sys.argv[1], sys.argv[2])
