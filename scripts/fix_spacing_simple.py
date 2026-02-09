#!/usr/bin/env python3
"""行距统一工具 - 简化版"""

import sys
from docx import Document
from docx.shared import Pt


def fix_line_spacing(input_path, output_path):
    """统一所有段落行距为29.45pt固定值"""
    print(f"Reading: {input_path}")
    doc = Document(input_path)

    # 修改所有段落的行距
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        pf = para.paragraph_format

        # 统一设置为29.45pt固定值
        pf.line_spacing = Pt(29.45)

    print("Fixed line spacing for all text paragraphs")
    doc.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python fix_spacing_simple.py input.docx output.docx")
        sys.exit(1)

    fix_line_spacing(sys.argv[1], sys.argv[2])
