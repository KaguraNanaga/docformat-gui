"""Read-only visual confirmation after learning a style from a sample Word doc.

The pure ``build_style_cards`` function maps learned settings to display rows
so it can be unit-tested without Tk.  The dialog itself only previews and
names the style; parameter editing stays in the main settings dialog.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

from window_layout import apply_parent_relative_layout


BG = "#f6f2eb"
CARD = "#fffdf9"
TEXT = "#292621"
MUTED = "#777068"
LINE = "#ded5ca"
ACCENT = "#c74928"
ACCENT_DARK = "#a9381d"
WARN_BG = "#fff2cf"

SLOT_ORDER = ("title", "heading1", "heading2", "heading3", "heading4", "body")
SLOT_LABELS = {
    "title": "大标题",
    "heading1": "一级标题",
    "heading2": "二级标题",
    "heading3": "三级标题",
    "heading4": "四级标题",
    "body": "正文",
}
ALIGN_LABELS = {"left": "左对齐", "center": "居中", "right": "右对齐", "justify": "两端对齐"}
SAMPLE_TEXT = "公文格式排版示例 Abc123"


def _float(value, fallback=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(fallback)


def _num(value):
    """Compact number: 2.0 -> '2', 2.5 -> '2.5'."""
    return f"{float(value):g}"


def build_style_cards(settings, warnings=(), available_fonts=None):
    """Map learned settings to preview-card data.  Pure and Tk-free.

    Only slots actually present in *settings* are returned; recipient,
    signature, date, attachment and closing are skipped on purpose (the
    generic classifier never produces them).
    """
    settings = settings or {}
    fonts = None if available_fonts is None else {str(name).lower() for name in available_fonts}
    rows = []
    for key in SLOT_ORDER:
        fmt = settings.get(key)
        if not isinstance(fmt, dict):
            continue
        size = _float(fmt.get("size"), 12)
        indent = _float(fmt.get("indent"), 0)
        line_spacing = _float(fmt.get("line_spacing"), 0)
        font_cn = str(fmt.get("font_cn") or "默认字体")
        parts = [font_cn, f"{_num(size)} 磅"]
        if fmt.get("bold"):
            parts.append("加粗")
        parts.append(ALIGN_LABELS.get(fmt.get("align"), ALIGN_LABELS["left"]))
        if indent > 0.05 and size > 0:
            parts.append(f"首行缩进 {_num(round(indent / size, 1))} 字")
        if line_spacing > 0:
            parts.append(f"行距 {_num(line_spacing)} 磅")
        rows.append({
            "key": key,
            "label": SLOT_LABELS[key],
            "font_cn": font_cn,
            "size": size,
            "bold": bool(fmt.get("bold")),
            "summary": " · ".join(parts),
            "font_missing": fonts is not None and font_cn.lower() not in fonts,
        })

    page = settings.get("page") or {}
    width = _float(page.get("width_mm"), 210)
    height = _float(page.get("height_mm"), 297)
    if {round(width), round(height)} == {210, 297}:
        paper = "A4"
    else:
        paper = f"{_num(width)}×{_num(height)} 毫米"
    orientation = "横向" if page.get("orientation") == "landscape" else "竖向"
    margins = (
        f"上 {_num(_float(page.get('top'), 2.5))} · 下 {_num(_float(page.get('bottom'), 2.5))} · "
        f"左 {_num(_float(page.get('left'), 2.5))} · 右 {_num(_float(page.get('right'), 2.5))} 厘米"
    )
    page_row = {
        "summary": f"{paper} {orientation} · 页边距 {margins} · "
                   f"{'有页码' if settings.get('page_number') else '无页码'}",
    }

    # 常识校验：学习结果明显可疑时给用户一句人话提示
    notes = []
    by_key = {row["key"]: row for row in rows}
    title_row = by_key.get("title")
    body_row = by_key.get("body")
    if title_row and body_row and title_row["size"] <= body_row["size"]:
        notes.append("大标题不比正文显眼，可能是选错了段落；保存后请在参数里检查。")
    source = settings.get("layout_source") or {}
    table_heavy = (
        _float(source.get("paragraph_count"), 99) < 12
        and _float(source.get("table_count"), 0) > 0
    )
    cleaned_warnings = [str(item) for item in warnings if str(item).strip()]
    if table_heavy:
        notes.append("该文档大部分内容可能在表格中，学习结果可能不准，表格部分请排版后检查。")
        cleaned_warnings = [item for item in cleaned_warnings if "表格" not in item]
    return {
        "rows": rows,
        "page_row": page_row,
        "warnings": cleaned_warnings,
        "notes": notes,
    }


class SampleStyleDialog:
    """Modal preview sheet.  ``result`` is the chosen style name or None."""

    def __init__(self, parent, *, sample_name, settings, warnings=(), default_name="我的样式"):
        self.result = None
        self.default_name = default_name or "我的样式"
        self.top = tk.Toplevel(parent)
        self.top.title("已从样例学到样式")
        self.top.configure(bg=BG)
        self.top.transient(parent)
        apply_parent_relative_layout(
            self.top, parent,
            preferred_width=700, preferred_height=740,
            min_width=560, min_height=480,
        )
        self.top.protocol("WM_DELETE_WINDOW", self._cancel)
        self._configure_styles()

        available = set(tkfont.families(self.top))
        cards = build_style_cards(settings, warnings, available)
        self._build(sample_name, cards)
        self.top.grab_set()
        self.top.focus_force()

    def _configure_styles(self):
        style = ttk.Style(self.top)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Sample.Primary.TButton", background=ACCENT, foreground="white",
                        font=("Arial", 11, "bold"), padding=(16, 9), borderwidth=0)
        style.map("Sample.Primary.TButton",
                  background=[("active", ACCENT_DARK), ("pressed", ACCENT_DARK)])
        style.configure("Sample.Secondary.TButton", background=CARD, foreground=TEXT,
                        font=("Arial", 11), padding=(16, 9), borderwidth=1)
        style.map("Sample.Secondary.TButton", background=[("active", "#eee7de")])

    def _build(self, sample_name, cards):
        header = tk.Frame(self.top, bg=BG, padx=24, pady=18)
        header.pack(fill="x")
        tk.Label(header, text="已从样例学到样式", font=("Arial", 18, "bold"),
                 fg=TEXT, bg=BG).pack(anchor="w")
        tk.Label(
            header,
            text=f"样例：《{sample_name}》 · 只读取页面和字体格式，不复制正文内容。",
            font=("Arial", 10), fg=MUTED, bg=BG,
        ).pack(anchor="w", pady=(5, 0))

        name_row = tk.Frame(self.top, bg=BG, padx=24)
        name_row.pack(fill="x", pady=(0, 12))
        tk.Label(name_row, text="样式名称", font=("Arial", 11, "bold"), fg=TEXT,
                 bg=BG).pack(side="left")
        self.name_var = tk.StringVar(value=self.default_name)
        tk.Entry(name_row, textvariable=self.name_var, font=("Arial", 11),
                 relief="solid", bd=1, highlightthickness=0).pack(
            side="left", fill="x", expand=True, padx=(10, 0), ipady=5)

        # 底部按钮先占位：内容再高也不挤压按钮
        footer = tk.Frame(self.top, bg=BG, padx=24, pady=14)
        footer.pack(side="bottom", fill="x")
        tk.Label(footer, text="保存后就是当前排版样式；参数仍可在「自定义格式」里随时调整。",
                 font=("Arial", 10), fg=MUTED, bg=BG).pack(side="left")
        ttk.Button(footer, text="取消", command=self._cancel,
                   style="Sample.Secondary.TButton").pack(side="right")
        ttk.Button(footer, text="存为我的样式", command=self._confirm,
                   style="Sample.Primary.TButton").pack(side="right", padx=(0, 8))

        # 卡片区可滚动，窗口不够高时滚动而不是压缩
        canvas = tk.Canvas(self.top, bg=BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(self.top, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True, padx=(24, 24))
        body = tk.Frame(canvas, bg=BG)
        body_window = canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind(
            "<Configure>",
            lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(body_window, width=event.width),
        )
        canvas.bind(
            "<Enter>",
            lambda _event: canvas.bind_all(
                "<MouseWheel>",
                lambda event: canvas.yview_scroll(int(-event.delta), "units"),
            ),
        )
        canvas.bind("<Leave>", lambda _event: canvas.unbind_all("<MouseWheel>"))

        for row in cards["rows"]:
            self._style_card(body, row)
        self._page_card(body, cards["page_row"])
        for note in cards["notes"] + cards["warnings"]:
            tk.Label(body, text="注意：" + note, wraplength=560, justify="left",
                     font=("Arial", 10), fg=ACCENT_DARK, bg=WARN_BG,
                     padx=10, pady=8).pack(fill="x", pady=(0, 8))

    def _style_card(self, parent, row):
        card = tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightthickness=1,
                        padx=14, pady=10)
        card.pack(fill="x", pady=(0, 8))
        top_line = tk.Frame(card, bg=CARD)
        top_line.pack(fill="x")
        tk.Label(top_line, text=row["label"], font=("Arial", 11, "bold"), fg=TEXT,
                 bg=CARD).pack(side="left")
        summary = row["summary"]
        if row["font_missing"]:
            summary += " · 本机未安装该字体，预览为近似效果"
        tk.Label(top_line, text=summary, font=("Arial", 9), fg=MUTED,
                 bg=CARD).pack(side="right")
        render_family = "Arial" if row["font_missing"] else row["font_cn"]
        display_size = max(12, min(24, int(round(row["size"]))))
        tk.Label(
            card, text=SAMPLE_TEXT, bg=CARD, fg=TEXT, anchor="w", pady=4,
            font=(render_family, display_size, "bold" if row["bold"] else "normal"),
        ).pack(fill="x")

    def _page_card(self, parent, page_row):
        card = tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightthickness=1,
                        padx=14, pady=10)
        card.pack(fill="x", pady=(0, 8))
        top_line = tk.Frame(card, bg=CARD)
        top_line.pack(fill="x")
        tk.Label(top_line, text="页面设置", font=("Arial", 11, "bold"), fg=TEXT,
                 bg=CARD).pack(side="left")
        tk.Label(top_line, text=page_row["summary"], font=("Arial", 9), fg=MUTED,
                 bg=CARD).pack(side="right")

    def _confirm(self):
        self.result = self.name_var.get().strip() or self.default_name
        self.top.destroy()

    def _cancel(self):
        self.result = None
        self.top.destroy()
