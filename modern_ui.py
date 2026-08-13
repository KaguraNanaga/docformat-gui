"""公文格式处理工具的统一视觉语言（视觉系统 v3）。

设计原则：
- 浅暖纸张页面、暖白控件、hairline 细描边和克制圆角；
- 按钮三级：primary（品牌朱砂实心）、secondary（暖白描边，默认）、tonal（浅色底 + 强调色文字）；
- 按钮宽度用 ``tkinter.font.Font.measure`` 实测，杜绝中英混排时的截断和拥挤；
- 单选卡片（ChoiceChip）常驻左侧指示圈并保持左对齐，选中态不再让文字跳动。

所有颜色、间距从调用方传入的 theme 读取，本模块不保存配色副本。
"""

from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

_UI_DENSITY = 1.0


def set_ui_density(density):
    """由宿主程序在读取显示器 DPI 后调用，让共享控件内部尺寸跟随缩放。"""
    global _UI_DENSITY
    try:
        _UI_DENSITY = max(1.0, min(2.5, float(density)))
    except (TypeError, ValueError):
        _UI_DENSITY = 1.0


def px(value):
    return max(1, int(round(float(value) * _UI_DENSITY)))


def rounded_rect_points(x1, y1, x2, y2, radius):
    radius = max(0, min(radius, (x2 - x1) / 2, (y2 - y1) / 2))
    return [
        x1 + radius, y1,
        x2 - radius, y1,
        x2, y1,
        x2, y1 + radius,
        x2, y2 - radius,
        x2, y2,
        x2 - radius, y2,
        x1 + radius, y2,
        x1, y2,
        x1, y2 - radius,
        x1, y1 + radius,
        x1, y1,
    ]


def measure_text(widget, text, font_spec):
    """用真实字体度量文字宽度；tkinter 不可用时退回粗略估算。"""
    try:
        font = tkfont.Font(root=widget, font=font_spec)
        return font.measure(str(text))
    except tk.TclError:
        size = font_spec[1] if isinstance(font_spec, (tuple, list)) and len(font_spec) > 1 else 10
        return int(len(str(text)) * size * 1.7)


def _mixed_font_runs(text):
    """Split text so ASCII uses the Latin font and CJK uses the UI font."""
    runs = []
    for char in str(text):
        kind = "latin" if ord(char) < 128 else "cjk"
        if runs and runs[-1][0] == kind:
            runs[-1] = (kind, runs[-1][1] + char)
        else:
            runs.append((kind, char))
    return runs


def _font_family_variant(font_spec, family):
    if isinstance(font_spec, (tuple, list)):
        return (family, *font_spec[1:])
    return (family, 10)


def measure_mixed_text(widget, text, font_spec, latin_font_family):
    total = 0
    for kind, run in _mixed_font_runs(text):
        run_font = (
            _font_family_variant(font_spec, latin_font_family)
            if kind == "latin" else font_spec
        )
        total += measure_text(widget, run, run_font)
    return total


class MixedFontLabel(tk.Frame):
    """One-line label with a dedicated Latin family and a CJK UI family."""

    def __init__(self, parent, text, font_spec, latin_font_family, bg, fg, cursor="arrow"):
        super().__init__(parent, bg=bg, cursor=cursor, bd=0, highlightthickness=0)
        self._text = str(text)
        self._font_spec = font_spec
        self._latin_font_family = latin_font_family
        self._bg = bg
        self._fg = fg
        self._cursor = cursor
        self._parts = []
        self._part_bindings = []
        self._render_parts()

    def _render_parts(self):
        for part in self._parts:
            part.destroy()
        self._parts = []
        for kind, run in _mixed_font_runs(self._text):
            run_font = (
                _font_family_variant(self._font_spec, self._latin_font_family)
                if kind == "latin" else self._font_spec
            )
            part = tk.Label(
                self,
                text=run,
                font=run_font,
                bg=self._bg,
                fg=self._fg,
                cursor=self._cursor,
                padx=0,
                pady=0,
                bd=0,
            )
            part.pack(side="left")
            for sequence, callback in self._part_bindings:
                part.bind(sequence, callback)
            self._parts.append(part)

    def set_style(self, bg=None, fg=None, cursor=None):
        options = {}
        if bg is not None:
            options["bg"] = bg
            self._bg = bg
            super().configure(bg=bg)
        if fg is not None:
            options["fg"] = fg
            self._fg = fg
        if cursor is not None:
            options["cursor"] = cursor
            self._cursor = cursor
            super().configure(cursor=cursor)
        if options:
            for part in self._parts:
                part.configure(**options)

    def bind_parts(self, sequence, callback):
        self.bind(sequence, callback)
        self._part_bindings.append((sequence, callback))
        for part in self._parts:
            part.bind(sequence, callback)

    def configure(self, cnf=None, **kwargs):
        options = dict(cnf or {})
        options.update(kwargs)
        text = options.pop("text", None)
        font_spec = options.pop("font", None)
        bg = options.pop("bg", options.pop("background", None))
        fg = options.pop("fg", options.pop("foreground", None))
        cursor = options.pop("cursor", None)
        if options:
            super().configure(options)
        rerender = False
        if text is not None and str(text) != self._text:
            self._text = str(text)
            rerender = True
        if font_spec is not None and font_spec != self._font_spec:
            self._font_spec = font_spec
            rerender = True
        self.set_style(bg=bg, fg=fg, cursor=cursor)
        if rerender:
            self._render_parts()

    config = configure

    def cget(self, key):
        if key == "text":
            return self._text
        if key == "font":
            return self._font_spec
        return super().cget(key)


def _hairline(theme):
    return getattr(theme, "HAIRLINE", theme.BORDER)


def _control_border(theme):
    return getattr(theme, "BORDER_STRONG", theme.BORDER)


def _accent_light(theme, accent):
    """把强调色映射到对应的浅色底（选中态 / tonal 按钮用）。"""
    mapping = {
        getattr(theme, "PRIMARY", None): getattr(theme, "PRIMARY_LIGHT", None),
        getattr(theme, "ACCENT", None): getattr(theme, "ACCENT_LIGHT", None),
    }
    return mapping.get(accent) or getattr(theme, "PRIMARY_LIGHT", theme.INPUT_BG)


def _accent_light_hover(theme, accent):
    if accent == getattr(theme, "ACCENT", None):
        return getattr(theme, "ACCENT_LIGHT_HOVER", _accent_light(theme, accent))
    return getattr(theme, "PRIMARY_LIGHT_HOVER", _accent_light(theme, accent))


def divider(parent, theme, vertical=False):
    """1 像素 hairline 分隔线。"""
    if vertical:
        return tk.Frame(parent, bg=_hairline(theme), width=1)
    return tk.Frame(parent, bg=_hairline(theme), height=1)


def configure_ttk_styles(root, theme, get_font):
    style = ttk.Style(root)
    try:
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except tk.TclError:
        pass

    # 下拉框：白底、细描边、聚焦变蓝，与 ModernEntry 一致。
    style.configure(
        "TCombobox",
        fieldbackground=theme.CARD,
        background=theme.CARD,
        foreground=theme.TEXT,
        bordercolor=_control_border(theme),
        lightcolor=_control_border(theme),
        darkcolor=_control_border(theme),
        arrowcolor=theme.TEXT_SECONDARY,
        padding=(10, 7),
        relief="flat",
        font=get_font(10),
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", theme.CARD), ("focus", theme.CARD), ("disabled", theme.INPUT_BG)],
        bordercolor=[("focus", theme.PRIMARY), ("readonly", _control_border(theme))],
        lightcolor=[("focus", theme.PRIMARY)],
        darkcolor=[("focus", theme.PRIMARY)],
        arrowcolor=[("disabled", theme.TEXT_MUTED), ("readonly", theme.TEXT_SECONDARY)],
    )
    try:
        root.option_add("*TCombobox*Listbox.background", theme.CARD)
        root.option_add("*TCombobox*Listbox.foreground", theme.TEXT)
        root.option_add("*TCombobox*Listbox.selectBackground", theme.PRIMARY_LIGHT)
        root.option_add("*TCombobox*Listbox.selectForeground", theme.TEXT)
        root.option_add("*TCombobox*Listbox.borderWidth", 0)
        root.option_add("*Scrollbar.width", px(18))
    except tk.TclError:
        pass

    # 标签页：去掉凸起灰块，改为安静的文字标签，选中项用主色标识。
    style.configure(
        "TNotebook",
        background=theme.BG,
        borderwidth=0,
        tabmargins=(0, 0, 0, 8),
    )
    style.configure(
        "TNotebook.Tab",
        background=theme.BG,
        foreground=theme.TEXT_SECONDARY,
        padding=(16, 8),
        borderwidth=0,
        font=get_font(10, "bold"),
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", theme.BG), ("active", theme.BG)],
        foreground=[("selected", theme.PRIMARY), ("active", theme.TEXT)],
        focuscolor=[("selected", theme.BG)],
    )

    style.configure(
        "TScrollbar",
        background=theme.BORDER_STRONG,
        troughcolor=theme.BG,
        bordercolor=theme.BG,
        lightcolor=theme.BORDER_STRONG,
        darkcolor=theme.BORDER_STRONG,
        arrowcolor=theme.TEXT_MUTED,
        relief="flat",
        width=px(18),
        arrowsize=px(15),
    )
    style.map("TScrollbar", background=[("active", theme.TEXT_MUTED)])

    style.configure(
        "Horizontal.TProgressbar",
        troughcolor=theme.BORDER_LIGHT,
        bordercolor=theme.BORDER_LIGHT,
        background=theme.PRIMARY,
        lightcolor=theme.PRIMARY,
        darkcolor=theme.PRIMARY,
        thickness=6,
    )


class RoundedCard(tk.Frame):
    """白底 hairline 描边分组卡片。默认不再绘制伪阴影。"""

    def __init__(self, parent, theme, padding=20, radius=12, fill=None, shadow=False, border=True, **kwargs):
        parent_bg = _widget_bg(parent, theme.BG)
        super().__init__(parent, bg=parent_bg, **kwargs)
        self.theme = theme
        self.padding = padding
        self.radius = radius
        self.fill = fill or theme.CARD
        self.shadow = shadow
        self.border = border

        self.canvas = tk.Canvas(self, bg=parent_bg, highlightthickness=0, bd=0, height=80)
        self.canvas.pack(fill="x")
        self.content = tk.Frame(self.canvas, bg=self.fill)
        self._window = self.canvas.create_window(
            padding,
            padding,
            anchor="nw",
            window=self.content,
        )
        self.content.bind("<Configure>", self._sync_height)
        self.canvas.bind("<Configure>", self._redraw)

    def _sync_height(self, _event=None):
        requested = self.content.winfo_reqheight() + self.padding * 2
        if int(float(self.canvas.cget("height"))) != requested:
            self.canvas.configure(height=requested)
        self._redraw()

    def _redraw(self, _event=None):
        width = max(2, self.canvas.winfo_width())
        height = max(2, int(float(self.canvas.cget("height"))))
        self.canvas.delete("surface")
        if self.shadow:
            # 柔和单层投影：仅向下偏移，避免左右边缘出现脏边。
            self.canvas.create_polygon(
                rounded_rect_points(2, 3, width - 3, height - 1, self.radius),
                smooth=True,
                fill=self.theme.SHADOW,
                outline="",
                tags="surface",
            )
        self.canvas.create_polygon(
            rounded_rect_points(1, 1, width - 2, height - 2, self.radius),
            smooth=True,
            fill=self.fill,
            outline=_hairline(self.theme) if self.border else "",
            width=1 if self.border else 0,
            tags="surface",
        )
        self.canvas.tag_lower("surface")
        inner_width = max(1, width - self.padding * 2 - 4)
        self.canvas.itemconfigure(self._window, width=inner_width)


class HoverTooltip:
    """Delayed tooltip for compact modern controls, including disabled ones."""

    def __init__(self, owner, text, theme, get_font, delay_ms=420, wraplength=330):
        self.owner = owner
        self.theme = theme
        self.get_font = get_font
        self.text = str(text or "").strip()
        self.delay_ms = delay_ms
        self.wraplength = wraplength
        self._after_id = None
        self._tip = None
        owner.bind("<Destroy>", self._on_owner_destroy, add="+")

    def set_text(self, text):
        self.text = str(text or "").strip()
        if not self.text:
            self.hide()

    def schedule(self):
        if not self.text:
            return
        self.hide()
        try:
            self._after_id = self.owner.after(self.delay_ms, self.show)
        except (RuntimeError, tk.TclError):
            self._after_id = None

    def show(self):
        self._after_id = None
        if self._tip is not None or not self.text:
            return
        try:
            if not self.owner.winfo_exists():
                return
            self._tip = tk.Toplevel(self.owner)
            self._tip.wm_overrideredirect(True)
            self._tip.configure(bg=self.theme.TEXT)
            try:
                self._tip.attributes("-topmost", True)
            except tk.TclError:
                pass
            tk.Label(
                self._tip,
                text=self.text,
                font=self.get_font(9),
                bg=self.theme.TEXT,
                fg="white",
                justify="left",
                wraplength=px(self.wraplength),
                padx=10,
                pady=8,
            ).pack()
            self._tip.update_idletasks()
            x = self.owner.winfo_rootx()
            y = self.owner.winfo_rooty() + self.owner.winfo_height() + px(8)
            tip_w = self._tip.winfo_reqwidth()
            tip_h = self._tip.winfo_reqheight()
            screen_w = self.owner.winfo_screenwidth()
            screen_h = self.owner.winfo_screenheight()
            x = max(8, min(x, screen_w - tip_w - 8))
            if y + tip_h > screen_h - 8:
                y = self.owner.winfo_rooty() - tip_h - px(8)
            self._tip.wm_geometry(f"+{x}+{max(8, y)}")
        except (RuntimeError, tk.TclError):
            self._tip = None

    def hide(self):
        if self._after_id is not None:
            try:
                self.owner.after_cancel(self._after_id)
            except (RuntimeError, tk.TclError):
                pass
            self._after_id = None
        if self._tip is not None:
            try:
                self._tip.destroy()
            except (RuntimeError, tk.TclError):
                pass
            self._tip = None

    def _on_owner_destroy(self, event=None):
        if event is None or event.widget is self.owner:
            self.hide()


class ModernButton(tk.Canvas):
    """统一按钮。

    - primary=True         → 品牌朱砂实心，白字（主要操作）
    - accent=..., tonal=True  → 浅色底 + 强调色文字（提示类操作）
    - accent=...           → 强调色实心（兼容旧调用）
    - 默认                  → 白底细描边（次要操作）
    """

    def __init__(
        self,
        parent,
        text,
        command,
        theme,
        get_font,
        primary=False,
        accent=None,
        tonal=False,
        height=None,
        width=None,
        min_width=72,
        font_size=10,
        horizontal_padding=36,
        latin_font_family="Times New Roman",
        radius=10,
        tooltip_text="",
        **kwargs,
    ):
        self.theme = theme
        self.get_font = get_font
        self.command = command
        self.enabled = True
        self.radius = radius
        self.disabled_fill = getattr(theme, "INPUT_BG", theme.BG)
        self.tonal = bool(tonal)
        self._is_secondary = not primary and accent is None

        if accent is not None and self.tonal:
            self.normal_fill = _accent_light(theme, accent)
            self.hover_fill = _accent_light_hover(theme, accent)
            self.text_color = accent
        elif accent is not None:
            self.normal_fill = accent
            self.hover_fill = accent
            self.text_color = "white"
        elif primary:
            self.normal_fill = theme.PRIMARY
            self.hover_fill = theme.PRIMARY_HOVER
            self.text_color = "white"
        else:
            self.normal_fill = theme.CARD
            self.hover_fill = theme.CONTROL_HOVER
            self.text_color = theme.TEXT

        parent_bg = _widget_bg(parent, theme.BG)
        font_spec = get_font(font_size, "bold")
        text_width = (
            measure_mixed_text(parent, text, font_spec, latin_font_family)
            if latin_font_family else measure_text(parent, text, font_spec)
        )
        requested_width = width or max(
            px(min_width), text_width + px(horizontal_padding)
        )
        height = height or px(38)
        super().__init__(
            parent,
            width=requested_width,
            height=height,
            bg=parent_bg,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
            **kwargs,
        )
        self._shape = self.create_polygon(
            0, 0, 1, 0, 1, 1,
            smooth=True,
            fill=self.normal_fill,
            outline="",
        )
        if latin_font_family:
            self.label = MixedFontLabel(
                self,
                text,
                font_spec,
                latin_font_family,
                self.normal_fill,
                self.text_color,
                cursor="hand2",
            )
        else:
            self.label = tk.Label(
                self,
                text=text,
                font=font_spec,
                bg=self.normal_fill,
                fg=self.text_color,
                cursor="hand2",
                padx=4,
                pady=0,
            )
        self._label_window = self.create_window(0, 0, window=self.label)
        self._tooltip = HoverTooltip(self, tooltip_text, theme, get_font) if tooltip_text else None
        self.bind("<Configure>", self._redraw)
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        if isinstance(self.label, MixedFontLabel):
            self.label.bind_parts("<Button-1>", self._click)
            self.label.bind_parts("<Enter>", self._enter)
            self.label.bind_parts("<Leave>", self._leave)
        else:
            self.label.bind("<Button-1>", self._click)
            self.label.bind("<Enter>", self._enter)
            self.label.bind("<Leave>", self._leave)
        self._redraw()

    def _outline_for(self, fill):
        if self._is_secondary and fill in (self.theme.CARD, self.theme.CONTROL_HOVER):
            return _control_border(self.theme)
        return ""

    def _redraw(self, _event=None, fill=None):
        width = max(2, self.winfo_width())
        height = max(2, self.winfo_height())
        color = self.disabled_fill if not self.enabled else (fill or self.normal_fill)
        outline = _hairline(self.theme) if not self.enabled else self._outline_for(color)
        self.coords(self._shape, *rounded_rect_points(1, 1, width - 2, height - 2, self.radius))
        self.itemconfigure(self._shape, fill=color, outline=outline, width=1 if outline else 0)
        self.coords(self._label_window, width / 2, height / 2)
        if isinstance(self.label, MixedFontLabel):
            self.label.set_style(
                bg=color,
                fg=self.text_color if self.enabled else self.theme.TEXT_MUTED,
            )
        else:
            self.label.configure(
                bg=color,
                fg=self.text_color if self.enabled else self.theme.TEXT_MUTED,
            )

    def _click(self, _event=None):
        if self.enabled and self.command:
            self.command()

    def _enter(self, _event=None):
        if self._tooltip is not None:
            self._tooltip.schedule()
        if self.enabled:
            self._redraw(fill=self.hover_fill)

    def _leave(self, _event=None):
        if self._tooltip is not None:
            self._tooltip.hide()
        self._redraw()

    def set_enabled(self, enabled):
        self.enabled = bool(enabled)
        cursor = "hand2" if self.enabled else "arrow"
        self.configure(cursor=cursor)
        if isinstance(self.label, MixedFontLabel):
            self.label.set_style(
                cursor=cursor,
                fg=self.text_color if self.enabled else self.theme.TEXT_MUTED,
            )
        else:
            self.label.configure(
                cursor=cursor,
                fg=self.text_color if self.enabled else self.theme.TEXT_MUTED,
            )
        self._redraw()

    def set_tooltip_text(self, text):
        if self._tooltip is None and text:
            self._tooltip = HoverTooltip(self, text, self.theme, self.get_font)
            return
        if self._tooltip is not None:
            self._tooltip.set_text(text)

    def set_text(self, text):
        """Update the visible caption without forwarding it to ``tk.Canvas``."""
        self.label.configure(text=str(text))
        self._redraw()

    def configure(self, cnf=None, **kwargs):
        if not hasattr(self, "normal_fill"):
            return super().configure(cnf or {}, **kwargs)
        text = kwargs.pop("text", None)
        fill = kwargs.pop("bg", kwargs.pop("background", None))
        fg = kwargs.pop("fg", kwargs.pop("foreground", None))
        state = kwargs.pop("state", None)
        kwargs.pop("highlightbackground", None)
        if fill is not None:
            self.normal_fill = fill
            self._is_secondary = fill == self.theme.CARD
            if fill == self.theme.PRIMARY:
                self.hover_fill = self.theme.PRIMARY_HOVER
            elif fill == self.theme.CARD:
                self.hover_fill = self.theme.CONTROL_HOVER
            elif fill == self.theme.CONTROL_BG:
                self.hover_fill = self.theme.CONTROL_HOVER
            elif fill == self.theme.TEXT_MUTED:
                self.hover_fill = fill
            else:
                self.hover_fill = fill
        if fg is not None:
            self.text_color = fg
            if isinstance(self.label, MixedFontLabel):
                self.label.set_style(fg=fg)
            else:
                self.label.configure(fg=fg)
        if state is not None:
            self.set_enabled(state != "disabled")
        result = super().configure(cnf or {}, **kwargs)
        if text is not None:
            self.label.configure(text=str(text))
        self._redraw()
        return result

    config = configure


class ChoiceChip(tk.Canvas):
    """单选卡片：左侧常驻指示圈，文字始终左对齐，选中态不发生布局跳动。"""

    def __init__(
        self,
        parent,
        title,
        value,
        variable,
        theme,
        get_font,
        command=None,
        subtitle="",
        compact=False,
        accent=None,
        tooltip_text="",
        **kwargs,
    ):
        self.theme = theme
        self.get_font = get_font
        self.value = value
        self.variable = variable
        self.command = command
        self.title = title
        self.subtitle = subtitle
        self.compact = compact
        self.accent = accent or theme.PRIMARY
        self.enabled = True
        self.selected = False
        self._tooltip = HoverTooltip(self, tooltip_text, theme, get_font) if tooltip_text else None
        height = px(40) if compact else px(64)
        parent_bg = _widget_bg(parent, theme.BG)
        super().__init__(
            parent,
            height=height,
            bg=parent_bg,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
            **kwargs,
        )
        self.bind("<Configure>", self._draw)
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self.variable.trace_add("write", lambda *_args: self._draw())
        self._draw()

    def _draw(self, _event=None, hover=False):
        self.delete("all")
        self.selected = self.variable.get() == self.value
        width = max(2, self.winfo_width())
        height = max(2, self.winfo_height())
        theme = self.theme
        if not self.enabled:
            fill, border, border_w = theme.INPUT_BG, _hairline(theme), 1
            title_fg = theme.TEXT_MUTED
            sub_fg = theme.TEXT_MUTED
        elif self.selected:
            fill, border, border_w = _accent_light(theme, self.accent), self.accent, 2
            title_fg = theme.TEXT
            sub_fg = theme.TEXT_SECONDARY
        elif hover:
            fill, border, border_w = theme.CONTROL_HOVER, _control_border(theme), 1
            title_fg = theme.TEXT
            sub_fg = theme.TEXT_SECONDARY
        else:
            fill, border, border_w = theme.CARD, _control_border(theme), 1
            title_fg = theme.TEXT
            sub_fg = theme.TEXT_SECONDARY
        self.create_polygon(
            rounded_rect_points(1, 1, width - 2, height - 2, 10),
            smooth=True,
            fill=fill,
            outline=border,
            width=border_w,
        )

        # 常驻指示圈：未选中为空心圆，选中为强调色圆 + 白点。
        ring_cx = px(20)
        ring_cy = height / 2
        ring_r = px(7)
        if not self.enabled:
            self.create_oval(
                ring_cx - ring_r, ring_cy - ring_r, ring_cx + ring_r, ring_cy + ring_r,
                outline=theme.BORDER, width=1.5, fill=theme.INPUT_BG,
            )
        elif self.selected:
            self.create_oval(
                ring_cx - ring_r, ring_cy - ring_r, ring_cx + ring_r, ring_cy + ring_r,
                outline="", fill=self.accent,
            )
            dot_r = max(2.5, px(2.5))
            self.create_oval(
                ring_cx - dot_r, ring_cy - dot_r, ring_cx + dot_r, ring_cy + dot_r,
                outline="", fill="white",
            )
        else:
            self.create_oval(
                ring_cx - ring_r, ring_cy - ring_r, ring_cx + ring_r, ring_cy + ring_r,
                outline=_control_border(theme), width=1.5, fill=theme.CARD,
            )

        text_x = ring_cx + ring_r + px(10)
        title_y = height / 2 if self.compact or not self.subtitle else height / 2 - px(10)
        self.create_text(
            text_x,
            title_y,
            text=self.title,
            anchor="w",
            fill=title_fg,
            font=self.get_font(10, "bold"),
        )
        if self.subtitle and not self.compact:
            self.create_text(
                text_x,
                height / 2 + px(12),
                text=self.subtitle,
                anchor="w",
                fill=sub_fg,
                font=self.get_font(9),
            )

    def _click(self, _event=None):
        if not self.enabled:
            return
        self.variable.set(self.value)
        if self.command:
            self.command()

    def _enter(self, _event=None):
        if self._tooltip is not None:
            self._tooltip.schedule()
        if self.enabled and not self.selected:
            self._draw(hover=True)

    def _leave(self, _event=None):
        if self._tooltip is not None:
            self._tooltip.hide()
        self._draw()

    def set_enabled(self, enabled):
        self.enabled = bool(enabled)
        self.configure(cursor="hand2" if self.enabled else "arrow")
        self._draw()

    def set_tooltip_text(self, text):
        if self._tooltip is None and text:
            self._tooltip = HoverTooltip(self, text, self.theme, self.get_font)
            return
        if self._tooltip is not None:
            self._tooltip.set_text(text)


class ModernSelectionControl(tk.Canvas):
    """Large, DPI-aware checkbox/radiobutton with a generous click target.

    Tk's native indicator stays extremely small on some Windows DPI/font
    combinations.  This canvas control keeps the familiar square/circle
    semantics while making the indicator and its check/dot visibly larger.
    """

    def __init__(
        self,
        parent,
        text,
        variable,
        theme,
        get_font,
        *,
        kind="check",
        value=None,
        command=None,
        font_size=11,
        bg=None,
        fg=None,
        indicator_size=22,
        enabled=True,
        **kwargs,
    ):
        self.theme = theme
        self.variable = variable
        self.kind = "radio" if kind == "radio" else "check"
        self.value = value
        self.command = command
        self.font_spec = get_font(font_size)
        self.text = str(text or "")
        self.base_bg = bg or _widget_bg(parent, theme.BG)
        self.text_color = fg or theme.TEXT
        self.enabled = bool(enabled)
        self.hovered = False
        self.indicator_size = px(max(20, indicator_size))
        self.control_height = max(px(34), self.indicator_size + px(8))
        text_width = measure_text(parent, self.text, self.font_spec)
        control_width = self.indicator_size + px(10) + text_width + px(6)
        super().__init__(
            parent,
            width=control_width,
            height=self.control_height,
            bg=self.base_bg,
            highlightthickness=0,
            bd=0,
            cursor="hand2" if self.enabled else "arrow",
            takefocus=1,
            **kwargs,
        )
        self.bind("<Button-1>", self._invoke)
        self.bind("<space>", self._invoke)
        self.bind("<Return>", self._invoke)
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self.bind("<FocusIn>", self._draw)
        self.bind("<FocusOut>", self._draw)
        self._trace_id = self.variable.trace_add("write", self._draw)
        self._draw()

    def _selected(self):
        if self.kind == "radio":
            return self.variable.get() == self.value
        return bool(self.variable.get())

    def _draw(self, *_args):
        if not self.winfo_exists():
            return
        self.delete("all")
        selected = self._selected()
        theme = self.theme
        size = self.indicator_size
        left = px(2)
        top = (self.control_height - size) / 2
        right = left + size
        bottom = top + size
        accent = theme.PRIMARY
        border = theme.TEXT_MUTED if self.enabled else theme.BORDER
        fill = accent if selected and self.enabled else (
            theme.INPUT_BG if not self.enabled else theme.CARD
        )

        if self.hovered and self.enabled and not selected:
            border = accent
        if self.kind == "radio":
            self.create_oval(
                left, top, right, bottom,
                fill=fill,
                outline=accent if selected and self.enabled else border,
                width=max(2, px(2)),
            )
            if selected:
                dot_r = max(px(4), size * 0.23)
                cx = (left + right) / 2
                cy = (top + bottom) / 2
                self.create_oval(
                    cx - dot_r, cy - dot_r, cx + dot_r, cy + dot_r,
                    fill="white" if self.enabled else theme.TEXT_MUTED,
                    outline="",
                )
        else:
            radius = px(4)
            self.create_polygon(
                rounded_rect_points(left, top, right, bottom, radius),
                smooth=True,
                fill=fill,
                outline=accent if selected and self.enabled else border,
                width=max(2, px(2)),
            )
            if selected:
                check_color = "white" if self.enabled else theme.TEXT_MUTED
                self.create_line(
                    left + size * 0.23,
                    top + size * 0.52,
                    left + size * 0.43,
                    top + size * 0.72,
                    left + size * 0.78,
                    top + size * 0.30,
                    fill=check_color,
                    width=max(2, px(2.5)),
                    capstyle=tk.ROUND,
                    joinstyle=tk.ROUND,
                )

        if self.focus_get() is self:
            pad = px(2)
            self.create_rectangle(
                0, pad, max(1, self.winfo_width() - 1), self.control_height - pad,
                outline=theme.PRIMARY,
                width=1,
                dash=(2, 2),
            )
        self.create_text(
            right + px(9),
            self.control_height / 2,
            text=self.text,
            anchor="w",
            fill=self.text_color if self.enabled else theme.TEXT_MUTED,
            font=self.font_spec,
        )

    def _invoke(self, _event=None):
        if not self.enabled:
            return "break"
        self.focus_set()
        if self.kind == "radio":
            self.variable.set(self.value)
        else:
            self.variable.set(not bool(self.variable.get()))
        if self.command:
            self.command()
        return "break"

    def _enter(self, _event=None):
        self.hovered = True
        self._draw()

    def _leave(self, _event=None):
        self.hovered = False
        self._draw()

    def set_enabled(self, enabled):
        self.enabled = bool(enabled)
        self.configure(cursor="hand2" if self.enabled else "arrow")
        self._draw()


class ModernCheckbutton(ModernSelectionControl):
    def __init__(self, parent, text, variable, theme, get_font, **kwargs):
        super().__init__(
            parent, text, variable, theme, get_font,
            kind="check", **kwargs,
        )


class ModernRadiobutton(ModernSelectionControl):
    def __init__(self, parent, text, value, variable, theme, get_font, **kwargs):
        super().__init__(
            parent, text, variable, theme, get_font,
            kind="radio", value=value, **kwargs,
        )


class ModernEntry(tk.Canvas):
    """白底细描边输入框，聚焦时描边变为主色。"""

    def __init__(self, parent, theme, get_font, textvariable=None, show=None, height=None, **kwargs):
        self.theme = theme
        self.radius = 9
        height = height or px(40)
        parent_bg = _widget_bg(parent, theme.BG)
        super().__init__(
            parent,
            height=height,
            bg=parent_bg,
            highlightthickness=0,
            bd=0,
            **kwargs,
        )
        self._shape = self.create_polygon(
            0, 0, 1, 0, 1, 1,
            smooth=True,
            fill=theme.CARD,
            outline=_control_border(theme),
        )
        self.entry = tk.Entry(
            self,
            textvariable=textvariable,
            show=show,
            font=get_font(11),
            bg=theme.CARD,
            fg=theme.TEXT,
            relief="flat",
            bd=0,
            highlightthickness=0,
            insertbackground=theme.TEXT,
        )
        self._entry_window = self.create_window(px(12), height / 2, anchor="w", window=self.entry)
        self.bind("<Configure>", self._redraw)
        self.entry.bind("<FocusIn>", lambda _event: self._redraw(border=theme.PRIMARY))
        self.entry.bind("<FocusOut>", lambda _event: self._redraw())
        self._redraw()

    def _redraw(self, _event=None, border=None):
        width = max(2, self.winfo_width())
        height = max(2, self.winfo_height())
        self.coords(self._shape, *rounded_rect_points(1, 1, width - 2, height - 2, self.radius))
        self.itemconfigure(
            self._shape,
            outline=border or _control_border(self.theme),
            width=2 if border else 1,
        )
        self.itemconfigure(self._entry_window, width=max(1, width - px(24)))
        self.coords(self._entry_window, px(12), height / 2)

    def configure(self, cnf=None, **kwargs):
        if not hasattr(self, "entry"):
            return super().configure(cnf or {}, **kwargs)
        state = kwargs.pop("state", None)
        if state is not None:
            self.entry.configure(state=state)
            fill = self.theme.INPUT_BG if state == "disabled" else self.theme.CARD
            self.itemconfigure(self._shape, fill=fill)
            self.entry.configure(bg=fill)
        return super().configure(cnf or {}, **kwargs)

    config = configure


def _widget_bg(widget, fallback):
    try:
        return widget.cget("bg")
    except (tk.TclError, AttributeError):
        return fallback


def fit_combobox_width(combo, values, min_width=8, max_chars=34):
    """实测当前字体下最长选项的像素宽度，换算为字符宽度赋给下拉框。

    ttk.Combobox 的 width 以字符计，同样的字符数在不同字体、字号和 DPI 下
    实际像素差异很大，写死容易裁掉文字（如“自定义(45.0pt)”）。这里按真实
    字体度量最长选项并预留下拉箭头空间；max_chars 防止异常长值把行撑爆。
    """
    try:
        font_spec = combo.cget('font')
        if not font_spec:
            font_spec = ttk.Style(combo).lookup('TCombobox', 'font')
        font = tkfont.Font(font=font_spec or 'TkDefaultFont')
        unit = max(1, font.measure('0'))
        longest = max(
            (font.measure(str(value)) for value in list(values) or ['']),
            default=0,
        )
        fitted = int((longest + px(30)) / unit) + 1
        combo.configure(width=max(int(min_width), min(int(max_chars), fitted)))
    except (tk.TclError, TypeError, ValueError):
        pass
