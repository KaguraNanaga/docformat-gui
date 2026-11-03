"""Tk dialogs for editable paragraph types and deterministic match rules."""

from __future__ import annotations

from copy import deepcopy
import tkinter as tk
from tkinter import messagebox, ttk

from window_layout import apply_parent_relative_layout

from scripts.layout_rules import (
    BASE_PARAGRAPH_TYPES,
    PARAGRAPH_RULE_VERSION,
    ParagraphRuleError,
    built_in_paragraph_rules,
    built_in_paragraph_types,
    ensure_paragraph_rule_defaults,
    merge_rule_pack,
    validate_manual_id,
    validate_paragraph_rules,
    validate_paragraph_types,
)


ALIGN_LABELS = {"left": "左对齐", "center": "居中", "right": "右对齐", "justify": "两端对齐"}


def _float(value, fallback=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(fallback)


def _wait_window_and_restore_grab(parent, dialog):
    """Restore the manager's modal grab after a nested editor closes."""
    try:
        parent.wait_window(dialog)
    finally:
        try:
            if parent.winfo_exists():
                parent.grab_set()
        except (tk.TclError, RuntimeError):
            pass


class ParagraphTypeEditor(tk.Toplevel):
    def __init__(self, parent, value=None, body_format=None):
        super().__init__(parent)
        self.result = None
        self.value = deepcopy(value or {})
        self.body_format = deepcopy(body_format or {})
        self.title("编辑段落类型")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        fmt = {**self.body_format, **(self.value.get("format") or {})}
        form = ttk.Frame(self, padding=16)
        form.pack(fill="both", expand=True)
        self.vars = {
            "id": tk.StringVar(value=self.value.get("id", "custom_type")),
            "name": tk.StringVar(value=self.value.get("name", "新段落类型")),
            "font_cn": tk.StringVar(value=fmt.get("font_cn", "宋体")),
            "font_en": tk.StringVar(value=fmt.get("font_en", "Times New Roman")),
            "size": tk.StringVar(value=str(fmt.get("size", 16))),
            "bold": tk.BooleanVar(value=bool(fmt.get("bold", False))),
            "align": tk.StringVar(value=ALIGN_LABELS.get(fmt.get("align", "left"), "左对齐")),
            "indent": tk.StringVar(value=str(fmt.get("indent", 0))),
            "line_spacing": tk.StringVar(value=str(fmt.get("line_spacing", 28))),
            "space_before": tk.StringVar(value=str(fmt.get("space_before", 0))),
            "space_after": tk.StringVar(value=str(fmt.get("space_after", 0))),
        }
        fields = [
            ("类型 ID（英文/数字/下划线）", "id"), ("显示名称", "name"), ("中文字体", "font_cn"),
            ("英数字体", "font_en"), ("字号（pt）", "size"),
            ("首行缩进（pt）", "indent"), ("行距（pt）", "line_spacing"),
            ("段前（pt）", "space_before"), ("段后（pt）", "space_after"),
        ]
        for row, (label, key) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="e", padx=(0, 8), pady=4)
            ttk.Entry(form, textvariable=self.vars[key], width=28).grid(row=row, column=1, sticky="ew", pady=4)
        row = len(fields)
        ttk.Label(form, text="对齐").grid(row=row, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Combobox(
            form, textvariable=self.vars["align"], state="readonly", width=25,
            values=list(ALIGN_LABELS.values()),
        ).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1
        ttk.Checkbutton(form, text="加粗", variable=self.vars["bold"]).grid(row=row, column=1, sticky="w", pady=4)
        buttons = ttk.Frame(form)
        buttons.grid(row=row + 1, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="确定", command=self._save).pack(side="right")
        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Return>", lambda _event: self._save())

    def _save(self):
        try:
            fmt = {
                "font_cn": self.vars["font_cn"].get().strip() or "宋体",
                "font_en": self.vars["font_en"].get().strip() or "Times New Roman",
                "size": _float(self.vars["size"].get(), 16),
                "bold": self.vars["bold"].get(),
                "align": next(
                    (key for key, label in ALIGN_LABELS.items() if label == self.vars["align"].get()),
                    "left",
                ),
                "indent": _float(self.vars["indent"].get(), 0),
                "line_spacing": _float(self.vars["line_spacing"].get(), 28),
                "space_before": _float(self.vars["space_before"].get(), 0),
                "space_after": _float(self.vars["space_after"].get(), 0),
            }
            candidate = {
                "id": validate_manual_id(self.vars["id"].get(), "类型 ID"),
                "name": self.vars["name"].get().strip(),
                "format": fmt,
            }
            self.result = validate_paragraph_types([candidate])[0]
        except ParagraphRuleError as exc:
            messagebox.showerror("输入错误", str(exc), parent=self)
            return
        self.destroy()


class ParagraphRuleEditor(tk.Toplevel):
    def __init__(self, parent, types, value=None):
        super().__init__(parent)
        self.result = None
        self.types = deepcopy(types)
        self.value = deepcopy(value or {})
        self.title("编辑匹配规则")
        self.transient(parent)
        apply_parent_relative_layout(
            self, parent,
            preferred_width=780, preferred_height=520,
            min_width=680, min_height=460,
            fraction=0.70, height_fraction=0.72,
        )
        self.grab_set()

        type_ids = [item["id"] for item in self.types]
        form = ttk.Frame(self, padding=16)
        form.pack(fill="both", expand=True)
        self.vars = {
            "id": tk.StringVar(value=self.value.get("id", "rule_custom")),
            "name": tk.StringVar(value=self.value.get("name", "新匹配规则")),
            "type_id": tk.StringVar(value=self.value.get("type_id", type_ids[0] if type_ids else "")),
            "pattern": tk.StringVar(value=self.value.get("pattern", r"^\S.+$")),
            "previous_type": tk.StringVar(value=self.value.get("previous_type", "")),
            "next_pattern": tk.StringVar(value=self.value.get("next_pattern", "")),
            "next_next_pattern": tk.StringVar(value=self.value.get("next_next_pattern", "")),
            "priority": tk.StringVar(value=str(self.value.get("priority", 100))),
            "enabled": tk.BooleanVar(value=bool(self.value.get("enabled", True))),
        }
        fields = [
            ("规则 ID", "id", "entry"), ("规则名称", "name", "entry"),
            ("输出类型", "type_id", "type"), ("本段正则", "pattern", "entry"),
            ("前一段类型", "previous_type", "previous"),
            ("下一段正则", "next_pattern", "entry"),
            ("下两段正则", "next_next_pattern", "entry"),
            ("优先级", "priority", "entry"),
        ]
        for row, (label, key, kind) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="ne", padx=(0, 8), pady=6)
            if kind == "type":
                widget = ttk.Combobox(form, textvariable=self.vars[key], values=type_ids, state="readonly")
            elif kind == "previous":
                widget = ttk.Combobox(
                    form, textvariable=self.vars[key],
                    values=[""] + sorted(BASE_PARAGRAPH_TYPES) + type_ids,
                    state="readonly",
                )
            else:
                widget = ttk.Entry(form, textvariable=self.vars[key])
            widget.grid(row=row, column=1, sticky="ew", pady=6)
        form.columnconfigure(1, weight=1)
        ttk.Label(
            form,
            text="留空表示不检查上下文。处理时按优先级从高到低，首个匹配规则生效。",
            foreground="#666666",
        ).grid(row=len(fields), column=0, columnspan=2, sticky="w", pady=(4, 8))
        ttk.Checkbutton(form, text="启用规则", variable=self.vars["enabled"]).grid(
            row=len(fields) + 1, column=1, sticky="w"
        )
        buttons = ttk.Frame(form)
        buttons.grid(row=len(fields) + 2, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(buttons, text="取消", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="确定", command=self._save).pack(side="right")
        self.bind("<Escape>", lambda _event: self.destroy())

    def _save(self):
        candidate = {key: var.get() for key, var in self.vars.items()}
        try:
            candidate["id"] = validate_manual_id(candidate.get("id"), "规则 ID")
            self.result = validate_paragraph_rules([candidate], self.types)[0]
        except (ParagraphRuleError, IndexError) as exc:
            messagebox.showerror("输入错误", str(exc), parent=self)
            return
        self.destroy()


class ParagraphRulesDialog(tk.Toplevel):
    """Manage local paragraph types and deterministic matching rules."""

    def __init__(self, parent, settings, on_save=None):
        super().__init__(parent)
        self.settings = deepcopy(settings or {})
        self.on_save = on_save
        ensure_paragraph_rule_defaults(self.settings)
        self.types = deepcopy(self.settings.get("paragraph_types", []))
        self.rules = deepcopy(self.settings.get("paragraph_rules", []))
        self.title("通用文档·段落类型与匹配规则")
        self.transient(parent)
        apply_parent_relative_layout(
            self, parent,
            preferred_width=1180, preferred_height=760,
            min_width=940, min_height=620,
        )
        self.grab_set()
        self._build()
        self._refresh()

    def _build(self):
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)
        ttk.Label(
            outer,
            text="所有规则均在本地匹配；保存前会检查正则表达式和上下文条件。",
            foreground="#555555",
        ).pack(anchor="w", pady=(0, 10))
        panes = ttk.Panedwindow(outer, orient="horizontal")
        panes.pack(fill="both", expand=True)

        type_panel = ttk.Labelframe(panes, text="段落类型（每类有独立字体、缩进和行距）", padding=8)
        rule_panel = ttk.Labelframe(panes, text="匹配规则（正则 + 上下文）", padding=8)
        panes.add(type_panel, weight=1)
        panes.add(rule_panel, weight=2)

        self.type_tree = ttk.Treeview(type_panel, columns=("name", "font", "size", "align"), show="headings", height=18)
        for key, label, width in (("name", "名称", 120), ("font", "字体", 120), ("size", "字号", 60), ("align", "对齐", 80)):
            self.type_tree.heading(key, text=label)
            self.type_tree.column(key, width=width, anchor="w")
        self.type_tree.pack(fill="both", expand=True)
        type_buttons = ttk.Frame(type_panel)
        type_buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(type_buttons, text="+新建", command=self._add_type).pack(side="left")
        ttk.Button(type_buttons, text="编辑", command=self._edit_type).pack(side="left", padx=4)
        ttk.Button(type_buttons, text="删除", command=self._delete_type).pack(side="left")

        self.rule_tree = ttk.Treeview(
            rule_panel, columns=("name", "type", "pattern", "previous", "priority"),
            show="headings", height=18,
        )
        for key, label, width in (
            ("name", "规则", 140), ("type", "类型", 100), ("pattern", "本段正则", 230),
            ("previous", "前置类型", 95), ("priority", "优先级", 65),
        ):
            self.rule_tree.heading(key, text=label)
            self.rule_tree.column(key, width=width, anchor="w")
        self.rule_tree.pack(fill="both", expand=True)
        rule_buttons = ttk.Frame(rule_panel)
        rule_buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(rule_buttons, text="+新建", command=self._add_rule).pack(side="left")
        ttk.Button(rule_buttons, text="编辑", command=self._edit_rule).pack(side="left", padx=4)
        ttk.Button(rule_buttons, text="删除", command=self._delete_rule).pack(side="left")

        tools = ttk.Frame(outer)
        tools.pack(fill="x", pady=(12, 0))
        ttk.Button(tools, text="恢复章程/语录基础规则", command=self._restore_builtin).pack(side="left")
        self.status_var = tk.StringVar(value="")
        ttk.Label(tools, textvariable=self.status_var, foreground="#666666").pack(side="left", padx=12)
        ttk.Button(tools, text="取消", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(tools, text="应用到当前预设", command=self._save).pack(side="right")

    def _refresh(self):
        self.type_tree.delete(*self.type_tree.get_children())
        for item in self.types:
            fmt = item.get("format") or {}
            self.type_tree.insert("", "end", iid=item["id"], values=(
                item.get("name"), fmt.get("font_cn"), fmt.get("size"), ALIGN_LABELS.get(fmt.get("align"), fmt.get("align")),
            ))
        self.rule_tree.delete(*self.rule_tree.get_children())
        type_names = {item["id"]: item.get("name") for item in self.types}
        for item in sorted(self.rules, key=lambda value: int(value.get("priority", 0)), reverse=True):
            state = "" if item.get("enabled", True) else "[停用] "
            self.rule_tree.insert("", "end", iid=item["id"], values=(
                state + item.get("name", ""), type_names.get(item.get("type_id"), item.get("type_id")),
                item.get("pattern"), type_names.get(item.get("previous_type"), item.get("previous_type", "")),
                item.get("priority", 100),
            ))
        self.status_var.set(f"当前 {len(self.types)} 个类型、{len(self.rules)} 条规则")

    def _selected(self, tree, values):
        selected = tree.selection()
        if not selected:
            return None, -1
        item_id = selected[0]
        for index, item in enumerate(values):
            if item.get("id") == item_id:
                return item, index
        return None, -1

    def _add_type(self):
        dialog = ParagraphTypeEditor(self, body_format=self.settings.get("body", {}))
        _wait_window_and_restore_grab(self, dialog)
        if not dialog.result:
            return
        if any(item["id"] == dialog.result["id"] for item in self.types):
            messagebox.showerror("重复 ID", "已存在同名类型 ID。", parent=self)
            return
        self.types.append(dialog.result)
        self._refresh()

    def _edit_type(self):
        item, index = self._selected(self.type_tree, self.types)
        if item is None:
            return
        dialog = ParagraphTypeEditor(self, item, self.settings.get("body", {}))
        _wait_window_and_restore_grab(self, dialog)
        if not dialog.result:
            return
        old_id, new_id = item["id"], dialog.result["id"]
        if new_id != old_id and any(value["id"] == new_id for value in self.types):
            messagebox.showerror("重复 ID", "已存在同名类型 ID。", parent=self)
            return
        self.types[index] = dialog.result
        if new_id != old_id:
            for rule in self.rules:
                if rule.get("type_id") == old_id:
                    rule["type_id"] = new_id
                if rule.get("previous_type") == old_id:
                    rule["previous_type"] = new_id
        self._refresh()

    def _delete_type(self):
        item, index = self._selected(self.type_tree, self.types)
        if item is None:
            return
        related = [rule for rule in self.rules if item["id"] in {rule.get("type_id"), rule.get("previous_type")}]
        message = f"删除类型「{item.get('name')}」？"
        if related:
            message += f"\n将同时删除 {len(related)} 条引用它的规则。"
        if not messagebox.askyesno("确认删除", message, parent=self):
            return
        self.types.pop(index)
        self.rules = [rule for rule in self.rules if item["id"] not in {rule.get("type_id"), rule.get("previous_type")}]
        self._refresh()

    def _add_rule(self):
        if not self.types:
            messagebox.showinfo("提示", "请先新建一个段落类型。", parent=self)
            return
        dialog = ParagraphRuleEditor(self, self.types)
        _wait_window_and_restore_grab(self, dialog)
        if not dialog.result:
            return
        if any(item["id"] == dialog.result["id"] for item in self.rules):
            messagebox.showerror("重复 ID", "已存在同名规则 ID。", parent=self)
            return
        self.rules.append(dialog.result)
        self._refresh()

    def _edit_rule(self):
        item, index = self._selected(self.rule_tree, self.rules)
        if item is None:
            return
        dialog = ParagraphRuleEditor(self, self.types, item)
        _wait_window_and_restore_grab(self, dialog)
        if not dialog.result:
            return
        if dialog.result["id"] != item["id"] and any(value["id"] == dialog.result["id"] for value in self.rules):
            messagebox.showerror("重复 ID", "已存在同名规则 ID。", parent=self)
            return
        self.rules[index] = dialog.result
        self._refresh()

    def _delete_rule(self):
        item, index = self._selected(self.rule_tree, self.rules)
        if item is not None:
            self.rules.pop(index)
            self._refresh()

    def _merge_into_current(self, rule_pack):
        current = deepcopy(self.settings)
        current["paragraph_types"] = self.types
        current["paragraph_rules"] = self.rules
        try:
            merged = merge_rule_pack(current, rule_pack)
        except ParagraphRuleError as exc:
            messagebox.showerror("规则合并失败", str(exc), parent=self)
            self._refresh()
            return False
        self.types = merged["paragraph_types"]
        self.rules = merged["paragraph_rules"]
        self._refresh()
        return True

    def _restore_builtin(self):
        rule_pack = {"types": built_in_paragraph_types(self.settings), "rules": built_in_paragraph_rules()}
        self._merge_into_current(rule_pack)

    def _save(self):
        try:
            self.types = validate_paragraph_types(self.types)
            self.rules = validate_paragraph_rules(self.rules, self.types)
        except ParagraphRuleError as exc:
            messagebox.showerror("规则无效", str(exc), parent=self)
            return
        result = deepcopy(self.settings)
        result["layout_mode"] = "generic"
        result["paragraph_rule_version"] = PARAGRAPH_RULE_VERSION
        result["paragraph_types"] = self.types
        result["paragraph_rules"] = self.rules
        if self.on_save:
            self.on_save(result)
        self.destroy()
