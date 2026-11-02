from types import SimpleNamespace

import pytest

tk = pytest.importorskip("tkinter")

import docformat_gui


def test_tkdnd_failure_uses_fallback_as_default_root(monkeypatch):
    try:
        tk.Tk().destroy()
    except tk.TclError as exc:
        pytest.skip(f"Tk display unavailable: {exc}")

    class BrokenDnDRoot(tk.Tk):
        def __init__(self):
            tk.Tk.__init__(self)
            raise RuntimeError("Unable to load tkdnd library.")

    monkeypatch.setattr(
        docformat_gui, "TkinterDnD", SimpleNamespace(Tk=BrokenDnDRoot), raising=False
    )
    monkeypatch.setattr(docformat_gui, "_DND_AVAILABLE", True)

    root = docformat_gui._create_root()
    try:
        assert tk._default_root is root
        v = tk.StringVar(value="3.7")
        e = tk.Entry(root, textvariable=v)
        root.update_idletasks()
        assert e.get() == "3.7"
        assert docformat_gui._DND_AVAILABLE is False

        e.delete(0, tk.END)
        e.insert(0, "4.2")
        assert v.get() == "4.2"
    finally:
        root.destroy()
