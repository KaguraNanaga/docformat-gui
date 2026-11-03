"""Consistent, parent-relative placement for the application's work windows."""

from __future__ import annotations


def parent_relative_size(
    preferred_width,
    preferred_height,
    *,
    parent_width,
    parent_height,
    screen_width,
    screen_height,
    fraction=0.80,
    height_fraction=0.88,
):
    """Return a roomy size that stays within both the parent and the screen.

    Work windows should feel like a focused page of the application, not a
    miniature dialog.  They start at 80% of the visible main-window width and
    88% of its height, so vertically dense work pages remain usable without
    outer scrolling.  A smaller main window is never exceeded, and a narrow
    or short display still retains a small outer safety margin.
    """
    parent_width = max(1, int(parent_width))
    parent_height = max(1, int(parent_height))
    available_width = max(480, int(screen_width) - 48)
    available_height = max(420, int(screen_height) - 64)
    target_width = max(int(round(parent_width * fraction)), min(int(preferred_width), parent_width))
    target_height = max(
        int(round(parent_height * height_fraction)),
        min(int(preferred_height), parent_height),
    )
    return (
        min(target_width, parent_width, available_width),
        min(target_height, parent_height, available_height),
    )


def parent_relative_geometry(
    preferred_width,
    preferred_height,
    *,
    parent_width,
    parent_height,
    parent_x,
    parent_y,
    screen_width,
    screen_height,
    screen_x=0,
    screen_y=0,
    fraction=0.80,
    height_fraction=0.88,
):
    """Return ``(width, height, x, y)`` centred in the visible parent area."""
    width, height = parent_relative_size(
        preferred_width,
        preferred_height,
        parent_width=parent_width,
        parent_height=parent_height,
        screen_width=screen_width,
        screen_height=screen_height,
        fraction=fraction,
        height_fraction=height_fraction,
    )
    preferred_x = int(parent_x) + (int(parent_width) - width) // 2
    preferred_y = int(parent_y) + (int(parent_height) - height) // 2
    x = max(int(screen_x) + 12, min(preferred_x, int(screen_x) + int(screen_width) - width - 12))
    y = max(int(screen_y) + 12, min(preferred_y, int(screen_y) + int(screen_height) - height - 12))
    return width, height, x, y


def apply_parent_relative_layout(
    window,
    parent,
    *,
    preferred_width,
    preferred_height,
    min_width,
    min_height,
    fraction=0.80,
    height_fraction=0.88,
):
    """Size and position a Tk window as an 80%-of-parent work surface."""
    parent.update_idletasks()
    window.update_idletasks()
    width, height, x, y = parent_relative_geometry(
        preferred_width,
        preferred_height,
        parent_width=parent.winfo_width(),
        parent_height=parent.winfo_height(),
        parent_x=parent.winfo_rootx(),
        parent_y=parent.winfo_rooty(),
        screen_width=window.winfo_vrootwidth(),
        screen_height=window.winfo_vrootheight(),
        screen_x=window.winfo_vrootx(),
        screen_y=window.winfo_vrooty(),
        fraction=fraction,
        height_fraction=height_fraction,
    )
    window.geometry(f"{width}x{height}+{x}+{y}")
    window.minsize(min(int(min_width), width), min(int(min_height), height))
    return width, height, x, y
