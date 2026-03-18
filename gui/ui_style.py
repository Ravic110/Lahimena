"""
Shared UI styling helpers for Lahimena Tours (Tkinter/CustomTkinter).
"""

import tkinter as tk
from tkinter import ttk

try:
    import customtkinter as ctk

    CTK_AVAILABLE = True
except Exception:
    ctk = None
    CTK_AVAILABLE = False

from config import (
    BUTTON_FONT,
    BUTTON_BLUE,
    BUTTON_GRAY,
    BUTTON_GREEN,
    BUTTON_GREEN_HOVER,
    BUTTON_ORANGE,
    BUTTON_RED,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MUTED_TEXT_COLOR,
    PANEL_BG_COLOR,
    READONLY_BG_COLOR,
    TEXT_COLOR,
)


def configure_combobox_style(root):
    """Apply consistent combobox style."""
    style = ttk.Style(root)
    style.configure(
        "TCombobox",
        font=ENTRY_FONT,
        foreground=TEXT_COLOR,
        fieldbackground=INPUT_BG_COLOR,
        background=INPUT_BG_COLOR,
        arrowcolor=TEXT_COLOR,
        padding=5,
        relief="flat",
    )
    style.map(
        "TCombobox",
        foreground=[("readonly", TEXT_COLOR)],
        fieldbackground=[("readonly", INPUT_BG_COLOR)],
        background=[("readonly", INPUT_BG_COLOR)],
        selectforeground=[("readonly", TEXT_COLOR)],
        selectbackground=[("readonly", INPUT_BG_COLOR)],
    )
    root.option_add("*TCombobox*Listbox.background", INPUT_BG_COLOR)
    root.option_add("*TCombobox*Listbox.foreground", TEXT_COLOR)
    root.option_add("*TCombobox*Listbox.selectBackground", BUTTON_GREEN)
    root.option_add("*TCombobox*Listbox.selectForeground", "white")


def create_card(parent, title=None, tabs=None, show_controls=False, on_add=None, on_remove=None):
    """Create a rounded card container with optional title and tabs."""
    if CTK_AVAILABLE:
        card = ctk.CTkFrame(
            parent,
            fg_color=PANEL_BG_COLOR,
            corner_radius=18,
            border_width=1,
            border_color="#C9DDE3",
        )
    else:
        card = tk.Frame(
            parent,
            bg=PANEL_BG_COLOR,
            highlightbackground="#C9DDE3",
            highlightthickness=1,
            bd=0,
        )
    card.pack(fill="x", pady=8)

    if CTK_AVAILABLE:
        inner = ctk.CTkFrame(card, fg_color="transparent")
    else:
        inner = tk.Frame(card, bg=PANEL_BG_COLOR)
    inner.pack(fill="both", expand=True, padx=12, pady=10)

    if tabs or title:
        header_row = tk.Frame(inner, bg=PANEL_BG_COLOR)
        header_row.pack(fill="x", pady=(0, 8))

        chips = tk.Frame(header_row, bg=PANEL_BG_COLOR)
        chips.pack(side="left", fill="x", expand=True)

        chip_items = list(tabs or [])
        if title and not tabs:
            chip_items = [(title, True)]

        for label, is_active in chip_items:
            bg = BUTTON_RED if is_active else BUTTON_BLUE
            tk.Label(
                chips,
                text=label,
                bg=bg,
                fg="white",
                font=("Poppins", 10, "bold"),
                padx=10,
                pady=3,
            ).pack(side="left", padx=(0, 6))

        if show_controls:
            controls = tk.Frame(header_row, bg=PANEL_BG_COLOR)
            controls.pack(side="right")
            for sym, cmd in (("+", on_add), ("-", on_remove)):
                btn = tk.Label(
                    controls,
                    text=sym,
                    bg=BUTTON_RED,
                    fg="white",
                    font=("Poppins", 8, "bold"),
                    width=2,
                    padx=0,
                    pady=1,
                    cursor="hand2" if cmd else "",
                )
                if cmd:
                    btn.bind("<Button-1>", lambda e, c=cmd: c())
                btn.pack(side="left", padx=(0, 4))

        tk.Frame(inner, bg=BUTTON_RED, height=2).pack(fill="x", pady=(0, 8))

    return inner


def row_two_columns(parent, left_label, left_widget, right_label, right_widget):
    """Create a two-column row with labels and widgets."""
    row = tk.Frame(parent, bg=PANEL_BG_COLOR)
    row.pack(fill="x", pady=(0, 4))

    left = tk.Frame(row, bg=PANEL_BG_COLOR)
    left.pack(side="left", fill="x", expand=True, padx=(0, 8))
    tk.Label(
        left,
        text=left_label,
        font=LABEL_FONT,
        fg=TEXT_COLOR,
        bg=PANEL_BG_COLOR,
    ).pack(anchor="w")
    left_widget(left)

    right = tk.Frame(row, bg=PANEL_BG_COLOR)
    right.pack(side="left", fill="x", expand=True, padx=(8, 0))
    tk.Label(
        right,
        text=right_label,
        font=LABEL_FONT,
        fg=TEXT_COLOR,
        bg=PANEL_BG_COLOR,
    ).pack(anchor="w")
    right_widget(right)


def styled_entry(parent, readonly=False, width=None):
    """Create a styled entry; returns the Entry widget."""
    bg = READONLY_BG_COLOR if readonly else INPUT_BG_COLOR
    if CTK_AVAILABLE:
        if width is None:
            pixel_width = 220
        else:
            try:
                pixel_width = max(80, int(width) * 8)
            except Exception:
                pixel_width = 220
        entry = ctk.CTkEntry(
            parent,
            width=pixel_width,
            height=28,
            fg_color=bg,
            text_color=TEXT_COLOR,
            border_color="#9EC7CF",
            corner_radius=14,
            font=ENTRY_FONT,
        )
        if readonly:
            entry.configure(state="readonly")
        return entry

    entry = tk.Entry(
        parent,
        font=ENTRY_FONT,
        width=width or 30,
        bg=bg,
        fg=TEXT_COLOR,
        readonlybackground=READONLY_BG_COLOR,
        disabledforeground=TEXT_COLOR,
        insertbackground=TEXT_COLOR,
        relief="flat",
        highlightthickness=1,
        highlightbackground="#9EC7CF",
        highlightcolor="#9EC7CF",
        bd=0,
    )
    if readonly:
        entry.config(state="readonly")
    return entry


def styled_label(parent, text):
    """Create a styled label."""
    return tk.Label(
        parent,
        text=text,
        font=LABEL_FONT,
        fg=TEXT_COLOR,
        bg=PANEL_BG_COLOR,
    )


def action_button(parent, text, variant="primary", command=None):
    """Create a styled button (primary/secondary/danger)."""
    if variant == "secondary":
        bg = BUTTON_GRAY
        hover = BUTTON_GRAY
    elif variant == "danger":
        bg = BUTTON_RED
        hover = BUTTON_RED
    elif variant == "accent":
        bg = BUTTON_ORANGE
        hover = BUTTON_ORANGE
    elif variant in ("info", "blue"):
        bg = BUTTON_BLUE
        hover = BUTTON_BLUE
    else:
        bg = BUTTON_GREEN
        hover = BUTTON_GREEN_HOVER
    if CTK_AVAILABLE:
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            fg_color=bg,
            hover_color=hover,
            text_color="white",
            font=BUTTON_FONT,
            corner_radius=14,
            height=30,
        )
    return tk.Button(
        parent,
        text=text,
        command=command,
        bg=bg,
        fg="white",
        font=BUTTON_FONT,
        relief="flat",
        padx=10,
        pady=4,
    )


def muted_label(parent, text):
    """Create a muted helper label."""
    return tk.Label(
        parent,
        text=text,
        font=("Poppins", 9),
        fg=MUTED_TEXT_COLOR,
        bg=PANEL_BG_COLOR,
    )
