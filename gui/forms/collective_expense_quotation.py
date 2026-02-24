"""
Placeholder collective expense quotation form
"""

import tkinter as tk

from config import MAIN_BG_COLOR, TEXT_COLOR, TITLE_FONT


class CollectiveExpenseQuotation:
    """Placeholder for collective expense quotation"""

    def __init__(self, parent, edit_data=None, row_number=None, callback_on_save=None):
        for widget in parent.winfo_children():
            widget.destroy()

        title = tk.Label(
            parent,
            text="COTATION FRAIS COLLECTIF (À implémenter)",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        title.pack(pady=40)

        note = tk.Label(
            parent,
            text="Cette section n'est pas encore disponible.",
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        note.pack()
