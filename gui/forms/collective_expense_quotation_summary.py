"""
Placeholder collective expense quotation summary
"""

import tkinter as tk

from config import MAIN_BG_COLOR, TEXT_COLOR, TITLE_FONT


class CollectiveExpenseQuotationSummary:
    """Placeholder for collective expense quotation summary"""

    def __init__(self, parent, callback_edit=None, callback_add=None):
        for widget in parent.winfo_children():
            widget.destroy()

        title = tk.Label(
            parent,
            text="RÉSUMÉ FRAIS COLLECTIF (À implémenter)",
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
