"""Empty expenses page placeholder."""

import tkinter as tk

from config import BUTTON_BLUE, BUTTON_FONT, MAIN_BG_COLOR, TEXT_COLOR, TITLE_FONT


class ExpensesPage:
    """Temporary empty page for expenses."""

    def __init__(self, parent, on_back_to_hub=None):
        self.parent = parent
        self.on_back_to_hub = on_back_to_hub
        self._build_ui()

    def _build_ui(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        if self.on_back_to_hub:
            tk.Button(
                container,
                text="⬅ Retour vers Factures / Devis",
                command=self.on_back_to_hub,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=10,
                pady=4,
            ).pack(anchor="w", pady=(0, 16))

        tk.Label(
            container,
            text="DEPENSES",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(anchor="w", pady=(0, 8))

        tk.Label(
            container,
            text="Cette section est vide pour le moment.",
            font=BUTTON_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(anchor="w")
