"""
Combined page for collective expense quotation form and summary.
"""

import tkinter as tk

from config import BUTTON_BLUE, BUTTON_FONT, MAIN_BG_COLOR
from gui.forms.collective_expense_quotation import CollectiveExpenseQuotation
from gui.forms.collective_expense_quotation_summary import (
    CollectiveExpenseQuotationSummary,
)


class CollectiveExpensePage:
    """
    Display collective expense quotation form and summary on the same page.
    """

    def __init__(self, parent, on_back_to_hub=None):
        self.parent = parent
        self.on_back_to_hub = on_back_to_hub
        self.form_container = None
        self.summary_container = None

        self._create_layout()
        self._show_form()
        self._show_summary()

    def _create_layout(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        if self.on_back_to_hub:
            top_actions = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
            top_actions.pack(fill="x", padx=20, pady=(4, 2))
            tk.Button(
                top_actions,
                text="⬅ Retour vers Factures / Devis",
                command=self._go_back_to_hub,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=8,
                pady=3,
            ).pack(side="left")

        self.form_container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.form_container.pack(fill="x", padx=0, pady=(0, 10))

        self.summary_container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.summary_container.pack(fill="both", expand=True, padx=0, pady=(0, 0))

    def _clear_container(self, container):
        for widget in container.winfo_children():
            widget.destroy()

    def _show_form(self, edit_data=None, row_number=None):
        self._clear_container(self.form_container)
        CollectiveExpenseQuotation(
            self.form_container,
            edit_data=edit_data,
            row_number=row_number,
            callback_on_save=self._on_form_saved,
        )

    def _show_summary(self):
        self._clear_container(self.summary_container)
        CollectiveExpenseQuotationSummary(
            self.summary_container,
            callback_edit=self._on_edit_requested,
            callback_add=self._on_add_requested,
        )

    def _on_form_saved(self):
        self._show_form()
        self._show_summary()

    def _on_edit_requested(self, data, row_number):
        self._show_form(edit_data=data, row_number=row_number)

    def _on_add_requested(self):
        self._show_form()

    def _go_back_to_hub(self):
        if self.on_back_to_hub:
            self.on_back_to_hub()
