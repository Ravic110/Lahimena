"""
Combined page for client quotation form and quotation history.
"""

import tkinter as tk

from config import MAIN_BG_COLOR
from gui.forms.client_quotation import ClientQuotation
from gui.forms.client_quotation_history import ClientQuotationHistory


class ClientQuotationPage:
    """Display client quotation and history on the same page."""

    def __init__(self, parent):
        self.parent = parent
        self.form_container = None
        self.history_container = None

        self._create_layout()
        self._show_form()
        self._show_history()

    def _create_layout(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        self.form_container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.form_container.pack(fill="x", padx=0, pady=(0, 10))

        self.history_container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.history_container.pack(fill="both", expand=True, padx=0, pady=(0, 0))

    def _clear_container(self, container):
        for widget in container.winfo_children():
            widget.destroy()

    def _show_form(self):
        self._clear_container(self.form_container)
        ClientQuotation(self.form_container)

    def _show_history(self):
        self._clear_container(self.history_container)
        ClientQuotationHistory(self.history_container)

