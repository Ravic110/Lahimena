"""
Combined page for hotel quotation form and summary.
"""

import tkinter as tk

from config import MAIN_BG_COLOR
from gui.forms.hotel_quotation import HotelQuotation
from gui.forms.hotel_quotation_summary import HotelQuotationSummary


class HotelQuotationPage:
    """Display hotel quotation and summary on the same page."""

    def __init__(self, parent):
        self.parent = parent
        self.form_container = None
        self.summary_container = None

        self._create_layout()
        self._show_form()
        self._show_summary()

    def _create_layout(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        self.form_container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.form_container.pack(fill="x", padx=0, pady=(0, 10))

        self.summary_container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.summary_container.pack(fill="both", expand=True, padx=0, pady=(0, 0))

    def _clear_container(self, container):
        for widget in container.winfo_children():
            widget.destroy()

    def _show_form(self):
        self._clear_container(self.form_container)
        HotelQuotation(self.form_container)

    def _show_summary(self):
        self._clear_container(self.summary_container)
        HotelQuotationSummary(self.summary_container)

