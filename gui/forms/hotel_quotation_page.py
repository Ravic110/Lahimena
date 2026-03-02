"""
Combined page for hotel quotation form and summary.
"""

import tkinter as tk

from config import BUTTON_BLUE, BUTTON_FONT, MAIN_BG_COLOR
from gui.forms.hotel_quotation import HotelQuotation
from gui.forms.hotel_quotation_summary import HotelQuotationSummary


class HotelQuotationPage:
    """Display hotel quotation and summary on the same page."""

    def __init__(self, parent, on_back_to_cotation=None):
        self.parent = parent
        self.on_back_to_cotation = on_back_to_cotation
        self.form_container = None
        self.summary_container = None

        self._create_layout()
        self._show_form()
        self._show_summary()

    def _create_layout(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        if self.on_back_to_cotation:
            top_actions = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
            top_actions.pack(fill="x", padx=20, pady=(4, 2))
            tk.Button(
                top_actions,
                text="⬅ Retour vers Cotation",
                command=self._go_back_to_cotation,
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

    def _show_form(self):
        self._clear_container(self.form_container)
        HotelQuotation(self.form_container)

    def _show_summary(self):
        self._clear_container(self.summary_container)
        HotelQuotationSummary(self.summary_container)

    def _go_back_to_cotation(self):
        if self.on_back_to_cotation:
            self.on_back_to_cotation()
