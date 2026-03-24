"""
Page formulaire client (sans liste).
"""

import tkinter as tk

from config import MAIN_BG_COLOR
from gui.forms.client_form import ClientForm


class ClientPage:
    """Display client form only."""

    def __init__(self, parent, client_to_edit=None):
        self.parent = parent
        self.form_container = None
        self._create_layout()
        self._show_form(client_to_edit=client_to_edit)

    def _create_layout(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        self.form_container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.form_container.pack(fill="both", expand=True, padx=0, pady=0)

    def _clear_container(self, container):
        for widget in container.winfo_children():
            widget.destroy()

    def _show_form(self, client_to_edit=None):
        self._clear_container(self.form_container)
        ClientForm(
            self.form_container,
            client_to_edit=client_to_edit,
            on_save_callback=self._on_client_saved,
        )

    def _on_client_saved(self):
        self._show_form()

    def _on_edit_requested(self, client_data):
        self._show_form(client_to_edit=client_data)
