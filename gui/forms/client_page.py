"""
Combined page for client form and client list.
"""

import tkinter as tk

from config import MAIN_BG_COLOR
from gui.forms.client_form import ClientForm
from gui.forms.client_list import ClientList


class ClientPage:
    """Display client form and client list on the same page."""

    def __init__(self, parent):
        self.parent = parent
        self.form_container = None
        self.list_container = None

        self._create_layout()
        self._show_form()
        self._show_list()

    def _create_layout(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        self.form_container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.form_container.pack(fill="x", padx=0, pady=(0, 10))

        self.list_container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.list_container.pack(fill="both", expand=True, padx=0, pady=(0, 0))

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

    def _show_list(self):
        self._clear_container(self.list_container)
        ClientList(
            self.list_container,
            on_edit_client=self._on_edit_requested,
            on_new_client=self._on_new_requested,
        )

    def _on_client_saved(self):
        self._show_form()
        self._show_list()

    def _on_edit_requested(self, client_data):
        self._show_form(client_to_edit=client_data)

    def _on_new_requested(self):
        self._show_form()

