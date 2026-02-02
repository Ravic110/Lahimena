"""
Main content GUI component
"""

import customtkinter as ctk
from config import *


class MainContent:
    """
    Main content area component
    """

    def __init__(self, parent):
        """
        Initialize main content area

        Args:
            parent: Parent widget
        """
        self.parent = parent

        # Create scrollable main content
        self.main_scroll = ctk.CTkScrollableFrame(
            parent,
            corner_radius=0
        )
        self.main_scroll.grid(row=0, column=1, sticky="nswe", padx=(0, 0))

        self._show_welcome()

    def update_content(self, content_type):
        """
        Update main content based on selected menu item

        Args:
            content_type (str): Type of content to display
        """
        # Clear current content
        for widget in self.main_scroll.winfo_children():
            widget.destroy()

        # Show appropriate content
        if content_type == "client_form":
            self._show_client_form()
        elif content_type == "welcome":
            self._show_welcome()
        else:
            self._show_placeholder(content_type)

    def _show_welcome(self):
        """Show welcome message"""
        title = ctk.CTkLabel(
            self.main_scroll,
            text="Bienvenue - Lahimena Tours: Gestion de devis",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title.pack(pady=40)

    def _show_client_form(self):
        """Show client form"""
        from gui.forms.client_form import ClientForm
        ClientForm(self.main_scroll)

    def _show_placeholder(self, content_type):
        """Show placeholder for unimplemented features"""
        title = ctk.CTkLabel(
            self.main_scroll,
            text=f"Fonction '{content_type}' - À implémenter",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.pack(pady=40)