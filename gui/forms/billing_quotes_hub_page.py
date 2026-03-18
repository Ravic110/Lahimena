"""
Dedicated hub page for quotes and invoices.
"""

import customtkinter as ctk

from config import (
    BUTTON_BLUE,
    BUTTON_GREEN,
    BUTTON_GREEN_HOVER,
    CARD_BG_COLOR,
    MUTED_TEXT_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
)


class BillingQuotesHubPage:
    """Landing page grouping client quotes and invoices."""

    def __init__(self, parent, navigate_callback=None):
        self.parent = parent
        self.navigate_callback = navigate_callback
        self._build_ui()

    def _build_ui(self):
        shell = ctk.CTkFrame(self.parent, fg_color="transparent")
        shell.pack(fill="both", expand=True, padx=24, pady=24)

        hero = ctk.CTkFrame(
            shell, fg_color=PANEL_BG_COLOR, corner_radius=18, border_width=1, border_color="#C9DDE3"
        )
        hero.pack(fill="x", pady=(0, 18))

        ctk.CTkLabel(
            hero,
            text="Factures / Devis",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=24, pady=(20, 6))

        ctk.CTkLabel(
            hero,
            text="Gestion centralisee des devis, factures et depenses.",
            font=ctk.CTkFont(size=15),
            text_color=MUTED_TEXT_COLOR,
        ).pack(anchor="w", padx=24, pady=(0, 14))

        ctk.CTkButton(
            hero,
            text="⬅ Retour Accueil",
            command=lambda: self._navigate("home"),
            height=36,
            corner_radius=10,
            fg_color=BUTTON_BLUE,
            hover_color=BUTTON_GREEN_HOVER,
            text_color="white",
        ).pack(anchor="w", padx=24, pady=(0, 16))

        grid = ctk.CTkFrame(shell, fg_color="transparent")
        grid.pack(fill="both", expand=True)
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=1)

        self._add_group(
            grid,
            row=0,
            col=0,
            title="Devis clients",
            fg_color=CARD_BG_COLOR,
            action=("Ouvrir", "client_quotes_page"),
        )
        self._add_group(
            grid,
            row=0,
            col=1,
            title="Facture clients",
            fg_color=CARD_BG_COLOR,
            action=("Ouvrir", "current_invoices"),
        )
        self._add_group(
            grid,
            row=1,
            col=0,
            title="Depenses",
            fg_color=CARD_BG_COLOR,
            action=("Ouvrir", "expenses_page"),
        )

    def _add_group(self, parent, row, col, title, fg_color, action):
        card = ctk.CTkFrame(
            parent, fg_color=fg_color, corner_radius=14, border_width=1, border_color="#C9DDE3"
        )
        card.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)

        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=16, pady=(16, 10))

        text, route = action
        ctk.CTkButton(
            card,
            text=text,
            command=lambda r=route: self._navigate(r),
            height=40,
            corner_radius=10,
            fg_color=BUTTON_GREEN,
            hover_color=BUTTON_GREEN_HOVER,
            text_color="white",
        ).pack(fill="x", padx=16, pady=(0, 14))

    def _navigate(self, route):
        if self.navigate_callback:
            self.navigate_callback(route)
