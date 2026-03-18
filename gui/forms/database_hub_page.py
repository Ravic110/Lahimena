"""
Dedicated database hub page.
"""

import customtkinter as ctk

from config import (
    BUTTON_GREEN,
    BUTTON_GREEN_HOVER,
    CARD_BG_COLOR,
    MUTED_TEXT_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
)


class DatabaseHubPage:
    """Landing page that groups all database submenus in one place."""

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
            text="Bases de données",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=24, pady=(20, 6))

        ctk.CTkLabel(
            hero,
            text="Toutes les gestions de reference sont centralisees ici.",
            font=ctk.CTkFont(size=15),
            text_color=MUTED_TEXT_COLOR,
        ).pack(anchor="w", padx=24, pady=(0, 20))

        grid = ctk.CTkFrame(shell, fg_color="transparent")
        grid.pack(fill="both", expand=True)
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=1)

        self._add_group(
            grid,
            row=0,
            col=0,
            title="Hotels",
            fg_color=CARD_BG_COLOR,
            actions=[
                ("Liste hotels", "hotel_list"),
                ("Ajouter hotel", "hotel_form"),
            ],
        )
        self._add_group(
            grid,
            row=0,
            col=1,
            title="Frais collectifs",
            fg_color=CARD_BG_COLOR,
            actions=[
                ("Liste frais collectifs", "collective_expense_db_list"),
                ("Ajouter frais collectif", "collective_expense_db_form"),
            ],
        )
        self._add_group(
            grid,
            row=1,
            col=0,
            title="Circuits et transport",
            fg_color=CARD_BG_COLOR,
            actions=[
                ("Base circuits", "circuit_db_page"),
                ("Base transport", "transport_db_page"),
            ],
        )
        self._add_group(
            grid,
            row=1,
            col=1,
            title="Billets avion",
            fg_color=CARD_BG_COLOR,
            actions=[
                ("Liste billets avion", "air_ticket_db_list"),
                ("Ajouter billet avion", "air_ticket_db_form"),
            ],
        )

    def _add_group(self, parent, row, col, title, fg_color, actions):
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

        for text, route in actions:
            ctk.CTkButton(
                card,
                text=text,
                command=lambda r=route: self._navigate(r),
                height=40,
                corner_radius=10,
                fg_color=BUTTON_GREEN,
                hover_color=BUTTON_GREEN_HOVER,
                text_color="white",
            ).pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkLabel(
            card,
            text="",
            height=8,
        ).pack()

    def _navigate(self, route):
        if self.navigate_callback:
            self.navigate_callback(route)
