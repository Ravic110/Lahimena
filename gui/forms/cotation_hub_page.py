"""
Dedicated quotation hub page.
"""

import customtkinter as ctk


class CotationHubPage:
    """Landing page that groups quotation submodules."""

    def __init__(self, parent, navigate_callback=None):
        self.parent = parent
        self.navigate_callback = navigate_callback
        self._build_ui()

    def _build_ui(self):
        shell = ctk.CTkFrame(self.parent, fg_color="transparent")
        shell.pack(fill="both", expand=True, padx=24, pady=24)

        hero = ctk.CTkFrame(shell, fg_color="#0F172A", corner_radius=18)
        hero.pack(fill="x", pady=(0, 18))

        ctk.CTkLabel(
            hero,
            text="Cotation",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color="#F8FAFC",
        ).pack(anchor="w", padx=24, pady=(20, 6))

        ctk.CTkLabel(
            hero,
            text="Accedez rapidement aux modules de cotation.",
            font=ctk.CTkFont(size=15),
            text_color="#CBD5E1",
        ).pack(anchor="w", padx=24, pady=(0, 20))

        grid = ctk.CTkFrame(shell, fg_color="transparent")
        grid.pack(fill="both", expand=True)
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=1)

        self._add_group(
            grid,
            row=0,
            col=0,
            title="Cotation hotel",
            fg_color="#1E3A8A",
            action=("Ouvrir", "hotel_quotation_page"),
        )
        self._add_group(
            grid,
            row=0,
            col=1,
            title="Transport",
            fg_color="#0F766E",
            action=("Ouvrir", "transport_page"),
        )
        self._add_group(
            grid,
            row=1,
            col=0,
            title="Billets avion",
            fg_color="#6D28D9",
            action=("Ouvrir", "air_ticket_page"),
        )

    def _add_group(self, parent, row, col, title, fg_color, action):
        card = ctk.CTkFrame(parent, fg_color=fg_color, corner_radius=14)
        card.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)

        ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#F8FAFC",
        ).pack(anchor="w", padx=16, pady=(16, 10))

        text, route = action
        ctk.CTkButton(
            card,
            text=text,
            command=lambda r=route: self._navigate(r),
            height=40,
            corner_radius=10,
            fg_color="#0B1220",
            hover_color="#111827",
        ).pack(fill="x", padx=16, pady=(0, 14))

    def _navigate(self, route):
        if self.navigate_callback:
            self.navigate_callback(route)
