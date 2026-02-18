"""
Home page view for the main content area.
"""

import customtkinter as ctk
from datetime import datetime


class HomePage:
    """Interactive home page with quick actions and dynamic highlights."""

    def __init__(self, parent, navigate_callback=None):
        self.parent = parent
        self.navigate_callback = navigate_callback
        self._tips = [
            "Astuce: commencez par creer un client avant la cotation.",
            "Astuce: verifiez la liste hotels pour reutiliser une fiche existante.",
            "Astuce: utilisez le resume des cotations pour un suivi rapide."
        ]
        self._tip_index = 0
        self.clock_label = None
        self.tip_label = None
        self._build_ui()
        self._start_clock()
        self._start_tip_rotation()

    def _build_ui(self):
        shell = ctk.CTkFrame(self.parent, fg_color="transparent")
        shell.pack(fill="both", expand=True, padx=24, pady=24)

        hero = ctk.CTkFrame(shell, fg_color="#0F172A", corner_radius=18)
        hero.pack(fill="x", pady=(0, 18))
        hero.grid_columnconfigure(0, weight=1)
        hero.grid_columnconfigure(1, weight=0)

        ctk.CTkLabel(
            hero,
            text="Bienvenue sur Lahimena Tours",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color="#F8FAFC"
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 6))

        ctk.CTkLabel(
            hero,
            text="Un espace clair pour vos clients, hotels et devis.",
            font=ctk.CTkFont(size=15),
            text_color="#CBD5E1"
        ).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 20))

        clock_box = ctk.CTkFrame(hero, fg_color="#1E293B", corner_radius=14)
        clock_box.grid(row=0, column=1, rowspan=2, sticky="e", padx=18, pady=14)

        ctk.CTkLabel(
            clock_box,
            text="Heure locale",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#93C5FD"
        ).pack(padx=14, pady=(10, 4))

        self.clock_label = ctk.CTkLabel(
            clock_box,
            text="--/--/----\n--:--:--",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#F8FAFC"
        )
        self.clock_label.pack(padx=14, pady=(0, 10))

        quick_actions = ctk.CTkFrame(shell, fg_color="transparent")
        quick_actions.pack(fill="x", pady=(0, 14))

        self._add_quick_action(
            quick_actions,
            "Nouveau client",
            "client_form",
            "#059669",
            "#047857"
        )
        self._add_quick_action(
            quick_actions,
            "Nouvelle cotation",
            "hotel_quotation",
            "#0284C7",
            "#0369A1"
        )
        self._add_quick_action(
            quick_actions,
            "Liste hotels",
            "hotel_list",
            "#6D28D9",
            "#5B21B6"
        )

        cards_container = ctk.CTkFrame(shell, fg_color="transparent")
        cards_container.pack(fill="both", expand=True)

        self._add_card(
            cards_container,
            "Demarrage rapide",
            "1) Creez un client\n2) Ajoutez ou choisissez un hotel\n3) Lancez une cotation"
        )
        self._add_card(
            cards_container,
            "Raccourcis utiles",
            "Client, cotation hotel, resume de cotations et liste hotels."
        )

        tip_card = self._add_card(
            cards_container,
            "Astuce du moment",
            self._tips[self._tip_index]
        )
        self.tip_label = tip_card["body_label"]
        tip_card["frame"].configure(fg_color="#172554")
        self.tip_label.configure(text_color="#DBEAFE")

    def _add_card(self, parent, title, body):
        normal = "#1F2937"
        hover = "#334155"

        card = ctk.CTkFrame(parent, fg_color=normal, corner_radius=14)
        card.pack(fill="x", pady=8)

        title_label = ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color="#F8FAFC"
        )
        title_label.pack(anchor="w", padx=16, pady=(12, 4))

        body_label = ctk.CTkLabel(
            card,
            text=body,
            justify="left",
            anchor="w",
            font=ctk.CTkFont(size=14),
            text_color="#E2E8F0"
        )
        body_label.pack(anchor="w", padx=16, pady=(0, 12))

        def on_enter(_event):
            card.configure(fg_color=hover)

        def on_leave(_event):
            card.configure(fg_color=normal)

        for widget in (card, title_label, body_label):
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)

        return {"frame": card, "body_label": body_label}

    def _add_quick_action(self, parent, text, route, color, hover):
        btn = ctk.CTkButton(
            parent,
            text=text,
            command=lambda: self._navigate(route),
            fg_color=color,
            hover_color=hover,
            corner_radius=12,
            height=42
        )
        btn.pack(side="left", padx=(0, 10), pady=4)

    def _navigate(self, route):
        if self.navigate_callback:
            self.navigate_callback(route)

    def _start_clock(self):
        if not self.clock_label or not self.clock_label.winfo_exists():
            return
        self.clock_label.configure(
            text=datetime.now().strftime("%d/%m/%Y\n%H:%M:%S")
        )
        self.parent.after(1000, self._start_clock)

    def _start_tip_rotation(self):
        if not self.tip_label or not self.tip_label.winfo_exists():
            return
        self._tip_index = (self._tip_index + 1) % len(self._tips)
        self.tip_label.configure(text=self._tips[self._tip_index])
        self.parent.after(5000, self._start_tip_rotation)
