"""
Home page view for the main content area.
"""

from datetime import datetime
import threading

import customtkinter as ctk

from config import (
    ACCENT_TEXT_COLOR,
    BUTTON_BLUE,
    BUTTON_GREEN,
    BUTTON_GREEN_HOVER,
    BUTTON_ORANGE,
    BUTTON_RED,
    CARD_BG_COLOR,
    CARD_HOVER_BG_COLOR,
    INPUT_BG_COLOR,
    MAIN_BG_COLOR,
    MUTED_TEXT_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
)
from utils.excel_handler import (
    load_all_clients,
    load_all_collective_expense_quotations,
    load_all_hotel_quotations,
    load_all_invoices,
)


class HomePage:
    """Interactive home page with quick actions and dynamic highlights."""

    def __init__(self, parent, navigate_callback=None):
        self.parent = parent
        self.navigate_callback = navigate_callback
        self._tips = [
            "Astuce: commencez par creer un client avant de lancer la cotation hotel.",
            "Astuce: en cotation hotel, ajoutez chaque ville/hotel avec le bouton 'Ajouter cet hotel'.",
            "Astuce: le resume des cotations hotel affiche maintenant un tableau simple des grandes lignes.",
            "Astuce: le bouton Parametre est disponible dans la fenetre Transport.",
        ]
        self._tip_index = 0
        self.clock_label = None
        self.tip_label = None
        self.dashboard_stats = self._empty_dashboard_stats()
        self._dashboard_value_labels = {}
        self._pending_dashboard_stats = None
        self._build_ui()
        self._load_dashboard_stats_async()
        self._poll_dashboard_stats()
        self._start_clock()
        self._start_tip_rotation()

    def _empty_dashboard_stats(self):
        return {
            "clients": 0,
            "collective_count": 0,
            "collective_total": 0.0,
            "quotes_count": 0,
            "quotes_total": 0.0,
            "invoices_count": 0,
            "invoices_total": 0.0,
        }

    def _load_dashboard_stats_async(self):
        """Load dashboard stats in background to keep UI responsive."""
        def worker():
            stats = self._load_dashboard_stats()
            self._pending_dashboard_stats = stats

        threading.Thread(target=worker, daemon=True).start()

    def _poll_dashboard_stats(self):
        """Apply dashboard stats from background thread when ready."""
        if not self.parent.winfo_exists():
            return
        if self._pending_dashboard_stats is not None:
            stats = self._pending_dashboard_stats
            self._pending_dashboard_stats = None
            self._apply_dashboard_stats(stats)
            return
        self.parent.after(100, self._poll_dashboard_stats)

    def _apply_dashboard_stats(self, stats):
        self.dashboard_stats = stats
        if not self._dashboard_value_labels:
            return
        self._dashboard_value_labels["clients"].configure(
            text=f"{stats['clients']}"
        )
        self._dashboard_value_labels["collective"].configure(
            text=f"{stats['collective_count']} | {stats['collective_total']:,.0f} MGA"
        )
        self._dashboard_value_labels["quotes"].configure(
            text=f"{stats['quotes_count']} | {stats['quotes_total']:,.0f} MGA"
        )
        self._dashboard_value_labels["invoices"].configure(
            text=f"{stats['invoices_count']} | {stats['invoices_total']:,.0f} MGA"
        )

    def _to_number(self, value):
        try:
            if value is None or value == "":
                return 0.0
            if isinstance(value, (int, float)):
                return float(value)
            text = str(value).strip().replace(" ", "").replace(",", ".")
            cleaned = ""
            dot_seen = False
            for ch in text:
                if ch.isdigit() or ch == "-":
                    cleaned += ch
                elif ch == "." and not dot_seen:
                    cleaned += ch
                    dot_seen = True
            return float(cleaned) if cleaned not in ("", "-", ".") else 0.0
        except Exception:
            return 0.0

    def _load_dashboard_stats(self):
        """Load small dashboard counters and totals for home page."""
        stats = {
            "clients": 0,
            "collective_count": 0,
            "collective_total": 0.0,
            "quotes_count": 0,
            "quotes_total": 0.0,
            "invoices_count": 0,
            "invoices_total": 0.0,
        }
        try:
            clients = load_all_clients()
            stats["clients"] = len(clients)
        except Exception:
            pass
        try:
            collective = load_all_collective_expense_quotations()
            stats["collective_count"] = len(collective)
            stats["collective_total"] = sum(
                self._to_number(item.get("Total_Devise", 0)) for item in collective
            )
        except Exception:
            pass
        try:
            quotes = load_all_hotel_quotations()
            stats["quotes_count"] = len(quotes)
            stats["quotes_total"] = sum(
                self._to_number(item.get("total_price", 0)) for item in quotes
            )
        except Exception:
            pass
        try:
            invoices = load_all_invoices()
            stats["invoices_count"] = len(invoices)
            stats["invoices_total"] = sum(
                self._to_number(item.get("Total_TTC", 0)) for item in invoices
            )
        except Exception:
            pass
        return stats

    def _build_ui(self):
        shell = ctk.CTkFrame(self.parent, fg_color="transparent")
        shell.pack(fill="both", expand=True, padx=24, pady=24)

        hero = ctk.CTkFrame(
            shell, fg_color=PANEL_BG_COLOR, corner_radius=18, border_width=1, border_color="#C9DDE3"
        )
        hero.pack(fill="x", pady=(0, 18))
        hero.grid_columnconfigure(0, weight=1)
        hero.grid_columnconfigure(1, weight=0)

        ctk.CTkLabel(
            hero,
            text="Bienvenue sur Lahimena Tours",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color=TEXT_COLOR,
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 6))

        ctk.CTkLabel(
            hero,
            text="Un espace clair pour vos clients, hotels et devis.",
            font=ctk.CTkFont(size=15),
            text_color=MUTED_TEXT_COLOR,
        ).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 20))

        clock_box = ctk.CTkFrame(hero, fg_color=CARD_BG_COLOR, corner_radius=14)
        clock_box.grid(row=0, column=1, rowspan=2, sticky="e", padx=18, pady=14)

        ctk.CTkLabel(
            clock_box,
            text="Heure locale",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=MUTED_TEXT_COLOR,
        ).pack(padx=14, pady=(10, 4))

        self.clock_label = ctk.CTkLabel(
            clock_box,
            text="--/--/----\n--:--:--",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=TEXT_COLOR,
        )
        self.clock_label.pack(padx=14, pady=(0, 10))

        quick_actions = ctk.CTkFrame(shell, fg_color="transparent")
        quick_actions.pack(fill="x", pady=(0, 14))

        self._add_quick_action(
            quick_actions,
            "Demande client",
            "client_page",
            BUTTON_GREEN,
            BUTTON_GREEN_HOVER,
        )
        self._add_quick_action(
            quick_actions,
            "Cotation hotel multi-villes",
            "hotel_quotation_page",
            BUTTON_BLUE,
            BUTTON_GREEN_HOVER,
        )
        self._add_quick_action(
            quick_actions,
            "Transport + Parametre",
            "transport_page",
            BUTTON_GREEN,
            BUTTON_GREEN_HOVER,
        )
        self._add_quick_action(
            quick_actions,
            "Frais collectifs",
            "collective_expense_page",
            BUTTON_ORANGE,
            "#D48806",
        )
        self._add_quick_action(
            quick_actions,
            "Devis clients",
            "client_quotes_page",
            BUTTON_BLUE,
            BUTTON_GREEN_HOVER,
        )
        self._add_quick_action(
            quick_actions,
            "Factures clients",
            "current_invoices",
            BUTTON_RED,
            "#B71C1C",
        )

        cards_container = ctk.CTkFrame(shell, fg_color="transparent")
        cards_container.pack(fill="both", expand=True)

        self._add_dashboard(cards_container)
        self._add_card(
            cards_container,
            "Demarrage rapide",
            "1) Creez/selectionnez un client\n2) Choisissez une ville de l'itineraire\n3) Calculez puis ajoutez l'hotel dans la liste evolutive\n4) Generez le devis final",
        )

        tip_card = self._add_card(
            cards_container, "Astuce du moment", self._tips[self._tip_index]
        )
        self.tip_label = tip_card["body_label"]
        tip_card["frame"].configure(fg_color=PANEL_BG_COLOR)
        self.tip_label.configure(text_color=TEXT_COLOR)

    def _add_card(self, parent, title, body):
        normal = CARD_BG_COLOR
        hover = CARD_HOVER_BG_COLOR

        card = ctk.CTkFrame(
            parent,
            fg_color=normal,
            corner_radius=14,
            border_width=1,
            border_color="#C9DDE3",
        )
        card.pack(fill="x", pady=8)

        title_label = ctk.CTkLabel(
            card,
            text=title,
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color=TEXT_COLOR,
        )
        title_label.pack(anchor="w", padx=16, pady=(12, 4))

        body_label = ctk.CTkLabel(
            card,
            text=body,
            justify="left",
            anchor="w",
            font=ctk.CTkFont(size=14),
            text_color=MUTED_TEXT_COLOR,
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

    def _add_dashboard(self, parent):
        """Small synthetic dashboard with key counters."""
        wrapper = ctk.CTkFrame(
            parent,
            fg_color=PANEL_BG_COLOR,
            corner_radius=14,
            border_width=1,
            border_color="#C9DDE3",
        )
        wrapper.pack(fill="x", pady=8)

        ctk.CTkLabel(
            wrapper,
            text="Dashboard synthetique",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=16, pady=(12, 10))

        grid = ctk.CTkFrame(wrapper, fg_color="transparent")
        grid.pack(fill="x", padx=12, pady=(0, 12))
        for col in range(4):
            grid.grid_columnconfigure(col, weight=1)

        cards = [
            ("Clients", "clients", BUTTON_BLUE),
            ("Frais collectifs", "collective", BUTTON_ORANGE),
            ("Devis hotel", "quotes", BUTTON_GREEN),
            ("Factures clients", "invoices", BUTTON_RED),
        ]

        for idx, (title, key, color) in enumerate(cards):
            card = ctk.CTkFrame(
                grid,
                fg_color=CARD_BG_COLOR,
                corner_radius=12,
                border_width=1,
                border_color="#D3E2E7",
            )
            card.grid(row=0, column=idx, sticky="nsew", padx=6, pady=4)
            ctk.CTkLabel(
                card,
                text=title,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=MUTED_TEXT_COLOR,
            ).pack(anchor="w", padx=10, pady=(10, 2))
            value_label = ctk.CTkLabel(
                card,
                text="--",
                font=ctk.CTkFont(size=15, weight="bold"),
                text_color=color,
            )
            value_label.pack(anchor="w", padx=10, pady=(0, 10))
            self._dashboard_value_labels[key] = value_label

    def _add_quick_action(self, parent, text, route, color, hover):
        btn = ctk.CTkButton(
            parent,
            text=text,
            command=lambda: self._navigate(route),
            fg_color=color,
            hover_color=hover,
            corner_radius=12,
            height=42,
            text_color="white",
        )
        btn.pack(side="left", padx=(0, 10), pady=4)

    def _navigate(self, route):
        if self.navigate_callback:
            self.navigate_callback(route)

    def _start_clock(self):
        if not self.clock_label or not self.clock_label.winfo_exists():
            return
        self.clock_label.configure(text=datetime.now().strftime("%d/%m/%Y\n%H:%M:%S"))
        self.parent.after(1000, self._start_clock)

    def _start_tip_rotation(self):
        if not self.tip_label or not self.tip_label.winfo_exists():
            return
        self._tip_index = (self._tip_index + 1) % len(self._tips)
        self.tip_label.configure(text=self._tips[self._tip_index])
        self.parent.after(5000, self._start_tip_rotation)
