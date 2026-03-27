"""
Home page view for the main content area.
"""

import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    BUTTON_BLUE,
    BUTTON_GREEN,
    BUTTON_GREEN_HOVER,
    BUTTON_ORANGE,
    BUTTON_RED,
    CARD_BG_COLOR,
    INPUT_BG_COLOR,
    MAIN_BG_COLOR,
    MUTED_TEXT_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
)

_STATUT_COLORS = {
    "En cours":   "#0097A7",   # cyan foncé
    "Accepté":    "#2E7D32",   # vert foncé
    "En circuit": "#E65100",   # orange foncé
    "Annulé":     "#C62828",   # rouge foncé
}
_STATUT_ROW_BG = {
    "En cours":   "#E0F7FA",   # cyan clair
    "Accepté":    "#E8F5E9",   # vert clair
    "En circuit": "#FFF3E0",   # orange clair
    "Annulé":     "#FFEBEE",   # rouge clair
}
# Emoji badge affiché dans la colonne Statut
_STATUT_BADGE = {
    "En cours":   "🔵  En cours",
    "Accepté":    "🟢  Accepté",
    "En circuit": "🟠  En circuit",
    "Annulé":     "🔴  Annulé",
}


class HomePage:
    """Interactive home page with quick actions, dashboard and client list."""

    def __init__(self, parent, navigate_callback=None):
        self.parent = parent
        self.navigate_callback = navigate_callback
        self.clock_label = None
        self.dashboard_stats = self._empty_dashboard_stats()
        self._dashboard_value_labels = {}
        self._pending_dashboard_stats = None
        # Client list state
        self._all_clients = []
        self._pending_clients = None
        self._client_tree = None
        self._search_var = None
        self._statut_filter_var = None
        self._build_ui()
        self._load_dashboard_stats_async()
        self._poll_dashboard_stats()
        self._load_clients_async()
        self._poll_clients()
        self._start_clock()

    # ── Dashboard async ────────────────────────────────────────────────────

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
        def worker():
            self._pending_dashboard_stats = self._load_dashboard_stats()
        threading.Thread(target=worker, daemon=True).start()

    def _poll_dashboard_stats(self):
        try:
            if not self.parent.winfo_exists():
                return
        except Exception:
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
        from utils.excel_handler import (
            load_all_clients,
            load_all_collective_expense_quotations,
            load_all_hotel_quotations,
            load_all_invoices,
        )
        stats = self._empty_dashboard_stats()
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

    # ── Client list async ──────────────────────────────────────────────────

    def _load_clients_async(self):
        def worker():
            try:
                from utils.excel_handler import load_all_clients
                self._pending_clients = load_all_clients()
            except Exception:
                self._pending_clients = []
        threading.Thread(target=worker, daemon=True).start()

    def _poll_clients(self):
        try:
            if not self.parent.winfo_exists():
                return
        except Exception:
            return
        if self._pending_clients is not None:
            clients = self._pending_clients
            self._pending_clients = None
            self._all_clients = clients
            self._refresh_client_tree()
            return
        self.parent.after(150, self._poll_clients)

    def _refresh_client_tree(self):
        if self._client_tree is None or not self._client_tree.winfo_exists():
            return
        search = (self._search_var.get() if self._search_var else "").lower().strip()
        statut_filter = (self._statut_filter_var.get() if self._statut_filter_var else "Tous")
        for item in self._client_tree.get_children():
            self._client_tree.delete(item)
        for client in self._all_clients:
            if search and not any(
                search in str(client.get(f, "")).lower()
                for f in ("nom", "ref_client", "email", "telephone", "circuit", "numero_dossier")
            ):
                continue
            statut = client.get("statut") or "En cours"
            if statut_filter != "Tous" and statut != statut_filter:
                continue
            values = (
                _STATUT_BADGE.get(statut, statut),
                client.get("numero_dossier", ""),
                client.get("nom", ""),
                client.get("nombre_participants", ""),
                client.get("duree_sejour", ""),
                client.get("date_arrivee", ""),
                client.get("date_depart", ""),
                client.get("restauration", ""),
                client.get("compagnie", ""),
                client.get("heure_arrivee", ""),
                client.get("heure_depart", ""),
            )
            self._client_tree.insert("", "end", values=values, tags=(statut,))

    def _on_search_change(self, *_):
        self._refresh_client_tree()

    def _on_statut_filter_change(self, *_):
        self._refresh_client_tree()

    # ── UI ─────────────────────────────────────────────────────────────────

    def _build_ui(self):
        shell = ctk.CTkFrame(self.parent, fg_color="transparent")
        shell.pack(fill="both", expand=True, padx=24, pady=24)

        # Hero banner
        hero = ctk.CTkFrame(
            shell, fg_color=PANEL_BG_COLOR, corner_radius=18,
            border_width=1, border_color="#C9DDE3",
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
            clock_box, text="Heure locale",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=MUTED_TEXT_COLOR,
        ).pack(padx=14, pady=(10, 4))

        self.clock_label = ctk.CTkLabel(
            clock_box, text="--/--/----\n--:--:--",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=TEXT_COLOR,
        )
        self.clock_label.pack(padx=14, pady=(0, 10))

        # Quick actions
        quick_actions = ctk.CTkFrame(shell, fg_color="transparent")
        quick_actions.pack(fill="x", pady=(0, 14))

        self._add_quick_action(quick_actions, "Demande client", "client_page", BUTTON_GREEN, BUTTON_GREEN_HOVER)
        self._add_quick_action(quick_actions, "Cotation hotel multi-villes", "hotel_quotation_page", BUTTON_BLUE, BUTTON_GREEN_HOVER)
        self._add_quick_action(quick_actions, "Transport + Parametre", "transport_page", BUTTON_GREEN, BUTTON_GREEN_HOVER)
        self._add_quick_action(quick_actions, "Frais collectifs", "collective_expense_page", BUTTON_ORANGE, "#D48806")
        self._add_quick_action(quick_actions, "Devis clients", "client_quotes_page", BUTTON_BLUE, BUTTON_GREEN_HOVER)
        self._add_quick_action(quick_actions, "Factures clients", "current_invoices", BUTTON_RED, "#B71C1C")

        cards_container = ctk.CTkFrame(shell, fg_color="transparent")
        cards_container.pack(fill="both", expand=True)

        self._add_dashboard(cards_container)
        self._add_client_list(cards_container)

    def _add_quick_action(self, parent, text, route, color, hover):
        try:
            from utils.auth_handler import current_role
            is_comptable = current_role() == "comptable"
        except Exception:
            is_comptable = False

        if is_comptable:
            btn_state = "disabled"
            fg = "#AAAAAA"
            hv = "#AAAAAA"
            txt_color = "#DDDDDD"
        else:
            btn_state = "normal"
            fg = color
            hv = hover
            txt_color = "white"

        ctk.CTkButton(
            parent, text=text,
            command=lambda: self._navigate(route),
            fg_color=fg, hover_color=hv,
            corner_radius=12, height=42, text_color=txt_color,
            state=btn_state,
        ).pack(side="left", padx=(0, 10), pady=4)

    def _add_dashboard(self, parent):
        wrapper = ctk.CTkFrame(
            parent, fg_color=PANEL_BG_COLOR,
            corner_radius=14, border_width=1, border_color="#C9DDE3",
        )
        wrapper.pack(fill="x", pady=8)

        ctk.CTkLabel(
            wrapper, text="Dashboard synthetique",
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
                grid, fg_color=CARD_BG_COLOR,
                corner_radius=12, border_width=1, border_color="#D3E2E7",
            )
            card.grid(row=0, column=idx, sticky="nsew", padx=6, pady=4)
            ctk.CTkLabel(
                card, text=title,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=MUTED_TEXT_COLOR,
            ).pack(anchor="w", padx=10, pady=(10, 2))
            value_label = ctk.CTkLabel(
                card, text="--",
                font=ctk.CTkFont(size=15, weight="bold"),
                text_color=color,
            )
            value_label.pack(anchor="w", padx=10, pady=(0, 10))
            self._dashboard_value_labels[key] = value_label

    def _add_client_list(self, parent):
        wrapper = ctk.CTkFrame(
            parent, fg_color=PANEL_BG_COLOR,
            corner_radius=14, border_width=1, border_color="#C9DDE3",
        )
        wrapper.pack(fill="both", expand=True, pady=8)

        # Header row: title + search + refresh
        header = tk.Frame(wrapper, bg=PANEL_BG_COLOR)
        header.pack(fill="x", padx=16, pady=(12, 8))

        tk.Label(
            header, text="Liste des clients",
            font=("Poppins", 17, "bold"),
            fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(side="left")

        # Refresh button
        tk.Button(
            header, text="🔄",
            command=self._reload_clients,
            bg=BUTTON_BLUE, fg="white",
            font=("Poppins", 11), relief="flat",
            padx=8, cursor="hand2",
        ).pack(side="right", padx=(6, 0))

        # Search bar
        self._search_var = tk.StringVar()
        self._search_var.trace("w", self._on_search_change)
        search_entry = tk.Entry(
            header,
            textvariable=self._search_var,
            font=("Poppins", 12),
            bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
            insertbackground=TEXT_COLOR,
            relief="flat", bd=2, width=28,
        )
        search_entry.pack(side="right", padx=(6, 0), ipady=4)
        tk.Label(
            header, text="🔍",
            font=("Poppins", 12),
            fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(side="right")

        # Filtre statut
        _STATUT_OPTIONS = ["Tous", "En cours", "Accepté", "En circuit", "Annulé"]
        self._statut_filter_var = tk.StringVar(value="Tous")
        self._statut_filter_var.trace("w", self._on_statut_filter_change)
        statut_combo = ttk.Combobox(
            header,
            textvariable=self._statut_filter_var,
            values=_STATUT_OPTIONS,
            state="readonly",
            width=12,
            font=("Poppins", 11),
        )
        statut_combo.pack(side="right", padx=(6, 0), ipady=3)
        tk.Label(
            header, text="Statut :",
            font=("Poppins", 11),
            fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(side="right", padx=(12, 0))

        # Hint
        tk.Label(
            wrapper,
            text="Double-cliquez sur un client pour le modifier ou changer son statut.",
            font=("Poppins", 10),
            fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(anchor="w", padx=16, pady=(0, 6))

        # Treeview
        tree_frame = tk.Frame(wrapper, bg=PANEL_BG_COLOR)
        tree_frame.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        cols = (
            "statut", "numero_dossier", "nom", "nombre_participants",
            "duree_sejour", "date_arrivee", "date_depart",
            "restauration", "compagnie", "heure_arrivee", "heure_depart",
        )
        self._client_tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
            height=10,
        )
        vsb.config(command=self._client_tree.yview)
        hsb.config(command=self._client_tree.xview)

        headers = {
            "statut":               "Statut",
            "numero_dossier":       "N° Dossier",
            "nom":                  "Nom clients",
            "nombre_participants":  "Nb pax",
            "duree_sejour":         "Durée",
            "date_arrivee":         "Début",
            "date_depart":          "Fin",
            "restauration":         "Formule",
            "compagnie":            "Compagnie",
            "heure_arrivee":        "H. Arrivée",
            "heure_depart":         "H. Départ",
        }
        widths = {
            "statut": 130, "numero_dossier": 130, "nom": 170,
            "nombre_participants": 60, "duree_sejour": 60,
            "date_arrivee": 90, "date_depart": 90,
            "restauration": 130, "compagnie": 120,
            "heure_arrivee": 80, "heure_depart": 80,
        }
        for c in cols:
            self._client_tree.heading(c, text=headers[c])
            self._client_tree.column(c, width=widths.get(c, 100), minwidth=60)

        # Row color tags per statut: fond pastel + texte couleur vive + gras
        for s in _STATUT_ROW_BG:
            self._client_tree.tag_configure(
                s,
                background=_STATUT_ROW_BG[s],
                foreground=_STATUT_COLORS[s],
                font=("Poppins", 10, "bold"),
            )

        self._client_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self._client_tree.bind("<Double-1>", self._on_client_double_click)

    # ── Actions ────────────────────────────────────────────────────────────

    def _reload_clients(self):
        from utils.cache import invalidate_client_cache
        invalidate_client_cache()
        self._load_clients_async()
        self._poll_clients()

    def _get_selected_client(self):
        if self._client_tree is None:
            return None
        sel = self._client_tree.selection()
        if not sel:
            return None
        values = self._client_tree.item(sel[0], "values")
        numero_dossier = values[1]  # numero_dossier at index 1
        nom = values[2]
        for c in self._all_clients:
            if c.get("nom") == nom and (
                not numero_dossier or c.get("numero_dossier") == numero_dossier
            ):
                return c
        return None

    def _on_client_double_click(self, event):
        item = self._client_tree.identify_row(event.y)
        if not item:
            return
        self._client_tree.selection_set(item)
        client = self._get_selected_client()
        if client:
            _ClientActionModal(self.parent, client, self._on_change_statut, self._on_modify_client)

    def _on_change_statut(self, client):
        _ChangeStatutDialog(self.parent, client, on_done=self._reload_clients)

    def _on_modify_client(self, client):
        if self.navigate_callback:
            self.navigate_callback("client_page", client_to_edit=client)

    def _navigate(self, route):
        if self.navigate_callback:
            self.navigate_callback(route)

    def _start_clock(self):
        try:
            if not self.clock_label or not self.clock_label.winfo_exists():
                return
            self.clock_label.configure(text=datetime.now().strftime("%d/%m/%Y\n%H:%M:%S"))
            self.parent.after(1000, self._start_clock)
        except Exception:
            return


# ── Modal: choisir l'action sur un client ─────────────────────────────────────

class _ClientActionModal(tk.Toplevel):

    def __init__(self, parent, client, on_change_statut, on_modify):
        super().__init__(parent)
        self.client = client
        self.on_change_statut = on_change_statut
        self.on_modify = on_modify
        nom = client.get("nom", "")
        statut = client.get("statut") or "En cours"
        self.title(f"Actions — {nom}")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, False)
        self.transient(parent)
        self.after(0, self._safe_grab)
        self._build()

    def _safe_grab(self):
        try:
            self.wait_visibility()
            self.grab_set()
        except Exception:
            pass

    def _build(self):
        nom = self.client.get("nom", "")
        statut = self.client.get("statut") or "En cours"
        color = _STATUT_COLORS.get(statut, BUTTON_BLUE)

        tk.Label(
            self, text=nom,
            font=("Poppins", 15, "bold"),
            fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(padx=32, pady=(20, 4))

        tk.Label(
            self, text=f"Statut : {statut}",
            font=("Poppins", 11),
            fg=color, bg=PANEL_BG_COLOR,
        ).pack(pady=(0, 20))

        btn_frame = tk.Frame(self, bg=PANEL_BG_COLOR)
        btn_frame.pack(padx=32, pady=(0, 24))

        tk.Button(
            btn_frame, text="🔄  Changer le statut",
            font=("Poppins", 11, "bold"),
            bg=BUTTON_ORANGE, fg="white",
            relief="flat", padx=18, pady=10, cursor="hand2",
            command=self._action_statut,
        ).pack(fill="x", pady=(0, 8))

        tk.Button(
            btn_frame, text="✏️  Modifier les informations",
            font=("Poppins", 11, "bold"),
            bg=BUTTON_GREEN, fg="white",
            relief="flat", padx=18, pady=10, cursor="hand2",
            command=self._action_modify,
        ).pack(fill="x")

        self.update_idletasks()
        w, h = self.winfo_reqwidth() + 40, self.winfo_reqheight() + 10
        px = self.master.winfo_rootx() + (self.master.winfo_width() - w) // 2
        py = self.master.winfo_rooty() + (self.master.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{px}+{py}")

    def _action_statut(self):
        client = self.client
        cb = self.on_change_statut
        parent = self.master
        self.destroy()
        parent.after(10, lambda: cb(client))

    def _action_modify(self):
        client = self.client
        cb = self.on_modify
        parent = self.master
        self.destroy()
        parent.after(10, lambda: cb(client))


# ── Dialog: changer le statut ─────────────────────────────────────────────────

class _ChangeStatutDialog(tk.Toplevel):

    _STATUTS = [
        ("En cours",   BUTTON_BLUE,   "#0097A7"),
        ("Accepté",    BUTTON_GREEN,  "#2E7D32"),
        ("En circuit", BUTTON_ORANGE, "#E65100"),
        ("Annulé",     BUTTON_RED,    "#B71C1C"),
    ]

    def __init__(self, parent, client, on_done=None):
        super().__init__(parent)
        self.client = client
        self.on_done = on_done
        nom = client.get("nom", "")
        self.title(f"Changer le statut — {nom}")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, False)
        self.transient(parent)
        self.after(0, self._safe_grab)
        self._build()

    def _safe_grab(self):
        try:
            self.wait_visibility()
            self.grab_set()
        except Exception:
            pass

    def _build(self):
        nom = self.client.get("nom", "")
        current = self.client.get("statut") or "En cours"

        tk.Label(
            self, text=f"Statut de : {nom}",
            font=("Poppins", 13, "bold"),
            fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(padx=32, pady=(20, 6))

        tk.Label(
            self, text=f"Statut actuel : {current}",
            font=("Poppins", 11),
            fg=_STATUT_COLORS.get(current, TEXT_COLOR), bg=PANEL_BG_COLOR,
        ).pack(pady=(0, 14))

        btn_frame = tk.Frame(self, bg=PANEL_BG_COLOR)
        btn_frame.pack(padx=32, pady=(0, 24))

        for label, color, hover in self._STATUTS:
            is_current = label == current
            tk.Button(
                btn_frame, text=f"{'✔  ' if is_current else '      '}{label}",
                font=("Poppins", 11, "bold"),
                bg=color, fg="white",
                relief="flat", padx=18, pady=8, cursor="hand2",
                state="disabled" if is_current else "normal",
                command=lambda s=label: self._apply_statut(s),
            ).pack(fill="x", pady=3)

        self.update_idletasks()
        w = self.winfo_reqwidth() + 60
        h = self.winfo_reqheight() + 10
        px = self.master.winfo_rootx() + (self.master.winfo_width() - w) // 2
        py = self.master.winfo_rooty() + (self.master.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{px}+{py}")

    def _apply_statut(self, new_statut):
        from utils.excel_handler import update_client_statut
        row = self.client.get("row_number")
        if not row:
            messagebox.showerror("Erreur", "Impossible d'identifier la ligne du client.", parent=self)
            return
        if update_client_statut(row, new_statut):
            on_done = self.on_done
            parent = self.master
            self.destroy()
            if on_done:
                parent.after(10, on_done)
        else:
            messagebox.showerror("Erreur", "La mise à jour du statut a échoué.", parent=self)
