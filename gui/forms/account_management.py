"""
Interface de gestion des comptes utilisateurs (admin uniquement).
Accessible depuis le topbar quand l'utilisateur est admin.
"""

from __future__ import annotations

import csv
import re
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk

    PIL_AVAILABLE = True
except Exception:
    Image = ImageDraw = ImageFont = ImageTk = None
    PIL_AVAILABLE = False

from config import (
    BUTTON_BLUE,
    BUTTON_GREEN,
    BUTTON_ORANGE,
    BUTTON_RED,
    INPUT_BG_COLOR,
    MUTED_TEXT_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
)
from utils.activity_log import (
    get_activity,
    get_brute_force_usernames,
    get_recent_activity_summary,
    get_user_stats,
)
from utils.auth_handler import (
    ROLES,
    change_password,
    create_user,
    delete_user,
    duplicate_user,
    get_users,
    reactivate_user,
    set_access_expiry,
    suspend_user,
    update_user_role,
)

_ACCENT = "#C0392B"
_SURFACE = "#F7FBFC"
_HEADER_BG = "#F0F6F8"
_BORDER_CLR = "#D8E8ED"
_ROW_ALT = "#F8FCFD"
_SECURITY_BG = "#FFF1F0"
_PANEL_SHADOW = "#C6D9DE"
_FONT_BOLD = ("Poppins", 10, "bold")
_FONT_BODY = ("Poppins", 10)
_ROLE_SHORT = {
    "admin": "Admin",
    "agent": "Agent",
    "comptable": "Comptable",
}
_ROLE_COLORS = {
    "admin": "#C62828",
    "agent": "#0F7D8A",
    "comptable": "#2E7D32",
}
_STATUS_META = {
    "active": {"label": "Actif", "icon": "●", "color": "#1B7A3E"},
    "pw_expired": {"label": "MDP expiré", "icon": "⚠", "color": "#E65100"},
    "suspended": {"label": "Suspendu", "icon": "⛔", "color": BUTTON_RED},
    "access_expired": {"label": "Accès expiré", "icon": "⏱", "color": "#7B3F00"},
}
_COUNTER_META = {
    "total": ("Total comptes", "👥", "#0F7D8A"),
    "active": ("Actifs", "✅", "#1B7A3E"),
    "suspended": ("Suspendus", "⛔", "#C62828"),
    "expired": ("Expirés", "⏱", "#7B3F00"),
}
_AVATAR_PALETTE = [
    "#0F7D8A",
    "#E07A5F",
    "#5C6BC0",
    "#2E7D32",
    "#D97706",
    "#C62828",
    "#6A4C93",
    "#00897B",
]


def _fmt_date(value: str, fallback: str = "—") -> str:
    if not value:
        return fallback
    return value[:10]


def _fmt_dt(value: str, fallback: str = "—") -> str:
    return value or fallback


def _mix_with_white(color: str, ratio: float = 0.2) -> str:
    color = color.lstrip("#")
    if len(color) != 6:
        return color
    r = int(color[0:2], 16)
    g = int(color[2:4], 16)
    b = int(color[4:6], 16)
    r = int(r + (255 - r) * ratio)
    g = int(g + (255 - g) * ratio)
    b = int(b + (255 - b) * ratio)
    return f"#{r:02X}{g:02X}{b:02X}"


def _status_label(status: str, password_expiring_soon: bool = False) -> str:
    meta = _STATUS_META.get(status, {"icon": "•", "label": status, "color": TEXT_COLOR})
    label = f"{meta['icon']} {meta['label']}"
    if password_expiring_soon and status == "active":
        label += "  •  MDP < 7j"
    return label


def _initials(username: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9 ]+", " ", username).strip()
    if not cleaned:
        return "?"
    parts = [p for p in cleaned.split() if p]
    if len(parts) >= 2:
        return (parts[0][0] + parts[1][0]).upper()
    return parts[0][:2].upper()


def _pick_date(parent, title: str) -> str | None:
    try:
        from gui.forms.client_form import CalendarDialog
    except Exception as exc:
        messagebox.showerror(
            "Erreur",
            f"Impossible d'ouvrir le calendrier.\n\n{exc}",
            parent=parent,
        )
        return None

    cal = CalendarDialog(parent, title)
    parent.wait_window(cal)
    if cal.selected_date:
        return cal.selected_date.strftime("%Y-%m-%d")
    return None


class AccountManagementWindow(tk.Toplevel):
    """Fenêtre de gestion des comptes — admin seulement."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Gestion des comptes")
        self.geometry("1220x760")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(True, True)
        self.minsize(980, 640)
        self.transient(parent)

        self._all_users: list[dict] = []
        self._filtered_users: list[dict] = []
        self._users_by_name: dict[str, dict] = {}
        self._row_by_username: dict[str, str] = {}
        self._avatar_cache: dict[str, object] = {}
        self._sort_col = "username"
        self._sort_desc = False
        self._detail_user: dict | None = None
        self._detail_width = 370
        self._detail_x = self._detail_width + 24
        self._detail_visible = False
        self._detail_anim_job = None
        self._notice_job = None
        self._detail_notice_job = None
        self._badge_job = None
        self._badge_phase = False
        self._syncing_selection = False

        self.after(0, self._safe_focus)
        self._build_ui()
        self._load_users()

    def _safe_focus(self):
        try:
            self.lift()
            self.focus_set()
        except Exception:
            pass

    # ── UI principale ────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_title_bar()

        body = tk.Frame(self, bg=PANEL_BG_COLOR)
        body.pack(fill="both", expand=True, padx=18, pady=16)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(3, weight=1)

        self._build_stats_row(body)
        self._build_filter_bar(body)

        self._notice_lbl = tk.Label(
            body,
            text="",
            font=("Poppins", 10, "bold"),
            fg=MUTED_TEXT_COLOR,
            bg=PANEL_BG_COLOR,
            anchor="w",
        )
        self._notice_lbl.grid(row=2, column=0, sticky="ew", pady=(8, 6))

        self._content = tk.Frame(body, bg=PANEL_BG_COLOR)
        self._content.grid(row=3, column=0, sticky="nsew")
        self._content.columnconfigure(0, weight=1)
        self._content.rowconfigure(0, weight=1)
        self._content.bind("<Configure>", self._on_content_resize)

        self._build_user_list(self._content)
        self._build_detail_panel(self._content)

    def _build_title_bar(self):
        title_bar = tk.Frame(self, bg=_ACCENT, height=56)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)

        left = tk.Frame(title_bar, bg=_ACCENT)
        left.pack(side="left", fill="y")
        tk.Label(
            left,
            text="👤  Gestion des comptes utilisateurs",
            font=("Poppins", 14, "bold"),
            fg="white",
            bg=_ACCENT,
        ).pack(side="left", padx=20)

        right = tk.Frame(title_bar, bg=_ACCENT)
        right.pack(side="right", fill="y")
        ctk.CTkButton(
            right,
            text="➕ Nouveau compte",
            command=self._create_account,
            fg_color="#FDF5F3",
            text_color=_ACCENT,
            hover_color="#F8E0DB",
            width=150,
            height=34,
            corner_radius=10,
            font=ctk.CTkFont(family="Poppins", size=12, weight="bold"),
        ).pack(side="left", padx=(0, 10), pady=11)
        tk.Button(
            right,
            text="✕",
            bg=_ACCENT,
            fg="white",
            font=("Poppins", 12, "bold"),
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=12,
            command=self.destroy,
        ).pack(side="right", padx=(0, 10))

    def _build_stats_row(self, parent):
        row = tk.Frame(parent, bg=PANEL_BG_COLOR)
        row.grid(row=0, column=0, sticky="ew")
        row.columnconfigure(0, weight=1)

        cards = tk.Frame(row, bg=PANEL_BG_COLOR)
        cards.pack(fill="x")

        self._counter_cards: dict[str, tuple[tk.Label, tk.Label]] = {}
        for key, (label, icon, color) in _COUNTER_META.items():
            card = tk.Frame(
                cards,
                bg=color,
                padx=18,
                pady=14,
                highlightbackground=_mix_with_white(color, 0.25),
                highlightthickness=1,
            )
            card.pack(side="left", padx=(0, 12), fill="x", expand=True)
            top = tk.Frame(card, bg=color)
            top.pack(fill="x")
            tk.Label(top, text=icon, font=("Poppins", 18), fg="white", bg=color).pack(side="left")
            value_lbl = tk.Label(top, text="0", font=("Poppins", 22, "bold"), fg="white", bg=color)
            value_lbl.pack(side="right")
            label_lbl = tk.Label(card, text=label, font=("Poppins", 9, "bold"), fg="white", bg=color)
            label_lbl.pack(anchor="w", pady=(8, 0))
            self._counter_cards[key] = (value_lbl, label_lbl)

        self._security_alert = tk.Label(
            parent,
            text="",
            font=("Poppins", 10, "bold"),
            fg="#8A1C1C",
            bg="#FFF1F0",
            padx=12,
            pady=8,
            wraplength=1080,
            justify="left",
            anchor="w",
            highlightbackground="#F2C7C2",
            highlightthickness=1,
        )
        self._security_alert.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self._security_alert.grid_remove()

    def _build_filter_bar(self, parent):
        bar = tk.Frame(parent, bg=_SURFACE, highlightbackground=_BORDER_CLR, highlightthickness=1)
        bar.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        for idx in range(7):
            bar.columnconfigure(idx, weight=0)
        bar.columnconfigure(1, weight=1)

        tk.Label(bar, text="🔍", font=("Poppins", 12), bg=_SURFACE).grid(row=0, column=0, padx=(14, 6), pady=12)
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_filters())
        self._search_entry = ctk.CTkEntry(
            bar,
            textvariable=self._search_var,
            placeholder_text="Rechercher un utilisateur",
            height=36,
            fg_color=INPUT_BG_COLOR,
            text_color=TEXT_COLOR,
            border_color=_BORDER_CLR,
            corner_radius=10,
            font=ctk.CTkFont(family="Poppins", size=12),
        )
        self._search_entry.grid(row=0, column=1, sticky="ew", pady=10, padx=(0, 14))

        tk.Label(bar, text="Rôle", font=_FONT_BOLD, fg=TEXT_COLOR, bg=_SURFACE).grid(row=0, column=2, padx=(0, 6))
        self._role_filter = tk.StringVar(value="Tous")
        self._role_cb = ttk.Combobox(
            bar,
            textvariable=self._role_filter,
            state="readonly",
            width=14,
            values=["Tous", "Admin", "Agent", "Comptable"],
            font=("Poppins", 10),
        )
        self._role_cb.grid(row=0, column=3, padx=(0, 14))
        self._role_cb.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        tk.Label(bar, text="Statut", font=_FONT_BOLD, fg=TEXT_COLOR, bg=_SURFACE).grid(row=0, column=4, padx=(0, 6))
        self._status_filter = tk.StringVar(value="Tous")
        self._status_cb = ttk.Combobox(
            bar,
            textvariable=self._status_filter,
            state="readonly",
            width=16,
            values=["Tous", "Actif", "Suspendu", "Expiré"],
            font=("Poppins", 10),
        )
        self._status_cb.grid(row=0, column=5, padx=(0, 14))
        self._status_cb.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        ctk.CTkButton(
            bar,
            text="Réinitialiser",
            width=120,
            height=34,
            fg_color="#E7F2F5",
            text_color=TEXT_COLOR,
            hover_color="#D7EAEE",
            border_width=1,
            border_color=_BORDER_CLR,
            corner_radius=10,
            font=ctk.CTkFont(family="Poppins", size=11, weight="bold"),
            command=self._reset_filters,
        ).grid(row=0, column=6, padx=(0, 12))

    def _build_user_list(self, parent):
        card = tk.Frame(parent, bg=_SURFACE, highlightbackground=_BORDER_CLR, highlightthickness=1)
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)

        header = tk.Frame(card, bg=_SURFACE)
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 10))
        header.columnconfigure(0, weight=1)
        tk.Label(
            header,
            text="Annuaire des comptes",
            font=("Poppins", 12, "bold"),
            fg=TEXT_COLOR,
            bg=_SURFACE,
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="Clique sur un utilisateur pour ouvrir le panneau de détail",
            font=("Poppins", 9),
            fg=MUTED_TEXT_COLOR,
            bg=_SURFACE,
        ).grid(row=1, column=0, sticky="w")

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Acc.Treeview",
            background=_SURFACE,
            fieldbackground=_SURFACE,
            rowheight=46,
            font=("Poppins", 10),
            borderwidth=0,
        )
        style.configure(
            "Acc.Treeview.Heading",
            background=_HEADER_BG,
            foreground=TEXT_COLOR,
            font=("Poppins", 10, "bold"),
            relief="flat",
            padding=(8, 8),
        )
        style.map(
            "Acc.Treeview",
            background=[("selected", "#DDEEF2")],
            foreground=[("selected", TEXT_COLOR)],
        )
        style.layout("Acc.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

        table_wrap = tk.Frame(card, bg=_BORDER_CLR, bd=1)
        table_wrap.grid(row=1, column=0, sticky="nsew", padx=16)
        table_wrap.columnconfigure(0, weight=1)
        table_wrap.rowconfigure(0, weight=1)

        cols = ("role", "created_at", "last_login", "actions", "status")
        self._tree = ttk.Treeview(
            table_wrap,
            columns=cols,
            show="tree headings",
            style="Acc.Treeview",
            selectmode="browse",
        )
        self._tree.heading("#0", text="Utilisateur", command=lambda: self._sort_by("username"))
        self._tree.column("#0", width=250, minwidth=180, anchor="w")

        headings = {
            "role": ("Rôle", "role"),
            "created_at": ("Création", "created_at"),
            "last_login": ("Dernière connexion", "last_login"),
            "actions": ("Actions", "total_actions"),
            "status": ("Statut", "status"),
        }
        widths = {
            "role": 110,
            "created_at": 120,
            "last_login": 155,
            "actions": 90,
            "status": 190,
        }
        for col, (label, sort_key) in headings.items():
            self._tree.heading(col, text=label, command=lambda key=sort_key: self._sort_by(key))
            anchor = "center" if col in ("actions", "status") else "w"
            self._tree.column(col, width=widths[col], minwidth=80, anchor=anchor)

        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        for status, meta in _STATUS_META.items():
            self._tree.tag_configure(status, foreground=meta["color"])
        self._tree.tag_configure("alt", background=_ROW_ALT)
        self._tree.tag_configure("security", background=_SECURITY_BG)

        self._tree.bind("<<TreeviewSelect>>", self._on_select)

        footer = tk.Frame(card, bg=_SURFACE)
        footer.grid(row=2, column=0, sticky="ew", padx=16, pady=14)
        footer.columnconfigure(0, weight=1)

        self._table_count = tk.Label(
            footer,
            text="0 compte",
            font=("Poppins", 9),
            fg=MUTED_TEXT_COLOR,
            bg=_SURFACE,
        )
        self._table_count.grid(row=0, column=0, sticky="w")

        actions = tk.Frame(footer, bg=_SURFACE)
        actions.grid(row=0, column=1, sticky="e")

        self._btn_reset = self._make_small_button(actions, "🔑 Réinitialiser MDP", BUTTON_ORANGE, self._reset_password)
        self._btn_reset.pack(side="left", padx=(0, 8))
        self._btn_history = self._make_small_button(actions, "📋 Historique", BUTTON_BLUE, self._show_history)
        self._btn_history.pack(side="left", padx=(0, 8))
        self._btn_delete = self._make_small_button(actions, "🗑 Supprimer", BUTTON_RED, self._delete_account)
        self._btn_delete.pack(side="left")
        self._set_list_actions_enabled(False)

    def _build_detail_panel(self, parent):
        self._detail_panel = tk.Frame(
            parent,
            bg=_SURFACE,
            highlightbackground=_PANEL_SHADOW,
            highlightthickness=1,
        )

        header = tk.Frame(self._detail_panel, bg=_SURFACE)
        header.pack(fill="x", padx=18, pady=(16, 10))
        header.columnconfigure(0, weight=1)

        tk.Label(
            header,
            text="Fiche utilisateur",
            font=("Poppins", 12, "bold"),
            fg=TEXT_COLOR,
            bg=_SURFACE,
        ).grid(row=0, column=0, sticky="w")
        tk.Button(
            header,
            text="✕",
            bg=_SURFACE,
            fg=MUTED_TEXT_COLOR,
            relief="flat",
            bd=0,
            font=("Poppins", 11, "bold"),
            cursor="hand2",
            command=self._close_detail_panel,
        ).grid(row=0, column=1, sticky="e")

        hero = tk.Frame(self._detail_panel, bg=_HEADER_BG, highlightbackground=_BORDER_CLR, highlightthickness=1)
        hero.pack(fill="x", padx=18)
        avatar_wrap = tk.Frame(hero, bg=_HEADER_BG)
        avatar_wrap.pack(fill="x", padx=16, pady=16)
        self._detail_avatar = tk.Label(avatar_wrap, bg=_HEADER_BG)
        self._detail_avatar.pack(side="left")
        text_block = tk.Frame(avatar_wrap, bg=_HEADER_BG)
        text_block.pack(side="left", padx=(12, 0), fill="x", expand=True)
        self._detail_name = tk.Label(text_block, text="", font=("Poppins", 14, "bold"), fg=TEXT_COLOR, bg=_HEADER_BG)
        self._detail_name.pack(anchor="w")
        self._detail_role_badge = tk.Label(text_block, text="", font=("Poppins", 9, "bold"), fg="white", bg=BUTTON_BLUE, padx=10, pady=4)
        self._detail_role_badge.pack(anchor="w", pady=(6, 0))
        self._detail_status_badge = tk.Label(hero, text="", font=("Poppins", 10, "bold"), fg="white", bg=BUTTON_GREEN, padx=12, pady=6)
        self._detail_status_badge.pack(anchor="w", padx=16, pady=(0, 14))

        self._detail_notice = tk.Label(
            self._detail_panel,
            text="",
            font=("Poppins", 9, "bold"),
            fg=BUTTON_GREEN,
            bg=_SURFACE,
            wraplength=320,
            justify="left",
            anchor="w",
        )
        self._detail_notice.pack(fill="x", padx=18, pady=(10, 4))

        info = tk.Frame(self._detail_panel, bg=_SURFACE)
        info.pack(fill="x", padx=18, pady=(6, 0))
        self._info_labels: dict[str, tk.Label] = {}
        for key, title in [
            ("created", "Créé le"),
            ("last_login", "Dernière connexion"),
            ("actions", "Actions enregistrées"),
            ("password", "Échéance du mot de passe"),
            ("access", "Accès autorisé jusqu'au"),
        ]:
            row = tk.Frame(info, bg=_SURFACE)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=title, width=22, anchor="w", font=("Poppins", 9, "bold"), fg=TEXT_COLOR, bg=_SURFACE).pack(side="left")
            lbl = tk.Label(row, text="—", anchor="w", font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=_SURFACE)
            lbl.pack(side="left", fill="x", expand=True)
            self._info_labels[key] = lbl

        self._build_role_editor()
        self._build_expiry_editor()
        self._build_action_buttons()
        self._build_activity_summary()

        self._position_detail_panel(self._hidden_panel_x())

    def _build_role_editor(self):
        box = tk.Frame(self._detail_panel, bg=_HEADER_BG, highlightbackground=_BORDER_CLR, highlightthickness=1)
        box.pack(fill="x", padx=18, pady=(14, 0))
        tk.Label(box, text="Modifier le rôle", font=_FONT_BOLD, fg=TEXT_COLOR, bg=_HEADER_BG).pack(anchor="w", padx=14, pady=(14, 8))
        row = tk.Frame(box, bg=_HEADER_BG)
        row.pack(fill="x", padx=14, pady=(0, 14))
        self._detail_role_var = tk.StringVar(value="agent")
        self._detail_role_cb = ttk.Combobox(
            row,
            textvariable=self._detail_role_var,
            state="readonly",
            values=list(ROLES.keys()),
            font=("Poppins", 10),
            width=16,
        )
        self._detail_role_cb.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._detail_role_apply = ctk.CTkButton(
            row,
            text="Appliquer",
            width=110,
            height=34,
            fg_color=BUTTON_BLUE,
            hover_color="#1565C0",
            corner_radius=10,
            font=ctk.CTkFont(family="Poppins", size=11, weight="bold"),
            command=self._apply_role_change,
        )
        self._detail_role_apply.pack(side="left")

    def _build_expiry_editor(self):
        box = tk.Frame(self._detail_panel, bg=_HEADER_BG, highlightbackground=_BORDER_CLR, highlightthickness=1)
        box.pack(fill="x", padx=18, pady=(14, 0))
        tk.Label(box, text="Expiration d'accès", font=_FONT_BOLD, fg=TEXT_COLOR, bg=_HEADER_BG).pack(anchor="w", padx=14, pady=(14, 4))
        tk.Label(box, text="Date libre ou raccourcis rapides", font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=_HEADER_BG).pack(anchor="w", padx=14)

        row = tk.Frame(box, bg=_HEADER_BG)
        row.pack(fill="x", padx=14, pady=(10, 8))
        self._expiry_var = tk.StringVar()
        self._expiry_entry = ctk.CTkEntry(
            row,
            textvariable=self._expiry_var,
            placeholder_text="Cliquer pour choisir une date",
            height=34,
            fg_color=INPUT_BG_COLOR,
            text_color=TEXT_COLOR,
            border_color=_BORDER_CLR,
            corner_radius=10,
            font=ctk.CTkFont(family="Poppins", size=12),
            state="readonly",
        )
        self._expiry_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._expiry_entry._entry.configure(cursor="hand2")
        self._expiry_entry._entry.bind("<Button-1>", lambda e: self._pick_expiry_date())
        ctk.CTkButton(
            row,
            text="Enregistrer",
            width=112,
            height=34,
            fg_color=BUTTON_BLUE,
            hover_color="#1565C0",
            corner_radius=10,
            font=ctk.CTkFont(family="Poppins", size=11, weight="bold"),
            command=self._apply_expiry,
        ).pack(side="left")

        quick = tk.Frame(box, bg=_HEADER_BG)
        quick.pack(fill="x", padx=14, pady=(0, 8))
        for text, days in [("+30j", 30), ("+90j", 90), ("+1 an", 365)]:
            self._make_ghost_button(quick, text, lambda d=days: self._extend_expiry(d), width=82).pack(side="left", padx=(0, 6))
        self._make_ghost_button(quick, "Illimité", self._remove_expiry, width=90).pack(side="left")

        self._expiry_hint = tk.Label(box, text="", font=("Poppins", 9, "bold"), fg=MUTED_TEXT_COLOR, bg=_HEADER_BG)
        self._expiry_hint.pack(anchor="w", padx=14, pady=(0, 14))

    def _build_action_buttons(self):
        box = tk.Frame(self._detail_panel, bg=_SURFACE)
        box.pack(fill="x", padx=18, pady=(14, 0))
        tk.Label(box, text="Actions rapides", font=_FONT_BOLD, fg=TEXT_COLOR, bg=_SURFACE).pack(anchor="w", pady=(0, 8))

        self._btn_suspend = self._make_full_button(box, "🚫 Suspendre l'accès", BUTTON_RED, self._toggle_suspend)
        self._btn_suspend.pack(fill="x", pady=(0, 8))
        self._btn_reset_side = self._make_full_button(box, "🔑 Réinitialiser le mot de passe", BUTTON_ORANGE, self._reset_password)
        self._btn_reset_side.pack(fill="x", pady=(0, 8))
        self._btn_duplicate = self._make_full_button(box, "🧬 Dupliquer ce compte", BUTTON_BLUE, self._duplicate_account)
        self._btn_duplicate.pack(fill="x", pady=(0, 8))
        self._btn_history_side = self._make_full_button(box, "📋 Voir l'historique complet", "#0E7C86", self._show_history)
        self._btn_history_side.pack(fill="x")

    def _build_activity_summary(self):
        box = tk.Frame(self._detail_panel, bg=_HEADER_BG, highlightbackground=_BORDER_CLR, highlightthickness=1)
        box.pack(fill="both", expand=True, padx=18, pady=(14, 18))
        tk.Label(box, text="Résumé d'activité récent", font=_FONT_BOLD, fg=TEXT_COLOR, bg=_HEADER_BG).pack(anchor="w", padx=14, pady=(14, 8))
        self._activity_summary = tk.Frame(box, bg=_HEADER_BG)
        self._activity_summary.pack(fill="both", expand=True, padx=14, pady=(0, 12))

    # ── Helpers UI ────────────────────────────────────────────────────────────

    def _make_small_button(self, parent, text: str, color: str, command):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            height=34,
            corner_radius=10,
            fg_color=color,
            hover_color=_mix_with_white(color, 0.1),
            font=ctk.CTkFont(family="Poppins", size=11, weight="bold"),
        )

    def _make_full_button(self, parent, text: str, color: str, command):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            height=38,
            corner_radius=10,
            fg_color=color,
            hover_color=_mix_with_white(color, 0.08),
            font=ctk.CTkFont(family="Poppins", size=12, weight="bold"),
        )

    def _make_ghost_button(self, parent, text: str, command, width: int = 90):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=30,
            fg_color="transparent",
            hover_color="#E8F0F4",
            border_width=1,
            border_color=_BORDER_CLR,
            text_color=TEXT_COLOR,
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=10, weight="bold"),
        )

    def _avatar_color(self, username: str) -> str:
        return _AVATAR_PALETTE[sum(ord(c) for c in username.lower()) % len(_AVATAR_PALETTE)]

    def _avatar_image(self, username: str, size: int = 34):
        if not PIL_AVAILABLE:
            return None

        key = f"{username}:{size}"
        if key in self._avatar_cache:
            return self._avatar_cache[key]

        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        color = self._avatar_color(username)
        draw.ellipse((0, 0, size - 1, size - 1), fill=color)
        initials = _initials(username)
        try:
            font = ImageFont.truetype("DejaVuSans-Bold.ttf", max(12, size // 2))
        except Exception:
            try:
                font = ImageFont.truetype("Arial.ttf", max(12, size // 2))
            except Exception:
                font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), initials, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        draw.text(((size - tw) / 2, (size - th) / 2 - 1), initials, fill="white", font=font)
        photo = ImageTk.PhotoImage(img)
        self._avatar_cache[key] = photo
        return photo

    def _set_list_actions_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self._btn_reset.configure(state=state)
        self._btn_history.configure(state=state)
        self._btn_delete.configure(state=state)

    def _show_notice(self, text: str, color: str = BUTTON_GREEN, detail: bool = False):
        label = self._detail_notice if detail else self._notice_lbl
        job_attr = "_detail_notice_job" if detail else "_notice_job"
        prev_job = getattr(self, job_attr)
        if prev_job:
            try:
                self.after_cancel(prev_job)
            except Exception:
                pass
        label.configure(text=text, fg=color)
        setattr(self, job_attr, self.after(4200, lambda: label.configure(text="")))

    # ── Chargement / filtre / tri ─────────────────────────────────────────────

    def _load_users(self, preserve_selection: bool = True):
        selected = self._detail_user["username"] if preserve_selection and self._detail_user else None
        brute_force_users = set(u.lower() for u in get_brute_force_usernames())

        users = []
        for user in get_users():
            stats = get_user_stats(user["username"])
            user = dict(user)
            user["role_label"] = _ROLE_SHORT.get(user["role"], user["role"].capitalize())
            user["created_at_label"] = _fmt_date(user.get("created_at", ""))
            user["last_login"] = stats.get("last_login", "")
            user["last_login_label"] = _fmt_dt(stats.get("last_login", ""))
            user["total_actions"] = stats.get("total_actions", 0)
            user["brute_force_detected"] = user["username"].lower() in brute_force_users or stats.get("brute_force_detected", False)
            user["status_label"] = _status_label(user.get("status", ""), user.get("password_expiring_soon", False))
            users.append(user)

        self._all_users = users
        self._users_by_name = {u["username"].lower(): u for u in users}
        self._refresh_counters()
        self._apply_filters()

        if selected:
            user = self._users_by_name.get(selected.lower())
            if user:
                self._open_user_by_name(user["username"], keep_panel=True)
            else:
                self._close_detail_panel()

    def _refresh_counters(self):
        counts = {
            "total": len(self._all_users),
            "active": sum(1 for u in self._all_users if u.get("status") == "active"),
            "suspended": sum(1 for u in self._all_users if u.get("status") == "suspended"),
            "expired": sum(1 for u in self._all_users if u.get("status") in {"access_expired", "pw_expired"}),
        }
        for key, (value_lbl, _) in self._counter_cards.items():
            value_lbl.configure(text=str(counts.get(key, 0)))

        brute_accounts = [u["username"] for u in self._all_users if u.get("brute_force_detected")]
        if brute_accounts:
            txt = "Alerte sécurité : tentatives de brute force détectées pour " + ", ".join(brute_accounts)
            self._security_alert.configure(text=txt)
            self._security_alert.grid()
        else:
            self._security_alert.grid_remove()

    def _reset_filters(self):
        self._search_var.set("")
        self._role_filter.set("Tous")
        self._status_filter.set("Tous")
        self._apply_filters()

    def _apply_filters(self):
        q = self._search_var.get().strip().lower()
        role_filter = self._role_filter.get()
        status_filter = self._status_filter.get()

        users = []
        for user in self._all_users:
            if q and q not in user["username"].lower():
                continue
            if role_filter != "Tous" and user.get("role_label") != role_filter:
                continue
            if status_filter == "Actif" and user.get("status") != "active":
                continue
            if status_filter == "Suspendu" and user.get("status") != "suspended":
                continue
            if status_filter == "Expiré" and user.get("status") not in {"access_expired", "pw_expired"}:
                continue
            users.append(user)

        self._filtered_users = self._sort_users(users)
        self._populate_tree()

    def _sort_users(self, users: list[dict]) -> list[dict]:
        def status_weight(user: dict) -> tuple[int, str]:
            order = {"active": 0, "pw_expired": 1, "access_expired": 2, "suspended": 3}
            return order.get(user.get("status", "active"), 9), user.get("status_label", "")

        key_funcs = {
            "username": lambda u: u["username"].lower(),
            "role": lambda u: u.get("role_label", "").lower(),
            "created_at": lambda u: u.get("created_at", ""),
            "last_login": lambda u: u.get("last_login", ""),
            "total_actions": lambda u: u.get("total_actions", 0),
            "status": status_weight,
        }
        key_func = key_funcs.get(self._sort_col, key_funcs["username"])
        return sorted(users, key=key_func, reverse=self._sort_desc)

    def _sort_by(self, column: str):
        if self._sort_col == column:
            self._sort_desc = not self._sort_desc
        else:
            self._sort_col = column
            self._sort_desc = False
        self._apply_filters()

    def _populate_tree(self):
        selected = self._detail_user["username"].lower() if self._detail_user else None
        for item in self._tree.get_children():
            self._tree.delete(item)
        self._row_by_username.clear()

        for idx, user in enumerate(self._filtered_users):
            tags = [user.get("status", "active")]
            if idx % 2 == 1:
                tags.append("alt")
            if user.get("brute_force_detected"):
                tags.append("security")
            image = self._avatar_image(user["username"])
            text = user["username"] + ("  ⚠" if user.get("password_expiring_soon") else "")
            if not image:
                text = f"{_initials(user['username'])}  {text}"
            iid = self._tree.insert(
                "",
                "end",
                text=text,
                image=image,
                values=(
                    user.get("role_label", ""),
                    user.get("created_at_label", "—"),
                    user.get("last_login_label", "—"),
                    user.get("total_actions", 0),
                    user.get("status_label", ""),
                ),
                tags=tuple(tags),
            )
            self._row_by_username[user["username"].lower()] = iid

        count = len(self._filtered_users)
        suffix = "s" if count > 1 else ""
        self._table_count.configure(text=f"{count} compte{suffix} affiché{suffix}")

        if selected and selected in self._row_by_username:
            iid = self._row_by_username[selected]
            self._tree.selection_set(iid)
            self._tree.see(iid)
            self._set_list_actions_enabled(True)
        else:
            self._set_list_actions_enabled(False)
            if self._detail_user and self._detail_user["username"].lower() not in self._row_by_username:
                self._close_detail_panel()

    # ── Sélection & panneau latéral ──────────────────────────────────────────

    def _on_select(self, event=None):
        if self._syncing_selection:
            return
        sel = self._tree.selection()
        if not sel:
            self._set_list_actions_enabled(False)
            return
        iid = sel[0]
        username = next(
            (name for name, row_id in self._row_by_username.items() if row_id == iid),
            None,
        )
        if not username:
            return
        self._set_list_actions_enabled(True)
        self._open_user_by_name(username, sync_selection=False)

    def _open_user_by_name(self, username: str, keep_panel: bool = False, sync_selection: bool = True):
        user = self._users_by_name.get(username.lower())
        if not user:
            return
        self._detail_user = user
        if sync_selection and username.lower() in self._row_by_username:
            iid = self._row_by_username[username.lower()]
            current_selection = self._tree.selection()
            if current_selection != (iid,):
                self._syncing_selection = True
                try:
                    self._tree.selection_set(iid)
                finally:
                    self._syncing_selection = False
            self._tree.focus(iid)
        self._refresh_detail_panel()
        if not keep_panel or not self._detail_visible:
            self._open_detail_panel()

    def _refresh_detail_panel(self):
        user = self._detail_user
        if not user:
            return

        avatar = self._avatar_image(user["username"], size=56)
        if avatar:
            self._detail_avatar.configure(image=avatar, text="", bg=_HEADER_BG)
            self._detail_avatar.image = avatar
        else:
            self._detail_avatar.configure(
                image="",
                text=_initials(user["username"]),
                font=("Poppins", 18, "bold"),
                fg="white",
                bg=self._avatar_color(user["username"]),
                width=4,
                height=2,
            )
            self._detail_avatar.image = None
        self._detail_name.configure(text=user["username"])

        role_color = _ROLE_COLORS.get(user["role"], BUTTON_BLUE)
        self._detail_role_badge.configure(text=user.get("role_label", user["role"]), bg=role_color)
        self._detail_role_var.set(user["role"])

        status_meta = _STATUS_META.get(user["status"], _STATUS_META["active"])
        self._detail_status_badge.configure(text=user.get("status_label", status_meta["label"]), bg=status_meta["color"])
        self._badge_phase = False
        self._schedule_badge_pulse()

        self._info_labels["created"].configure(text=_fmt_dt(user.get("created_at", "")))
        self._info_labels["last_login"].configure(text=user.get("last_login_label", "—"))
        self._info_labels["actions"].configure(text=str(user.get("total_actions", 0)))

        pw_text = user.get("password_expires_at") or "—"
        if user.get("password_expiring_soon"):
            days_left = user.get("password_days_left")
            label = f"{pw_text}  •  attention, {max(days_left, 0)} jour(s) restant(s)" if days_left is not None else pw_text
            self._info_labels["password"].configure(text=label, fg="#E65100")
        elif user.get("status") == "pw_expired":
            self._info_labels["password"].configure(text=f"{pw_text}  •  expiré", fg=BUTTON_RED)
        else:
            self._info_labels["password"].configure(text=pw_text, fg=MUTED_TEXT_COLOR)

        access_until = user.get("access_expires_at") or "Illimité"
        self._info_labels["access"].configure(text=access_until)
        self._expiry_var.set("" if access_until == "Illimité" else access_until)
        self._expiry_hint.configure(text=f"Valeur actuelle : {access_until}")

        if user.get("status") == "suspended":
            self._btn_suspend.configure(text="✅ Réactiver l'accès", fg_color=BUTTON_GREEN, hover_color="#1B5E20")
        else:
            self._btn_suspend.configure(text="🚫 Suspendre l'accès", fg_color=BUTTON_RED, hover_color="#7F0000")

        self._populate_activity_summary()
        self._detail_notice.configure(text="")

    def _populate_activity_summary(self):
        for child in self._activity_summary.winfo_children():
            child.destroy()

        if not self._detail_user:
            return

        entries = get_recent_activity_summary(self._detail_user["username"], limit=5)
        if not entries:
            tk.Label(
                self._activity_summary,
                text="Aucune activité récente.",
                font=("Poppins", 9),
                fg=MUTED_TEXT_COLOR,
                bg=_HEADER_BG,
            ).pack(anchor="w")
            return

        for entry in entries:
            row = tk.Frame(self._activity_summary, bg=_HEADER_BG)
            row.pack(fill="x", pady=(0, 8))
            tk.Label(
                row,
                text=entry.get("timestamp", "")[:16],
                width=16,
                anchor="w",
                font=("Poppins", 8, "bold"),
                fg=MUTED_TEXT_COLOR,
                bg=_HEADER_BG,
            ).pack(side="left")
            text = entry.get("label", entry.get("action", ""))
            if entry.get("details"):
                text += f" — {entry['details']}"
            tk.Label(
                row,
                text=text,
                anchor="w",
                justify="left",
                wraplength=240,
                font=("Poppins", 9),
                fg=TEXT_COLOR,
                bg=_HEADER_BG,
            ).pack(side="left", fill="x", expand=True)

    def _hidden_panel_x(self):
        width = max(self._content.winfo_width(), 1)
        return width + 18

    def _visible_panel_x(self):
        width = max(self._content.winfo_width(), self._detail_width + 40)
        return max(width - self._detail_width, 0)

    def _position_detail_panel(self, x: int):
        self._detail_x = x
        self._detail_panel.place(x=x, y=0, width=self._detail_width, relheight=1)

    def _on_content_resize(self, event=None):
        target_x = self._visible_panel_x() if self._detail_visible else self._hidden_panel_x()
        self._position_detail_panel(target_x)

    def _animate_panel(self, show: bool):
        if self._detail_anim_job:
            try:
                self.after_cancel(self._detail_anim_job)
            except Exception:
                pass
        self._detail_visible = show
        target = self._visible_panel_x() if show else self._hidden_panel_x()

        def _step():
            current = self._detail_x
            distance = target - current
            if abs(distance) <= 12:
                self._position_detail_panel(target)
                self._detail_anim_job = None
                return
            self._position_detail_panel(current + int(distance * 0.35))
            self._detail_anim_job = self.after(18, _step)

        _step()

    def _open_detail_panel(self):
        self._animate_panel(True)
        self.after(40, self._schedule_badge_pulse)

    def _close_detail_panel(self):
        self._detail_user = None
        self._tree.selection_remove(self._tree.selection())
        self._set_list_actions_enabled(False)
        if self._badge_job:
            try:
                self.after_cancel(self._badge_job)
            except Exception:
                pass
            self._badge_job = None
        self._animate_panel(False)

    def _schedule_badge_pulse(self):
        if self._badge_job:
            try:
                self.after_cancel(self._badge_job)
            except Exception:
                pass

        if not self._detail_user or not self._detail_visible:
            return

        status = self._detail_user.get("status", "active")
        base = _STATUS_META.get(status, _STATUS_META["active"])["color"]
        alt = _mix_with_white(base, 0.18)
        self._badge_phase = not self._badge_phase
        self._detail_status_badge.configure(bg=alt if self._badge_phase else base)
        self._badge_job = self.after(700, self._schedule_badge_pulse)

    # ── Actions ──────────────────────────────────────────────────────────────

    def _get_selected_username(self):
        return self._detail_user["username"] if self._detail_user else None

    def _create_account(self):
        _CreateAccountDialog(self, on_created=lambda: self._load_users(preserve_selection=False))

    def _reset_password(self):
        username = self._get_selected_username()
        if not username:
            return
        _AdminResetPasswordDialog(self, username, on_done=self._load_users)

    def _delete_account(self):
        username = self._get_selected_username()
        if not username:
            return
        if not messagebox.askyesno(
            "Confirmation",
            f"Supprimer le compte « {username} » ?\nCette action est irréversible.",
            icon="warning",
            parent=self,
        ):
            return
        success, err = delete_user(username)
        if success:
            self._show_notice(f"Compte « {username} » supprimé.")
            self._close_detail_panel()
            self._load_users(preserve_selection=False)
        else:
            self._show_notice(err, color=BUTTON_RED, detail=True)

    def _show_history(self):
        username = self._get_selected_username()
        if not username:
            return
        _ActivityHistoryDialog(self, username)

    def _pick_expiry_date(self):
        if not self._detail_user:
            return
        selected_date = _pick_date(self, "Choisir la date d'expiration")
        if selected_date:
            self._expiry_var.set(selected_date)

    def _apply_expiry(self):
        username = self._get_selected_username()
        if not username:
            return
        date_str = self._expiry_var.get().strip()
        if date_str and not re.match(r"^\d{4}-\d{2}-\d{2}$", date_str):
            self._show_notice("Format invalide. Utilisez AAAA-MM-JJ.", color=BUTTON_RED, detail=True)
            return
        success, err = set_access_expiry(username, date_str)
        if success:
            self._show_notice("Date d'expiration mise à jour.", detail=True)
            self._load_users()
        else:
            self._show_notice(err, color=BUTTON_RED, detail=True)

    def _remove_expiry(self):
        username = self._get_selected_username()
        if not username:
            return
        success, err = set_access_expiry(username, "")
        if success:
            self._show_notice("Limite d'accès supprimée.", detail=True)
            self._load_users()
        else:
            self._show_notice(err, color=BUTTON_RED, detail=True)

    def _extend_expiry(self, days: int):
        username = self._get_selected_username()
        if not username or not self._detail_user:
            return

        base_date = datetime.now().date()
        current_expiry = self._detail_user.get("access_expires_at", "")
        if current_expiry:
            try:
                parsed = datetime.strptime(current_expiry, "%Y-%m-%d").date()
                if parsed > base_date:
                    base_date = parsed
            except ValueError:
                pass
        new_date = (base_date + timedelta(days=days)).strftime("%Y-%m-%d")
        success, err = set_access_expiry(username, new_date)
        if success:
            self._show_notice(f"Accès prolongé jusqu'au {new_date}.", detail=True)
            self._load_users()
        else:
            self._show_notice(err, color=BUTTON_RED, detail=True)

    def _toggle_suspend(self):
        username = self._get_selected_username()
        if not username or not self._detail_user:
            return
        if self._detail_user.get("status") == "suspended":
            ok, err = reactivate_user(username)
            ok_msg = f"Compte « {username} » réactivé."
        else:
            ok, err = suspend_user(username)
            ok_msg = f"Compte « {username} » suspendu."
        if ok:
            self._show_notice(ok_msg, detail=True)
            self._load_users()
        else:
            self._show_notice(err, color=BUTTON_RED, detail=True)

    def _apply_role_change(self):
        username = self._get_selected_username()
        if not username:
            return
        new_role = self._detail_role_var.get().strip()
        success, err = update_user_role(username, new_role)
        if success:
            label = _ROLE_SHORT.get(new_role, new_role)
            self._show_notice(f"Rôle mis à jour : {label}.", detail=True)
            self._load_users()
        else:
            self._show_notice(err, color=BUTTON_RED, detail=True)

    def _duplicate_account(self):
        if not self._detail_user:
            return
        _DuplicateAccountDialog(
            self,
            source_user=self._detail_user,
            on_done=lambda: self._load_users(preserve_selection=False),
        )


# ── Dialog : historique d'activité ───────────────────────────────────────────

class _ActivityHistoryDialog(tk.Toplevel):
    """Historique d'activité complet avec filtres, tri et export CSV."""

    _CAT_COLORS = {
        "auth": "#1565C0",
        "client": "#1B7A3E",
        "user": "#C62828",
        "quota": "#6A1B9A",
        "invoice": "#E65100",
        "nav": "#78909C",
    }
    _CAT_LABELS = {
        "": "Toutes",
        "auth": "🔐 Connexion",
        "client": "👥 Clients",
        "user": "👤 Comptes",
        "quota": "📄 Cotations",
        "invoice": "💶 Factures",
        "nav": "🧭 Navigation",
    }

    def __init__(self, parent, username: str):
        super().__init__(parent)
        self.username = username
        self.title(f"Historique d'activité — {username}")
        self.geometry("980x640")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(True, True)
        self.minsize(780, 480)
        self.transient(parent)
        self.after(0, self._safe_focus)
        self._sort_col = "timestamp"
        self._sort_desc = True
        self._all_entries: list[dict] = []
        self._visible_entries: list[dict] = []
        self._build_ui()

    def _safe_focus(self):
        try:
            self.lift()
            self.focus_set()
        except Exception:
            pass

    def _build_ui(self):
        title_bar = tk.Frame(self, bg=_ACCENT, height=48)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)
        tk.Label(
            title_bar,
            text=f"📋  Historique d'activité — {self.username}",
            font=("Poppins", 12, "bold"),
            fg="white",
            bg=_ACCENT,
        ).pack(side="left", padx=16)
        tk.Button(
            title_bar,
            text="✕",
            bg=_ACCENT,
            fg="white",
            font=("Poppins", 11, "bold"),
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=10,
            command=self.destroy,
        ).pack(side="right")

        body = tk.Frame(self, bg=PANEL_BG_COLOR)
        body.pack(fill="both", expand=True, padx=14, pady=10)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        fbar = tk.Frame(body, bg=_HEADER_BG, highlightbackground="#C9DDE3", highlightthickness=1)
        fbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        tk.Label(fbar, text="Afficher :", font=("Poppins", 9, "bold"), fg=TEXT_COLOR, bg=_HEADER_BG).grid(row=0, column=0, padx=(10, 4), pady=8, sticky="w")
        self._filter_var = tk.StringVar(value="user")
        for col_idx, (val, lbl) in enumerate([("user", f"Seulement {self.username}"), ("all", "Tous les utilisateurs")], start=1):
            tk.Radiobutton(
                fbar,
                text=lbl,
                value=val,
                variable=self._filter_var,
                bg=_HEADER_BG,
                fg=TEXT_COLOR,
                selectcolor=BUTTON_BLUE,
                font=("Poppins", 9),
                command=self._reload,
            ).grid(row=0, column=col_idx, padx=(0, 10), pady=8)

        tk.Frame(fbar, bg="#C9DDE3", width=1).grid(row=0, column=3, padx=8, pady=4, sticky="ns")
        tk.Label(fbar, text="Catégorie :", font=("Poppins", 9, "bold"), fg=TEXT_COLOR, bg=_HEADER_BG).grid(row=0, column=4, padx=(0, 4), pady=8)
        self._cat_var = tk.StringVar(value="")
        cat_cb = ttk.Combobox(fbar, textvariable=self._cat_var, state="readonly", values=list(self._CAT_LABELS.values()), width=16, font=("Poppins", 9))
        cat_cb.grid(row=0, column=5, padx=(0, 8), pady=8)
        cat_cb.bind("<<ComboboxSelected>>", lambda e: self._reload())

        tk.Label(fbar, text="Du :", font=("Poppins", 9, "bold"), fg=TEXT_COLOR, bg=_HEADER_BG).grid(row=0, column=6, padx=(0, 4))
        self._date_from = tk.StringVar()
        tk.Entry(fbar, textvariable=self._date_from, width=11, font=("Poppins", 9)).grid(row=0, column=7, padx=(0, 6))
        tk.Label(fbar, text="Au :", font=("Poppins", 9, "bold"), fg=TEXT_COLOR, bg=_HEADER_BG).grid(row=0, column=8, padx=(0, 4))
        self._date_to = tk.StringVar()
        tk.Entry(fbar, textvariable=self._date_to, width=11, font=("Poppins", 9)).grid(row=0, column=9, padx=(0, 6))
        tk.Button(
            fbar,
            text="Appliquer",
            font=("Poppins", 8, "bold"),
            fg="white",
            bg=BUTTON_BLUE,
            relief="flat",
            bd=0,
            padx=8,
            pady=3,
            cursor="hand2",
            command=self._reload,
        ).grid(row=0, column=10, padx=(0, 6))
        tk.Button(
            fbar,
            text="↺",
            font=("Poppins", 10),
            fg=TEXT_COLOR,
            bg=_HEADER_BG,
            relief="flat",
            bd=0,
            cursor="hand2",
            command=self._reset_filters,
        ).grid(row=0, column=11, padx=(0, 6))

        tk.Frame(fbar, bg="#C9DDE3", width=1).grid(row=0, column=12, padx=8, pady=4, sticky="ns")
        tk.Label(fbar, text="🔍", font=("Poppins", 11), bg=_HEADER_BG).grid(row=0, column=13, padx=(0, 2))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_search())
        tk.Entry(fbar, textvariable=self._search_var, width=18, font=("Poppins", 9)).grid(row=0, column=14, padx=(0, 10))

        tab_bar = tk.Frame(body, bg=PANEL_BG_COLOR)
        tab_bar.grid(row=1, column=0, sticky="ew")

        self._tab_hist_btn = tk.Label(tab_bar, text="Historique", font=("Poppins", 10, "bold"), bg=BUTTON_RED, fg="white", padx=14, pady=4, cursor="hand2")
        self._tab_hist_btn.pack(side="left", padx=(0, 4))
        self._tab_stat_btn = tk.Label(tab_bar, text="Statistiques", font=("Poppins", 10, "bold"), bg=BUTTON_BLUE, fg="white", padx=14, pady=4, cursor="hand2")
        self._tab_stat_btn.pack(side="left")
        self._tab_hist_btn.bind("<Button-1>", lambda e: self._show_tab("hist"))
        self._tab_stat_btn.bind("<Button-1>", lambda e: self._show_tab("stat"))
        tk.Button(
            tab_bar,
            text="⬇ Exporter CSV",
            font=("Poppins", 9, "bold"),
            fg="white",
            bg=BUTTON_BLUE,
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._export_csv,
        ).pack(side="right")

        body.rowconfigure(2, weight=1)
        self._hist_frame = tk.Frame(body, bg=PANEL_BG_COLOR)
        self._hist_frame.grid(row=2, column=0, sticky="nsew", pady=(4, 0))
        self._hist_frame.columnconfigure(0, weight=1)
        self._hist_frame.rowconfigure(0, weight=1)
        self._build_treeview(self._hist_frame)

        self._count_lbl = tk.Label(body, text="", font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR)
        self._count_lbl.grid(row=3, column=0, sticky="e", pady=(2, 0))

        self._stat_frame = tk.Frame(body, bg=PANEL_BG_COLOR)
        self._reload()

    def _build_treeview(self, parent):
        style = ttk.Style()
        style.configure("Hist.Treeview", background=PANEL_BG_COLOR, fieldbackground=PANEL_BG_COLOR, font=("Poppins", 9), rowheight=26)
        style.configure("Hist.Treeview.Heading", font=("Poppins", 9, "bold"), background=_HEADER_BG)

        cols = ("timestamp", "username", "role", "label", "details")
        self._tree = ttk.Treeview(parent, columns=cols, show="headings", style="Hist.Treeview")
        widths = {"timestamp": 140, "username": 100, "role": 80, "label": 160, "details": 0}
        heads = {"timestamp": "Date / Heure", "username": "Utilisateur", "role": "Rôle", "label": "Action", "details": "Détails"}
        for col in cols:
            self._tree.heading(col, text=heads[col], command=lambda c=col: self._sort_by_col(c))
            if widths[col]:
                self._tree.column(col, width=widths[col], minwidth=60, stretch=False)
            else:
                self._tree.column(col, stretch=True, minwidth=120)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        for cat, color in self._CAT_COLORS.items():
            self._tree.tag_configure(cat, foreground=color)
        self._tree.tag_configure("odd", background="#F5F9FB")
        self._tree.tag_configure("auth_alert", foreground="#C62828", font=("Poppins", 9, "bold"))

    def _build_stats_panel(self, parent):
        stats = get_user_stats(self.username)
        for w in parent.winfo_children():
            w.destroy()

        cards_row = tk.Frame(parent, bg=PANEL_BG_COLOR)
        cards_row.pack(fill="x", pady=(4, 12))

        def _stat_card(frame, icon, value, label, color):
            card = tk.Frame(frame, bg=color, padx=14, pady=10)
            card.pack(side="left", padx=(0, 10), fill="y")
            tk.Label(card, text=icon, font=("Poppins", 18), bg=color, fg="white").pack()
            tk.Label(card, text=str(value), font=("Poppins", 20, "bold"), bg=color, fg="white").pack()
            tk.Label(card, text=label, font=("Poppins", 8), bg=color, fg="white").pack()

        _stat_card(cards_row, "🔐", stats["login_count"], "Connexions", "#1565C0")
        _stat_card(cards_row, "⚠", stats["failed_logins"], "Échecs login", "#C62828")
        _stat_card(cards_row, "📋", stats["total_actions"], "Total actions", "#1B7A3E")
        if stats["brute_force_detected"]:
            _stat_card(cards_row, "🚨", "Alerte", "Brute force", "#7F0000")

        info = tk.Frame(parent, bg=PANEL_BG_COLOR)
        info.pack(fill="x", pady=(0, 10))
        for lbl, val in [("Dernière connexion :", stats["last_login"] or "—"), ("Dernière action :", stats["last_action_ts"] or "—")]:
            row = tk.Frame(info, bg=PANEL_BG_COLOR)
            row.pack(anchor="w")
            tk.Label(row, text=lbl, font=("Poppins", 9, "bold"), fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left", padx=(0, 6))
            tk.Label(row, text=val, font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left")

        if stats["top_actions"]:
            tk.Label(parent, text="Actions les plus fréquentes (30 jours)", font=("Poppins", 10, "bold"), fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(4, 6))
            for action_lbl, count in stats["top_actions"]:
                bar_frame = tk.Frame(parent, bg=PANEL_BG_COLOR)
                bar_frame.pack(fill="x", pady=(0, 4))
                tk.Label(bar_frame, text=action_lbl, width=30, anchor="w", font=("Poppins", 9), fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left")
                max_count = stats["top_actions"][0][1] if stats["top_actions"] else 1
                bar_w = max(4, int(200 * count / max_count))
                tk.Frame(bar_frame, bg=BUTTON_BLUE, width=bar_w, height=14).pack(side="left", padx=(4, 6))
                tk.Label(bar_frame, text=str(count), font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left")

    def _show_tab(self, tab: str):
        if tab == "hist":
            self._stat_frame.grid_forget()
            self._hist_frame.grid(row=2, column=0, sticky="nsew", pady=(4, 0))
            self._tab_hist_btn.configure(bg=BUTTON_RED)
            self._tab_stat_btn.configure(bg=BUTTON_BLUE)
        else:
            self._hist_frame.grid_forget()
            self._stat_frame.grid(row=2, column=0, sticky="nsew", pady=(4, 0))
            self._tab_hist_btn.configure(bg=BUTTON_BLUE)
            self._tab_stat_btn.configure(bg=BUTTON_RED)
            self._build_stats_panel(self._stat_frame)

    def _get_cat_filter(self) -> str:
        label = self._cat_var.get()
        for key, value in self._CAT_LABELS.items():
            if value == label:
                return key
        return ""

    def _reset_filters(self):
        self._filter_var.set("user")
        self._cat_var.set("")
        self._date_from.set("")
        self._date_to.set("")
        self._search_var.set("")
        self._reload()

    def _reload(self):
        username = "" if self._filter_var.get() == "all" else self.username
        self._all_entries = get_activity(
            username=username,
            limit=1000,
            action_filter=self._get_cat_filter(),
            date_from=self._date_from.get().strip(),
            date_to=self._date_to.get().strip(),
        )
        self._apply_search()

    def _apply_search(self):
        q = self._search_var.get().strip().lower()
        entries = self._all_entries
        if q:
            entries = [e for e in entries if any(q in str(v).lower() for v in e.values())]
        self._visible_entries = list(entries)
        self._populate_history_tree(entries)

    def _populate_history_tree(self, entries: list[dict]):
        for item in self._tree.get_children():
            self._tree.delete(item)
        for idx, entry in enumerate(entries):
            cat = entry.get("category", "")
            action = entry.get("action", "")
            tags = [cat] if cat else []
            if action == "brute_force_alert":
                tags = ["auth_alert"]
            elif idx % 2 == 1:
                tags.append("odd")
            self._tree.insert(
                "",
                "end",
                values=(
                    entry.get("timestamp", ""),
                    entry.get("username", ""),
                    entry.get("role", ""),
                    entry.get("label", entry.get("action", "")),
                    entry.get("details", ""),
                ),
                tags=tuple(tags),
            )
        count = len(entries)
        suffix = "s" if count > 1 else ""
        self._count_lbl.configure(text=f"{count} entrée{suffix} affichée{suffix}")

    def _sort_by_col(self, col: str):
        if self._sort_col == col:
            self._sort_desc = not self._sort_desc
        else:
            self._sort_col = col
            self._sort_desc = True
        self._all_entries.sort(key=lambda e: str(e.get(col, "")).lower(), reverse=self._sort_desc)
        heads = {"timestamp": "Date / Heure", "username": "Utilisateur", "role": "Rôle", "label": "Action", "details": "Détails"}
        arrow = " ▼" if self._sort_desc else " ▲"
        for key, head in heads.items():
            self._tree.heading(key, text=head + (arrow if key == col else ""))
        self._apply_search()

    def _export_csv(self):
        default = f"historique_{self.username}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        path = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Tous", "*.*")],
            initialfile=default,
            title="Exporter l'historique",
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=["timestamp", "username", "role", "label", "category", "details"],
                    extrasaction="ignore",
                )
                writer.writeheader()
                writer.writerows(self._visible_entries)
            messagebox.showinfo("Export réussi", f"Fichier exporté :\n{path}", parent=self)
        except Exception as ex:
            messagebox.showerror("Erreur", str(ex), parent=self)


# ── Dialog : créer un compte ─────────────────────────────────────────────────

class _CreateAccountDialog(tk.Toplevel):
    def __init__(self, parent, on_created=None):
        super().__init__(parent)
        self.on_created = on_created
        self.title("Nouveau compte")
        self.geometry("540x640")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, True)
        self.minsize(520, 600)
        self.transient(parent)
        self.after(0, self._safe_focus)
        self._build_ui()

    def _safe_focus(self):
        try:
            self.lift()
            self.focus_set()
        except Exception:
            pass

    def _build_ui(self):
        bar = tk.Frame(self, bg=BUTTON_GREEN, height=46)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="➕  Créer un nouveau compte", font=("Poppins", 12, "bold"), fg="white", bg=BUTTON_GREEN).pack(side="left", padx=20)

        footer = tk.Frame(self, bg=PANEL_BG_COLOR)
        footer.pack(side="bottom", fill="x", padx=24, pady=12)
        self._msg = ctk.CTkLabel(footer, text="", font=ctk.CTkFont(family="Poppins", size=11), text_color=BUTTON_RED, wraplength=440)
        self._msg.pack(anchor="w", pady=(0, 6))

        btn_row = tk.Frame(footer, bg=PANEL_BG_COLOR)
        btn_row.pack(fill="x")
        ctk.CTkButton(
            btn_row,
            text="✓  Créer le compte",
            height=40,
            fg_color=BUTTON_GREEN,
            hover_color="#1B5E20",
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=13, weight="bold"),
            command=self._create,
        ).pack(side="left", expand=True, fill="x", padx=(0, 8))
        ctk.CTkButton(
            btn_row,
            text="Annuler",
            width=110,
            height=40,
            fg_color="#AAAAAA",
            hover_color="#888888",
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12),
            command=self.destroy,
        ).pack(side="left")

        body = tk.Frame(self, bg=PANEL_BG_COLOR)
        body.pack(fill="both", expand=True, padx=24, pady=(12, 0))

        def _field(label, attr, show=""):
            tk.Label(body, text=label, font=("Poppins", 10, "bold"), fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(6, 2))
            entry = ctk.CTkEntry(
                body,
                height=34,
                fg_color=INPUT_BG_COLOR,
                text_color=TEXT_COLOR,
                border_color=_BORDER_CLR,
                corner_radius=8,
                font=ctk.CTkFont(family="Poppins", size=12),
                show=show,
            )
            entry.pack(fill="x")
            setattr(self, attr, entry)

        _field("Nom d'utilisateur", "_e_user")
        _field("Mot de passe  (min. 6 caractères)", "_e_pass", show="•")
        _field("Confirmer le mot de passe", "_e_confirm", show="•")

        tk.Label(body, text="Date d'expiration d'accès  (optionnel)", font=("Poppins", 10, "bold"), fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(8, 2))
        self._expiry_var = tk.StringVar(value="")
        self._e_expiry = ctk.CTkEntry(
            body,
            textvariable=self._expiry_var,
            height=36,
            fg_color=INPUT_BG_COLOR,
            text_color=TEXT_COLOR,
            border_color=_BORDER_CLR,
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12),
            placeholder_text="Cliquer pour choisir une date",
            state="readonly",
        )
        self._e_expiry.pack(fill="x")
        self._e_expiry._entry.configure(cursor="hand2")
        self._e_expiry._entry.bind("<Button-1>", lambda e: self._pick_expiry_date())

        quick = tk.Frame(body, bg=PANEL_BG_COLOR)
        quick.pack(anchor="w", pady=(8, 0))
        for text, days in [("+30j", 30), ("+90j", 90), ("+1 an", 365)]:
            ctk.CTkButton(
                quick,
                text=text,
                width=72,
                height=28,
                fg_color="transparent",
                hover_color="#E8F0F4",
                border_width=1,
                border_color=_BORDER_CLR,
                text_color=TEXT_COLOR,
                corner_radius=8,
                font=ctk.CTkFont(family="Poppins", size=10, weight="bold"),
                command=lambda d=days: self._set_expiry_days(d),
            ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            quick,
            text="Illimité",
            width=84,
            height=28,
            fg_color="transparent",
            hover_color="#E8F0F4",
            border_width=1,
            border_color=_BORDER_CLR,
            text_color=TEXT_COLOR,
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=10, weight="bold"),
            command=lambda: self._expiry_var.set(""),
        ).pack(side="left")

        tk.Label(body, text="Rôle", font=("Poppins", 10, "bold"), fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(12, 6))
        self._role_var = tk.StringVar(value="agent")
        role_box = tk.Frame(body, bg=PANEL_BG_COLOR)
        role_box.pack(fill="x")
        for i in range(3):
            role_box.columnconfigure(i, weight=1)

        cfg = {
            "admin": {"icon": "🛡", "color": BUTTON_RED, "desc": "Accès complet"},
            "agent": {"icon": "👤", "color": BUTTON_BLUE, "desc": "Clients, cotations"},
            "comptable": {"icon": "📊", "color": "#2E7D32", "desc": "Comptabilité"},
        }
        self._role_cards = {}

        def _select(role):
            self._role_var.set(role)
            _refresh()

        def _refresh():
            selected = self._role_var.get()
            for role, (widgets, role_cfg) in self._role_cards.items():
                if role == selected:
                    bg, fg, muted = role_cfg["color"], "white", "#E4E4E4"
                else:
                    bg, fg, muted = "#E8F4F8", TEXT_COLOR, MUTED_TEXT_COLOR
                for widget in widgets:
                    widget.configure(bg=bg)
                    if isinstance(widget, tk.Label):
                        widget.configure(fg=fg)
                widgets[-1].configure(fg=muted)

        for col, role in enumerate(cfg):
            card = tk.Frame(role_box, bg="#E8F4F8", cursor="hand2", relief="solid", bd=1, highlightbackground=_BORDER_CLR)
            card.grid(row=0, column=col, sticky="nsew", padx=(0, 4) if col < 2 else (4, 0), pady=2)
            top = tk.Frame(card, bg="#E8F4F8")
            top.pack(fill="x", padx=8, pady=(8, 2))
            icon = tk.Label(top, text=cfg[role]["icon"], font=("Poppins", 15), bg="#E8F4F8", fg=TEXT_COLOR)
            icon.pack(side="left")
            name = tk.Label(top, text=role.capitalize(), font=("Poppins", 10, "bold"), bg="#E8F4F8", fg=TEXT_COLOR)
            name.pack(side="left", padx=(6, 0))
            desc = tk.Label(card, text=cfg[role]["desc"], font=("Poppins", 9), bg="#E8F4F8", fg=MUTED_TEXT_COLOR, wraplength=140, justify="left")
            desc.pack(anchor="w", padx=8, pady=(0, 10))
            widgets = [card, top, icon, name, desc]
            self._role_cards[role] = (widgets, cfg[role])
            for widget in widgets:
                widget.bind("<Button-1>", lambda e, r=role: _select(r))

        _refresh()
        self._e_user.focus_set()

    def _set_expiry_days(self, days: int):
        self._expiry_var.set((datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d"))

    def _pick_expiry_date(self):
        selected_date = _pick_date(self, "Choisir la date d'expiration")
        if selected_date:
            self._expiry_var.set(selected_date)

    def _create(self):
        user = self._e_user.get().strip()
        pw = self._e_pass.get()
        confirm = self._e_confirm.get()
        role = self._role_var.get()
        expiry = self._expiry_var.get().strip()

        if pw != confirm:
            self._msg.configure(text="Les mots de passe ne correspondent pas.")
            return
        if expiry and not re.match(r"^\d{4}-\d{2}-\d{2}$", expiry):
            self._msg.configure(text="Format de date invalide. Utilisez AAAA-MM-JJ.")
            return

        success, err = create_user(user, pw, role, access_expires_at=expiry)
        if not success:
            self._msg.configure(text=err)
            return
        self.destroy()
        if self.on_created:
            self.on_created()


# ── Dialog : dupliquer un compte ─────────────────────────────────────────────

class _DuplicateAccountDialog(tk.Toplevel):
    def __init__(self, parent, source_user: dict, on_done=None):
        super().__init__(parent)
        self.source_user = source_user
        self.on_done = on_done
        self.title(f"Dupliquer le compte — {source_user['username']}")
        self.geometry("460x360")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, False)
        self.transient(parent)
        self.after(0, self._safe_focus)
        self._build_ui()

    def _safe_focus(self):
        try:
            self.lift()
            self.focus_set()
        except Exception:
            pass

    def _build_ui(self):
        bar = tk.Frame(self, bg=BUTTON_BLUE, height=46)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text=f"🧬  Dupliquer — {self.source_user['username']}", font=("Poppins", 12, "bold"), fg="white", bg=BUTTON_BLUE).pack(side="left", padx=20)

        body = tk.Frame(self, bg=PANEL_BG_COLOR)
        body.pack(fill="both", expand=True, padx=24, pady=18)

        info_txt = f"Le nouveau compte reprendra le rôle {self.source_user['role_label']}"
        if self.source_user.get("access_expires_at"):
            info_txt += f" et l'accès jusqu'au {self.source_user['access_expires_at']}"
        else:
            info_txt += " avec accès illimité"
        tk.Label(body, text=info_txt, font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR, wraplength=390, justify="left").pack(anchor="w", pady=(0, 14))

        def _field(label, attr, show=""):
            tk.Label(body, text=label, font=("Poppins", 10, "bold"), fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(6, 2))
            entry = ctk.CTkEntry(
                body,
                height=34,
                fg_color=INPUT_BG_COLOR,
                text_color=TEXT_COLOR,
                border_color=_BORDER_CLR,
                corner_radius=8,
                font=ctk.CTkFont(family="Poppins", size=12),
                show=show,
            )
            entry.pack(fill="x")
            setattr(self, attr, entry)

        _field("Nouveau nom d'utilisateur", "_e_user")
        _field("Mot de passe", "_e_pass", show="•")
        _field("Confirmer le mot de passe", "_e_confirm", show="•")
        self._msg = tk.Label(body, text="", font=("Poppins", 9, "bold"), fg=BUTTON_RED, bg=PANEL_BG_COLOR)
        self._msg.pack(anchor="w", pady=(10, 0))

        btns = tk.Frame(body, bg=PANEL_BG_COLOR)
        btns.pack(fill="x", pady=(18, 0))
        ctk.CTkButton(
            btns,
            text="Créer le doublon",
            height=38,
            fg_color=BUTTON_BLUE,
            hover_color="#1565C0",
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12, weight="bold"),
            command=self._duplicate,
        ).pack(side="left", expand=True, fill="x", padx=(0, 8))
        ctk.CTkButton(
            btns,
            text="Annuler",
            width=100,
            height=38,
            fg_color="#AAAAAA",
            hover_color="#888888",
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12),
            command=self.destroy,
        ).pack(side="left")

        self._e_user.focus_set()

    def _duplicate(self):
        username = self._e_user.get().strip()
        pw = self._e_pass.get()
        confirm = self._e_confirm.get()
        if pw != confirm:
            self._msg.configure(text="Les mots de passe ne correspondent pas.")
            return
        success, err = duplicate_user(self.source_user["username"], username, pw)
        if not success:
            self._msg.configure(text=err)
            return
        self.destroy()
        if self.on_done:
            self.on_done()


# ── Dialog : réinitialiser MDP (admin) ──────────────────────────────────────

class _AdminResetPasswordDialog(tk.Toplevel):
    def __init__(self, parent, username, on_done=None):
        super().__init__(parent)
        self.username = username
        self.on_done = on_done
        self.title(f"Réinitialiser MDP — {username}")
        self.geometry("420x300")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, False)
        self.transient(parent)
        self.after(0, self._safe_focus)
        self._build_ui()

    def _safe_focus(self):
        try:
            self.lift()
            self.focus_set()
        except Exception:
            pass

    def _build_ui(self):
        bar = tk.Frame(self, bg=BUTTON_ORANGE, height=46)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text=f"🔑  Réinitialiser — {self.username}", font=("Poppins", 12, "bold"), fg="white", bg=BUTTON_ORANGE).pack(side="left", padx=20)

        body = ctk.CTkFrame(self, fg_color=PANEL_BG_COLOR)
        body.pack(fill="both", expand=True, padx=32, pady=20)

        def _field(label, attr):
            ctk.CTkLabel(body, text=label, font=ctk.CTkFont(family="Poppins", size=12, weight="bold"), text_color=TEXT_COLOR).pack(anchor="w", pady=(8, 3))
            entry = ctk.CTkEntry(
                body,
                width=340,
                height=36,
                show="•",
                fg_color=INPUT_BG_COLOR,
                text_color=TEXT_COLOR,
                border_color=_BORDER_CLR,
                corner_radius=8,
                font=ctk.CTkFont(family="Poppins", size=12),
            )
            entry.pack(anchor="w")
            setattr(self, attr, entry)

        _field("Nouveau mot de passe (min. 6 car.)", "_e_new")
        _field("Confirmer", "_e_confirm")
        self._msg = ctk.CTkLabel(body, text="", font=ctk.CTkFont(family="Poppins", size=11), text_color=BUTTON_RED)
        self._msg.pack(pady=(8, 0))
        ctk.CTkButton(
            body,
            text="Valider",
            width=340,
            height=38,
            fg_color=BUTTON_ORANGE,
            hover_color="#E65100",
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=13, weight="bold"),
            command=self._reset,
        ).pack(pady=(10, 0))

    def _reset(self):
        pw = self._e_new.get()
        if pw != self._e_confirm.get():
            self._msg.configure(text="Les mots de passe ne correspondent pas.")
            return
        success, err = change_password(self.username, pw)
        if not success:
            self._msg.configure(text=err)
            return
        self.destroy()
        if self.on_done:
            self.on_done()
