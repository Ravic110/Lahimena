"""
Interface de gestion des comptes utilisateurs (admin uniquement).
Accessible depuis le topbar quand l'utilisateur est admin.
"""

import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    BUTTON_BLUE,
    BUTTON_GREEN,
    BUTTON_ORANGE,
    BUTTON_RED,
    INPUT_BG_COLOR,
    MAIN_BG_COLOR,
    MUTED_TEXT_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
)
from gui.forms.client_form import CalendarDialog
from utils.auth_handler import (
    ROLES,
    change_password,
    create_user,
    delete_user,
    get_users,
    reactivate_user,
    set_access_expiry,
    suspend_user,
)

_ACCENT     = "#C0392B"
_HEADER_BG  = "#F0F6F8"
_ROW_ALT    = "#F5FBFC"
_BORDER_CLR = "#D8E8ED"
_FONT_BOLD  = ("Poppins", 10, "bold")
_FONT_BODY  = ("Poppins", 10)

_STATUS_COLOR = {
    "active":         "#1B7A3E",
    "pw_expired":     "#E65100",
    "suspended":      BUTTON_RED,
    "access_expired": "#7B3F00",
}
_STATUS_LABEL = {
    "active":         "✓ Actif",
    "pw_expired":     "⚠ MDP expiré",
    "suspended":      "🚫 Suspendu",
    "access_expired": "⏱ Accès expiré",
}


class AccountManagementWindow(tk.Toplevel):
    """Fenêtre de gestion des comptes — admin seulement."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Gestion des comptes")
        self.geometry("820x580")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(True, True)
        self.minsize(760, 500)
        self.transient(parent)
        self.after(0, self._safe_grab)
        self._build_ui()
        self._load_users()

    def _safe_grab(self):
        try:
            self.wait_visibility()
            self.grab_set()
        except Exception:
            pass

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Barre de titre
        title_bar = tk.Frame(self, bg=_ACCENT, height=52)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)
        tk.Label(title_bar, text="👤  Gestion des comptes utilisateurs",
                 font=("Poppins", 13, "bold"), fg="white", bg=_ACCENT,
                 ).pack(side="left", padx=20)
        tk.Button(title_bar, text="✕", bg=_ACCENT, fg="white",
                  font=("Poppins", 12, "bold"), relief="flat",
                  cursor="hand2", bd=0, padx=12,
                  command=self.destroy).pack(side="right")

        # Corps principal — deux colonnes
        main = tk.Frame(self, bg=PANEL_BG_COLOR)
        main.pack(fill="both", expand=True, padx=20, pady=16)
        main.columnconfigure(0, weight=3)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        # ── Colonne gauche : liste ────────────────────────────────────────
        left = tk.Frame(main, bg=PANEL_BG_COLOR)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        tk.Label(left, text="Comptes enregistrés",
                 font=_FONT_BOLD, fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR,
                 ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        # Style treeview
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Acc.Treeview",
                        background=PANEL_BG_COLOR,
                        fieldbackground=PANEL_BG_COLOR,
                        rowheight=30, font=_FONT_BODY, borderwidth=0)
        style.configure("Acc.Treeview.Heading",
                        background=_HEADER_BG, foreground=TEXT_COLOR,
                        font=_FONT_BOLD, relief="flat", padding=(6, 5))
        style.map("Acc.Treeview",
                  background=[("selected", "#D6EAF0")],
                  foreground=[("selected", TEXT_COLOR)])
        style.layout("Acc.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

        tree_wrap = tk.Frame(left, bg=_BORDER_CLR, bd=1)
        tree_wrap.grid(row=1, column=0, sticky="nsew")

        cols = ("username", "role", "status", "access_until")
        self._tree = ttk.Treeview(tree_wrap, columns=cols,
                                   show="headings", height=12,
                                   style="Acc.Treeview")
        hdrs   = {"username": "Utilisateur", "role": "Rôle",
                  "status": "Statut", "access_until": "Accès jusqu'au"}
        widths = {"username": 140, "role": 130, "status": 120, "access_until": 110}
        for c in cols:
            self._tree.heading(c, text=hdrs[c])
            self._tree.column(c, width=widths[c], minwidth=60,
                              anchor="w" if c in ("username", "role") else "center")

        vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        for s, color in _STATUS_COLOR.items():
            self._tree.tag_configure(s, foreground=color)
        self._tree.tag_configure("alt", background=_ROW_ALT)

        # Boutons sous la liste
        btn_row = tk.Frame(left, bg=PANEL_BG_COLOR)
        btn_row.grid(row=2, column=0, sticky="ew", pady=(10, 0))

        def _ctk(text, cmd, color, hover, w=None):
            kw = dict(text=text, command=cmd, fg_color=color,
                      hover_color=hover, corner_radius=8, height=34,
                      font=ctk.CTkFont(family="Poppins", size=11, weight="bold"))
            if w:
                kw["width"] = w
            return ctk.CTkButton(btn_row, **kw)

        _ctk("➕ Nouveau compte", self._create_account, BUTTON_GREEN, "#1B5E20", 160).pack(side="left", padx=(0, 6))
        self._btn_reset = _ctk("🔑 MDP", self._reset_password, BUTTON_ORANGE, "#E65100", 80)
        self._btn_reset.pack(side="left", padx=(0, 6))
        self._btn_del   = _ctk("🗑 Supprimer", self._delete_account, BUTTON_RED, "#7F0000", 110)
        self._btn_del.pack(side="left", padx=(0, 6))
        self._btn_history = _ctk("📋 Historique", self._show_history, BUTTON_BLUE, "#1565C0", 120)
        self._btn_history.pack(side="left")

        self._btn_reset.configure(state="disabled")
        self._btn_del.configure(state="disabled")
        self._btn_history.configure(state="disabled")

        self._tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── Colonne droite : panneau de gestion des accès ─────────────────
        right = tk.Frame(main, bg=PANEL_BG_COLOR)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)

        tk.Label(right, text="Gestion des accès",
                 font=_FONT_BOLD, fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR,
                 ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        panel = tk.Frame(right, bg=_HEADER_BG,
                         highlightbackground=_BORDER_CLR, highlightthickness=1)
        panel.grid(row=1, column=0, sticky="nsew")
        right.rowconfigure(1, weight=1)

        # Utilisateur sélectionné
        self._sel_lbl = tk.Label(panel, text="— Aucun utilisateur sélectionné —",
                                  font=("Poppins", 11, "bold"),
                                  fg=MUTED_TEXT_COLOR, bg=_HEADER_BG,
                                  wraplength=260, justify="center")
        self._sel_lbl.pack(pady=(18, 4))

        self._status_badge = tk.Label(panel, text="",
                                       font=("Poppins", 10, "bold"),
                                       bg=_HEADER_BG, fg=TEXT_COLOR)
        self._status_badge.pack(pady=(0, 14))

        sep = tk.Frame(panel, bg=_BORDER_CLR, height=1)
        sep.pack(fill="x", padx=16, pady=(0, 14))

        # Durée d'accès
        tk.Label(panel, text="Date d'expiration d'accès",
                 font=_FONT_BOLD, fg=TEXT_COLOR, bg=_HEADER_BG,
                 ).pack(anchor="w", padx=16)
        tk.Label(panel, text="Laisser vide = accès illimité",
                 font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=_HEADER_BG,
                 ).pack(anchor="w", padx=16, pady=(0, 4))

        date_row = tk.Frame(panel, bg=_HEADER_BG)
        date_row.pack(fill="x", padx=16, pady=(0, 4))

        self._expiry_var = tk.StringVar()
        self._expiry_entry = ctk.CTkEntry(
            date_row, textvariable=self._expiry_var,
            placeholder_text="Cliquer pour choisir une date",
            height=34, fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
            border_color=_BORDER_CLR, corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12),
            state="readonly",
        )
        self._expiry_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self._expiry_entry._entry.configure(cursor="hand2")
        self._expiry_entry._entry.bind("<Button-1>", lambda e: self._pick_expiry_date())

        self._expiry_apply_btn = ctk.CTkButton(
            date_row, text="Appliquer", width=90, height=34,
            fg_color=BUTTON_BLUE, hover_color="#1565C0",
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=11, weight="bold"),
            command=self._apply_expiry,
        )
        self._expiry_apply_btn.pack(side="left")

        ctk.CTkButton(
            panel, text="✕ Supprimer la limite", height=30, width=200,
            fg_color="transparent", hover_color="#E8F0F4",
            border_color=_BORDER_CLR, border_width=1,
            text_color=MUTED_TEXT_COLOR, corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=10),
            command=self._remove_expiry,
        ).pack(pady=(0, 14))

        sep2 = tk.Frame(panel, bg=_BORDER_CLR, height=1)
        sep2.pack(fill="x", padx=16, pady=(0, 14))

        # Suspension / Réactivation
        tk.Label(panel, text="Suspension du compte",
                 font=_FONT_BOLD, fg=TEXT_COLOR, bg=_HEADER_BG,
                 ).pack(anchor="w", padx=16, pady=(0, 8))

        self._btn_suspend = ctk.CTkButton(
            panel, text="🚫 Suspendre l'accès",
            height=38, fg_color=BUTTON_RED, hover_color="#7F0000",
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12, weight="bold"),
            command=self._toggle_suspend,
        )
        self._btn_suspend.pack(fill="x", padx=16, pady=(0, 6))

        # Message de feedback
        self._panel_msg = tk.Label(panel, text="",
                                    font=("Poppins", 10), fg=BUTTON_GREEN,
                                    bg=_HEADER_BG, wraplength=240)
        self._panel_msg.pack(pady=(6, 0))

        self._set_panel_enabled(False)

    # ── Données ───────────────────────────────────────────────────────────────

    def _load_users(self):
        for item in self._tree.get_children():
            self._tree.delete(item)
        for i, u in enumerate(get_users()):
            status     = u["status"]
            expires    = u.get("access_expires_at", "") or "Illimité"
            role_label = ROLES.get(u["role"], u["role"])
            tags = [status]
            if i % 2 == 1:
                tags.append("alt")
            self._tree.insert("", "end", values=(
                u["username"],
                role_label,
                _STATUS_LABEL.get(status, status),
                expires,
            ), tags=tuple(tags))

    def _get_selected_username(self):
        sel = self._tree.selection()
        if not sel:
            return None
        return self._tree.item(sel[0], "values")[0]

    def _get_selected_user_data(self):
        sel = self._tree.selection()
        if not sel:
            return None
        vals = self._tree.item(sel[0], "values")
        return {"username": vals[0], "status_label": vals[2], "access_until": vals[3]}

    def _on_select(self, event=None):
        sel = bool(self._tree.selection())
        state = "normal" if sel else "disabled"
        self._btn_reset.configure(state=state)
        self._btn_del.configure(state=state)
        self._btn_history.configure(state=state)
        self._set_panel_enabled(sel)
        if sel:
            self._refresh_panel()

    def _refresh_panel(self):
        ud = self._get_selected_user_data()
        if not ud:
            return
        self._sel_lbl.configure(text=ud["username"], fg=TEXT_COLOR)

        # Badge statut
        raw_status = self._tree.item(self._tree.selection()[0], "tags")[0]
        badge_text  = _STATUS_LABEL.get(raw_status, raw_status)
        badge_color = _STATUS_COLOR.get(raw_status, TEXT_COLOR)
        self._status_badge.configure(text=badge_text, fg=badge_color)

        # Date d'expiration
        exp = ud["access_until"]
        self._expiry_var.set("" if exp == "Illimité" else exp)

        # Bouton suspend / réactiver
        if raw_status == "suspended":
            self._btn_suspend.configure(
                text="✅ Réactiver l'accès",
                fg_color=BUTTON_GREEN, hover_color="#1B5E20",
            )
        else:
            self._btn_suspend.configure(
                text="🚫 Suspendre l'accès",
                fg_color=BUTTON_RED, hover_color="#7F0000",
            )
        self._panel_msg.configure(text="")

    def _set_panel_enabled(self, enabled: bool):
        self._expiry_entry.configure(state="readonly" if enabled else "disabled")
        self._expiry_apply_btn.configure(state="normal" if enabled else "disabled")
        self._btn_suspend.configure(state="normal" if enabled else "disabled")

    # ── Actions ───────────────────────────────────────────────────────────────

    def _create_account(self):
        _CreateAccountDialog(self, on_created=self._load_users)

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
            icon="warning", parent=self,
        ):
            return
        success, err = delete_user(username)
        if success:
            messagebox.showinfo("Succès", f"Compte « {username} » supprimé.", parent=self)
            self._load_users()
            self._set_panel_enabled(False)
            self._sel_lbl.configure(text="— Aucun utilisateur sélectionné —",
                                     fg=MUTED_TEXT_COLOR)
            self._status_badge.configure(text="")
        else:
            messagebox.showerror("Erreur", err, parent=self)

    def _pick_expiry_date(self):
        cal = CalendarDialog(self, "Choisir la date d'expiration")
        self.wait_window(cal)
        if cal.selected_date:
            self._expiry_var.set(cal.selected_date.strftime("%Y-%m-%d"))

    def _apply_expiry(self):
        username = self._get_selected_username()
        if not username:
            return
        date_str = self._expiry_var.get().strip()
        if date_str:
            # Validation basique du format
            import re
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_str):
                self._panel_msg.configure(
                    text="Format invalide. Utilisez AAAA-MM-JJ.", fg=BUTTON_RED)
                return
        success, err = set_access_expiry(username, date_str)
        if success:
            self._panel_msg.configure(
                text="✓ Date d'expiration mise à jour.", fg=BUTTON_GREEN)
            self._load_users()
            # Re-sélectionner la même ligne
            for item in self._tree.get_children():
                if self._tree.item(item, "values")[0] == username:
                    self._tree.selection_set(item)
                    self._refresh_panel()
                    break
        else:
            self._panel_msg.configure(text=err, fg=BUTTON_RED)

    def _remove_expiry(self):
        username = self._get_selected_username()
        if not username:
            return
        success, err = set_access_expiry(username, "")
        if success:
            self._expiry_var.set("")
            self._panel_msg.configure(text="✓ Limite supprimée.", fg=BUTTON_GREEN)
            self._load_users()
            for item in self._tree.get_children():
                if self._tree.item(item, "values")[0] == username:
                    self._tree.selection_set(item)
                    self._refresh_panel()
                    break
        else:
            self._panel_msg.configure(text=err, fg=BUTTON_RED)

    def _toggle_suspend(self):
        username = self._get_selected_username()
        if not username:
            return
        sel_tags = self._tree.item(self._tree.selection()[0], "tags")
        is_susp  = "suspended" in sel_tags

        if is_susp:
            ok, err = reactivate_user(username)
            msg_ok  = f"✓ Compte « {username} » réactivé."
        else:
            if not messagebox.askyesno(
                "Confirmation",
                f"Suspendre l'accès de « {username} » ?\nL'utilisateur ne pourra plus se connecter.",
                icon="warning", parent=self,
            ):
                return
            ok, err = suspend_user(username)
            msg_ok  = f"✓ Compte « {username} » suspendu."

        if ok:
            self._panel_msg.configure(text=msg_ok, fg=BUTTON_GREEN)
            self._load_users()
            for item in self._tree.get_children():
                if self._tree.item(item, "values")[0] == username:
                    self._tree.selection_set(item)
                    self._refresh_panel()
                    break
        else:
            self._panel_msg.configure(text=err, fg=BUTTON_RED)

    def _show_history(self):
        username = self._get_selected_username()
        if not username:
            return
        _ActivityHistoryDialog(self, username)


# ── Dialog : historique d'activité ────────────────────────────────────────────

class _ActivityHistoryDialog(tk.Toplevel):
    """
    Historique d'activité complet :
    - Filtres : utilisateur, catégorie, plage de dates, recherche texte libre
    - Code couleur par catégorie
    - Tri par colonne
    - Panneau statistiques
    - Export CSV
    """

    _CAT_COLORS = {
        "auth":    "#1565C0",
        "client":  "#1B7A3E",
        "user":    "#C62828",
        "quota":   "#6A1B9A",
        "invoice": "#E65100",
        "nav":     "#78909C",
    }
    _CAT_LABELS = {
        "":        "Toutes",
        "auth":    "🔐 Connexion",
        "client":  "👥 Clients",
        "user":    "👤 Comptes",
        "quota":   "📄 Cotations",
        "invoice": "💶 Factures",
        "nav":     "🧭 Navigation",
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
        self.after(0, lambda: (self.wait_visibility(), self.grab_set()))
        self._sort_col  = "timestamp"
        self._sort_desc = True
        self._all_entries: list = []
        self._build_ui()

    # ── Construction UI ───────────────────────────────────────────────────────

    def _build_ui(self):
        # Barre de titre
        title_bar = tk.Frame(self, bg=_ACCENT, height=48)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)
        tk.Label(title_bar,
                 text=f"📋  Historique d'activité — {self.username}",
                 font=("Poppins", 12, "bold"), fg="white", bg=_ACCENT,
                 ).pack(side="left", padx=16)
        tk.Button(title_bar, text="✕", bg=_ACCENT, fg="white",
                  font=("Poppins", 11, "bold"), relief="flat",
                  cursor="hand2", bd=0, padx=10,
                  command=self.destroy).pack(side="right")

        # Corps principal : filtre + contenu
        body = tk.Frame(self, bg=PANEL_BG_COLOR)
        body.pack(fill="both", expand=True, padx=14, pady=10)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        # ── Barre de filtres ──────────────────────────────────────────────
        fbar = tk.Frame(body, bg=_HEADER_BG,
                        highlightbackground="#C9DDE3", highlightthickness=1)
        fbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        # Utilisateur
        tk.Label(fbar, text="Afficher :", font=("Poppins", 9, "bold"),
                 fg=TEXT_COLOR, bg=_HEADER_BG).grid(row=0, column=0,
                 padx=(10, 4), pady=8, sticky="w")
        self._filter_var = tk.StringVar(value="user")
        for col_idx, (val, lbl) in enumerate(
            [("user", f"Seulement {self.username}"), ("all", "Tous les utilisateurs")],
            start=1
        ):
            tk.Radiobutton(fbar, text=lbl, value=val,
                           variable=self._filter_var,
                           bg=_HEADER_BG, fg=TEXT_COLOR,
                           selectcolor=BUTTON_BLUE,
                           font=("Poppins", 9),
                           command=self._reload,
                           ).grid(row=0, column=col_idx, padx=(0, 10), pady=8)

        # Séparateur vertical
        tk.Frame(fbar, bg="#C9DDE3", width=1).grid(row=0, column=3,
                 padx=8, pady=4, sticky="ns")

        # Catégorie
        tk.Label(fbar, text="Catégorie :", font=("Poppins", 9, "bold"),
                 fg=TEXT_COLOR, bg=_HEADER_BG).grid(row=0, column=4,
                 padx=(0, 4), pady=8)
        self._cat_var = tk.StringVar(value="")
        cat_cb = ttk.Combobox(fbar, textvariable=self._cat_var, state="readonly",
                               values=list(self._CAT_LABELS.values()), width=16,
                               font=("Poppins", 9))
        cat_cb.grid(row=0, column=5, padx=(0, 8), pady=8)
        cat_cb.bind("<<ComboboxSelected>>", lambda e: self._reload())

        # Dates
        tk.Label(fbar, text="Du :", font=("Poppins", 9, "bold"),
                 fg=TEXT_COLOR, bg=_HEADER_BG).grid(row=0, column=6, padx=(0, 4))
        self._date_from = tk.StringVar()
        tk.Entry(fbar, textvariable=self._date_from, width=11,
                 font=("Poppins", 9)).grid(row=0, column=7, padx=(0, 6))
        tk.Label(fbar, text="Au :", font=("Poppins", 9, "bold"),
                 fg=TEXT_COLOR, bg=_HEADER_BG).grid(row=0, column=8, padx=(0, 4))
        self._date_to = tk.StringVar()
        tk.Entry(fbar, textvariable=self._date_to, width=11,
                 font=("Poppins", 9)).grid(row=0, column=9, padx=(0, 6))
        tk.Button(fbar, text="Appliquer", font=("Poppins", 8, "bold"),
                  fg="white", bg=BUTTON_BLUE, relief="flat", bd=0,
                  padx=8, pady=3, cursor="hand2",
                  command=self._reload).grid(row=0, column=10, padx=(0, 6))
        tk.Button(fbar, text="↺", font=("Poppins", 10),
                  fg=TEXT_COLOR, bg=_HEADER_BG, relief="flat", bd=0,
                  cursor="hand2", command=self._reset_filters,
                  ).grid(row=0, column=11, padx=(0, 6))

        # Recherche
        tk.Frame(fbar, bg="#C9DDE3", width=1).grid(row=0, column=12,
                 padx=8, pady=4, sticky="ns")
        tk.Label(fbar, text="🔍", font=("Poppins", 11),
                 bg=_HEADER_BG).grid(row=0, column=13, padx=(0, 2))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_search())
        tk.Entry(fbar, textvariable=self._search_var, width=18,
                 font=("Poppins", 9)).grid(row=0, column=14, padx=(0, 10))

        # ── Onglets : Historique | Statistiques ───────────────────────────
        tab_bar = tk.Frame(body, bg=PANEL_BG_COLOR)
        tab_bar.grid(row=1, column=0, sticky="ew")

        self._tab_hist_btn = tk.Label(tab_bar, text="Historique",
                                       font=("Poppins", 10, "bold"),
                                       bg=BUTTON_RED, fg="white",
                                       padx=14, pady=4, cursor="hand2")
        self._tab_hist_btn.pack(side="left", padx=(0, 4))
        self._tab_stat_btn = tk.Label(tab_bar, text="Statistiques",
                                       font=("Poppins", 10, "bold"),
                                       bg=BUTTON_BLUE, fg="white",
                                       padx=14, pady=4, cursor="hand2")
        self._tab_stat_btn.pack(side="left")
        self._tab_hist_btn.bind("<Button-1>", lambda e: self._show_tab("hist"))
        self._tab_stat_btn.bind("<Button-1>", lambda e: self._show_tab("stat"))

        # Bouton export
        tk.Button(tab_bar, text="⬇ Exporter CSV",
                  font=("Poppins", 9, "bold"), fg="white", bg=BUTTON_BLUE,
                  relief="flat", bd=0, padx=10, pady=4, cursor="hand2",
                  command=self._export_csv).pack(side="right")

        body.rowconfigure(2, weight=1)

        # ── Panneau historique ────────────────────────────────────────────
        self._hist_frame = tk.Frame(body, bg=PANEL_BG_COLOR)
        self._hist_frame.grid(row=2, column=0, sticky="nsew", pady=(4, 0))
        self._hist_frame.columnconfigure(0, weight=1)
        self._hist_frame.rowconfigure(0, weight=1)
        self._build_treeview(self._hist_frame)

        # Compteur
        self._count_lbl = tk.Label(body, text="",
                                    font=("Poppins", 9), fg=MUTED_TEXT_COLOR,
                                    bg=PANEL_BG_COLOR)
        self._count_lbl.grid(row=3, column=0, sticky="e", pady=(2, 0))

        # ── Panneau statistiques ──────────────────────────────────────────
        self._stat_frame = tk.Frame(body, bg=PANEL_BG_COLOR)
        # Non affiché par défaut

        self._reload()

    def _build_treeview(self, parent):
        style = ttk.Style()
        style.configure("Hist.Treeview",
                        background=PANEL_BG_COLOR,
                        fieldbackground=PANEL_BG_COLOR,
                        font=("Poppins", 9), rowheight=26)
        style.configure("Hist.Treeview.Heading",
                        font=("Poppins", 9, "bold"),
                        background=_HEADER_BG)

        cols = ("timestamp", "username", "role", "label", "details")
        self._tree = ttk.Treeview(parent, columns=cols, show="headings",
                                   style="Hist.Treeview")
        widths = {"timestamp": 140, "username": 100, "role": 80,
                  "label": 160, "details": 0}
        heads  = {"timestamp": "Date / Heure", "username": "Utilisateur",
                  "role": "Rôle", "label": "Action", "details": "Détails"}
        for c in cols:
            self._tree.heading(c, text=heads[c],
                               command=lambda col=c: self._sort_by(col))
            if widths[c]:
                self._tree.column(c, width=widths[c], minwidth=60, stretch=False)
            else:
                self._tree.column(c, stretch=True, minwidth=120)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # Tags couleur par catégorie
        for cat, color in self._CAT_COLORS.items():
            self._tree.tag_configure(cat, foreground=color)
        self._tree.tag_configure("odd", background="#F5F9FB")
        self._tree.tag_configure("auth_alert", foreground="#C62828",
                                  font=("Poppins", 9, "bold"))

    def _build_stats_panel(self, parent):
        from utils.activity_log import get_user_stats
        stats = get_user_stats(self.username)

        for w in parent.winfo_children():
            w.destroy()

        bg = PANEL_BG_COLOR
        # Cartes statistiques
        cards_row = tk.Frame(parent, bg=bg)
        cards_row.pack(fill="x", pady=(4, 12))

        def _stat_card(frame, icon, value, label, color):
            card = tk.Frame(frame, bg=color, padx=14, pady=10)
            card.pack(side="left", padx=(0, 10), fill="y")
            tk.Label(card, text=icon, font=("Poppins", 18), bg=color,
                     fg="white").pack()
            tk.Label(card, text=str(value), font=("Poppins", 20, "bold"),
                     bg=color, fg="white").pack()
            tk.Label(card, text=label, font=("Poppins", 8), bg=color,
                     fg="white").pack()

        _stat_card(cards_row, "🔐", stats["login_count"],   "Connexions",      "#1565C0")
        _stat_card(cards_row, "⚠", stats["failed_logins"], "Échecs login",    "#C62828")
        _stat_card(cards_row, "📋", stats["total_actions"], "Total actions",   "#1B7A3E")
        if stats["brute_force_detected"]:
            _stat_card(cards_row, "🚨", "Alerte", "Brute force détecté", "#7F0000")

        # Dernière activité
        info = tk.Frame(parent, bg=bg)
        info.pack(fill="x", pady=(0, 10))
        for lbl, val in [
            ("Dernière connexion :", stats["last_login"] or "—"),
            ("Dernière action :",   stats["last_action_ts"] or "—"),
        ]:
            row = tk.Frame(info, bg=bg)
            row.pack(anchor="w")
            tk.Label(row, text=lbl, font=("Poppins", 9, "bold"),
                     fg=TEXT_COLOR, bg=bg).pack(side="left", padx=(0, 6))
            tk.Label(row, text=val, font=("Poppins", 9),
                     fg=MUTED_TEXT_COLOR, bg=bg).pack(side="left")

        # Top actions
        if stats["top_actions"]:
            tk.Label(parent, text="Actions les plus fréquentes (30 jours)",
                     font=("Poppins", 10, "bold"), fg=TEXT_COLOR,
                     bg=bg).pack(anchor="w", pady=(4, 6))
            for action_lbl, count in stats["top_actions"]:
                bar_frame = tk.Frame(parent, bg=bg)
                bar_frame.pack(fill="x", pady=(0, 4))
                tk.Label(bar_frame, text=action_lbl, width=30, anchor="w",
                         font=("Poppins", 9), fg=TEXT_COLOR,
                         bg=bg).pack(side="left")
                max_count = stats["top_actions"][0][1] if stats["top_actions"] else 1
                bar_w = max(4, int(200 * count / max_count))
                tk.Frame(bar_frame, bg=BUTTON_BLUE,
                         width=bar_w, height=14).pack(side="left", padx=(4, 6))
                tk.Label(bar_frame, text=str(count), font=("Poppins", 9),
                         fg=MUTED_TEXT_COLOR, bg=bg).pack(side="left")

    # ── Onglets ───────────────────────────────────────────────────────────────

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

    # ── Données ───────────────────────────────────────────────────────────────

    def _get_cat_filter(self) -> str:
        label = self._cat_var.get()
        for k, v in self._CAT_LABELS.items():
            if v == label:
                return k
        return ""

    def _reset_filters(self):
        self._filter_var.set("user")
        self._cat_var.set("")
        self._date_from.set("")
        self._date_to.set("")
        self._search_var.set("")
        self._reload()

    def _reload(self):
        from utils.activity_log import get_activity, get_all_activity
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
            entries = [e for e in entries if any(
                q in str(v).lower() for v in e.values()
            )]
        self._populate_tree(entries)

    def _populate_tree(self, entries: list):
        for item in self._tree.get_children():
            self._tree.delete(item)

        for i, e in enumerate(entries):
            cat   = e.get("category", "")
            action = e.get("action", "")
            tags  = []
            if action == "brute_force_alert":
                tags = ["auth_alert"]
            else:
                if cat:
                    tags.append(cat)
                if i % 2 == 1:
                    tags.append("odd")
            self._tree.insert("", "end", values=(
                e.get("timestamp", ""),
                e.get("username", ""),
                e.get("role", ""),
                e.get("label", e.get("action", "")),
                e.get("details", ""),
            ), tags=tags)

        n = len(entries)
        self._count_lbl.configure(
            text=f"{n} entrée{'s' if n > 1 else ''} affichée{'s' if n > 1 else ''}"
        )

    # ── Tri par colonne ───────────────────────────────────────────────────────

    def _sort_by(self, col: str):
        if self._sort_col == col:
            self._sort_desc = not self._sort_desc
        else:
            self._sort_col  = col
            self._sort_desc = True
        key_map = {"timestamp": "timestamp", "username": "username",
                   "role": "role", "label": "label", "details": "details"}
        key = key_map.get(col, col)
        self._all_entries.sort(key=lambda e: e.get(key, "").lower(),
                               reverse=self._sort_desc)
        arrow = " ▼" if self._sort_desc else " ▲"
        heads = {"timestamp": "Date / Heure", "username": "Utilisateur",
                 "role": "Rôle", "label": "Action", "details": "Détails"}
        for c, h in heads.items():
            self._tree.heading(c, text=h + (arrow if c == col else ""))
        self._apply_search()

    # ── Export CSV ────────────────────────────────────────────────────────────

    def _export_csv(self):
        import csv
        from tkinter import filedialog
        from datetime import datetime as dt

        default = f"historique_{self.username}_{dt.now().strftime('%Y%m%d_%H%M%S')}.csv"
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Tous", "*.*")],
            initialfile=default,
            title="Exporter l'historique",
        )
        if not path:
            return

        entries = self._all_entries
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=["timestamp", "username", "role",
                                "label", "category", "details"],
                    extrasaction="ignore",
                )
                writer.writeheader()
                writer.writerows(entries)
            import subprocess, os
            messagebox.showinfo("Export réussi",
                                f"Fichier exporté :\n{path}", parent=self)
        except Exception as ex:
            messagebox.showerror("Erreur", str(ex), parent=self)


# ── Dialog : créer un compte ──────────────────────────────────────────────────

class _CreateAccountDialog(tk.Toplevel):

    def __init__(self, parent, on_created=None):
        super().__init__(parent)
        self.on_created = on_created
        self.title("Nouveau compte")
        self.geometry("520x620")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, True)
        self.minsize(520, 580)
        self.transient(parent)
        self.after(0, self._safe_grab)
        self._build_ui()

    def _safe_grab(self):
        try:
            self.wait_visibility()
            self.grab_set()
        except Exception:
            pass

    def _build_ui(self):
        # Barre titre
        bar = tk.Frame(self, bg=BUTTON_GREEN, height=46)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="➕  Créer un nouveau compte",
                 font=("Poppins", 12, "bold"), fg="white", bg=BUTTON_GREEN,
                 ).pack(side="left", padx=20)

        # Footer épinglé en bas
        footer = tk.Frame(self, bg=PANEL_BG_COLOR)
        footer.pack(side="bottom", fill="x", padx=24, pady=12)

        self._msg = ctk.CTkLabel(footer, text="",
                                  font=ctk.CTkFont(family="Poppins", size=11),
                                  text_color=BUTTON_RED, wraplength=440)
        self._msg.pack(anchor="w", pady=(0, 6))

        btn_row = tk.Frame(footer, bg=PANEL_BG_COLOR)
        btn_row.pack(fill="x")
        ctk.CTkButton(
            btn_row, text="✓  Créer le compte",
            height=40, fg_color=BUTTON_GREEN, hover_color="#1B5E20",
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=13, weight="bold"),
            command=self._create,
        ).pack(side="left", expand=True, fill="x", padx=(0, 8))
        ctk.CTkButton(
            btn_row, text="Annuler", width=110, height=40,
            fg_color="#AAAAAA", hover_color="#888888", corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12),
            command=self.destroy,
        ).pack(side="left")

        # Corps
        body = tk.Frame(self, bg=PANEL_BG_COLOR)
        body.pack(fill="both", expand=True, padx=24, pady=(12, 0))

        def _field(label, attr, show=""):
            tk.Label(body, text=label, font=("Poppins", 10, "bold"),
                     fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(6, 2))
            e = ctk.CTkEntry(body, height=34,
                             fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                             border_color=_BORDER_CLR, corner_radius=8,
                             font=ctk.CTkFont(family="Poppins", size=12), show=show)
            e.pack(fill="x")
            setattr(self, attr, e)

        _field("Nom d'utilisateur", "_e_user")
        _field("Mot de passe  (min. 6 caractères)", "_e_pass", show="•")
        _field("Confirmer le mot de passe", "_e_confirm", show="•")

        # Date d'expiration d'accès (optionnel)
        tk.Label(body, text="Date d'expiration d'accès  (optionnel)",
                 font=("Poppins", 10, "bold"),
                 fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(8, 2))
        from datetime import datetime, timedelta
        _default_expiry = (datetime.now() + timedelta(days=90)).strftime("%Y-%m-%d")
        self._expiry_var = tk.StringVar(value=_default_expiry)
        self._e_expiry = ctk.CTkEntry(
            body, textvariable=self._expiry_var, height=36,
            fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
            border_color=_BORDER_CLR, corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12),
            placeholder_text="Cliquer pour choisir une date",
            state="readonly",
        )
        self._e_expiry.pack(fill="x")
        self._e_expiry._entry.configure(cursor="hand2")
        self._e_expiry._entry.bind("<Button-1>", lambda e: self._pick_expiry_date())

        # Rôle — 3 cartes
        tk.Label(body, text="Rôle", font=("Poppins", 10, "bold"),
                 fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(12, 6))

        self._role_var = tk.StringVar(value="agent")
        role_box = tk.Frame(body, bg=PANEL_BG_COLOR)
        role_box.pack(fill="x")
        for i in range(3):
            role_box.columnconfigure(i, weight=1)

        _ROLE_CFG = {
            "admin":     {"icon": "🛡", "color": BUTTON_RED,   "desc": "Accès complet"},
            "agent":     {"icon": "👤", "color": BUTTON_BLUE,  "desc": "Clients, cotations"},
            "comptable": {"icon": "📊", "color": "#2E7D32",    "desc": "Comptabilité"},
        }
        self._role_cards = {}

        def _make_card(col, role, cfg):
            card = tk.Frame(role_box, bg="#E8F4F8", cursor="hand2",
                            relief="solid", bd=1, highlightbackground=_BORDER_CLR)
            pad = (0, 4) if col < 2 else (4, 0)
            card.grid(row=0, column=col, sticky="nsew", padx=pad, pady=2)
            top = tk.Frame(card, bg="#E8F4F8")
            top.pack(fill="x", padx=8, pady=(8, 2))
            icon_l = tk.Label(top, text=cfg["icon"], font=("Poppins", 15),
                              bg="#E8F4F8", fg=TEXT_COLOR)
            icon_l.pack(side="left")
            name_l = tk.Label(top, text=role.capitalize(),
                              font=("Poppins", 10, "bold"),
                              bg="#E8F4F8", fg=TEXT_COLOR)
            name_l.pack(side="left", padx=(6, 0))
            desc_l = tk.Label(card, text=cfg["desc"], font=("Poppins", 9),
                              bg="#E8F4F8", fg=MUTED_TEXT_COLOR,
                              wraplength=140, justify="left")
            desc_l.pack(anchor="w", padx=8, pady=(0, 10))
            widgets = [card, top, icon_l, name_l, desc_l]
            self._role_cards[role] = (card, widgets, cfg)
            for w in widgets:
                w.bind("<Button-1>", lambda e, r=role: _select(r))

        def _select(r):
            self._role_var.set(r)
            _refresh()

        def _refresh():
            sel = self._role_var.get()
            for r, (card, widgets, cfg) in self._role_cards.items():
                if r == sel:
                    bg, fg, muted = cfg["color"], "white", "#DDDDDD"
                else:
                    bg, fg, muted = "#E8F4F8", TEXT_COLOR, MUTED_TEXT_COLOR
                for w in widgets:
                    w.configure(bg=bg)
                    if isinstance(w, tk.Label):
                        w.configure(fg=fg)
                widgets[4].configure(fg=muted)

        for col, (role, cfg) in enumerate(_ROLE_CFG.items()):
            _make_card(col, role, cfg)

        _refresh()
        self._e_user.focus_set()

    def _pick_expiry_date(self):
        cal = CalendarDialog(self, "Choisir la date d'expiration")
        self.wait_window(cal)
        if cal.selected_date:
            self._expiry_var.set(cal.selected_date.strftime("%Y-%m-%d"))

    def _create(self):
        import re
        user    = self._e_user.get().strip()
        pw      = self._e_pass.get()
        confirm = self._e_confirm.get()
        role    = self._role_var.get()
        expiry  = self._expiry_var.get().strip()

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
        messagebox.showinfo("Succès", f"Compte « {user} » créé avec le rôle {role}.", parent=self)
        self.destroy()
        if self.on_created:
            self.on_created()


# ── Dialog : réinitialiser MDP (admin) ────────────────────────────────────────

class _AdminResetPasswordDialog(tk.Toplevel):

    def __init__(self, parent, username, on_done=None):
        super().__init__(parent)
        self.username = username
        self.on_done  = on_done
        self.title(f"Réinitialiser MDP — {username}")
        self.geometry("420x300")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, False)
        self.transient(parent)
        self.after(0, self._safe_grab)
        self._build_ui()

    def _safe_grab(self):
        try:
            self.wait_visibility()
            self.grab_set()
        except Exception:
            pass

    def _build_ui(self):
        bar = tk.Frame(self, bg=BUTTON_ORANGE, height=46)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text=f"🔑  Réinitialiser — {self.username}",
                 font=("Poppins", 12, "bold"), fg="white", bg=BUTTON_ORANGE,
                 ).pack(side="left", padx=20)

        body = ctk.CTkFrame(self, fg_color=PANEL_BG_COLOR)
        body.pack(fill="both", expand=True, padx=32, pady=20)

        def _field(label, attr):
            ctk.CTkLabel(body, text=label,
                         font=ctk.CTkFont(family="Poppins", size=12, weight="bold"),
                         text_color=TEXT_COLOR).pack(anchor="w", pady=(8, 3))
            e = ctk.CTkEntry(body, width=340, height=36, show="•",
                             fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                             border_color=_BORDER_CLR, corner_radius=8,
                             font=ctk.CTkFont(family="Poppins", size=12))
            e.pack(anchor="w")
            setattr(self, attr, e)

        _field("Nouveau mot de passe (min. 6 car.)", "_e_new")
        _field("Confirmer", "_e_confirm")

        self._msg = ctk.CTkLabel(body, text="",
                                  font=ctk.CTkFont(family="Poppins", size=11),
                                  text_color=BUTTON_RED)
        self._msg.pack(pady=(8, 0))

        ctk.CTkButton(
            body, text="Valider", width=340, height=38,
            fg_color=BUTTON_ORANGE, hover_color="#E65100", corner_radius=8,
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
        messagebox.showinfo("Succès", f"Mot de passe de « {self.username} » réinitialisé.", parent=self)
        self.destroy()
        if self.on_done:
            self.on_done()
