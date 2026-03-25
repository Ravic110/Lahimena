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
        self._btn_del.pack(side="left")

        self._btn_reset.configure(state="disabled")
        self._btn_del.configure(state="disabled")

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
            placeholder_text="AAAA-MM-JJ",
            height=34, fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
            border_color=_BORDER_CLR, corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12),
        )
        self._expiry_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            date_row, text="Appliquer", width=90, height=34,
            fg_color=BUTTON_BLUE, hover_color="#1565C0",
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=11, weight="bold"),
            command=self._apply_expiry,
        ).pack(side="left")

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
        state = "normal" if enabled else "disabled"
        self._expiry_entry.configure(state=state)
        self._btn_suspend.configure(state=state)

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


# ── Dialog : créer un compte ──────────────────────────────────────────────────

class _CreateAccountDialog(tk.Toplevel):

    def __init__(self, parent, on_created=None):
        super().__init__(parent)
        self.on_created = on_created
        self.title("Nouveau compte")
        self.geometry("500x530")
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
                     fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(8, 2))
            e = ctk.CTkEntry(body, height=36,
                             fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                             border_color=_BORDER_CLR, corner_radius=8,
                             font=ctk.CTkFont(family="Poppins", size=12), show=show)
            e.pack(fill="x")
            setattr(self, attr, e)

        _field("Nom d'utilisateur", "_e_user")
        _field("Mot de passe  (min. 6 caractères)", "_e_pass", show="•")
        _field("Confirmer le mot de passe", "_e_confirm", show="•")

        # Date d'expiration d'accès (optionnel)
        tk.Label(body, text="Date d'expiration d'accès  (optionnel — AAAA-MM-JJ)",
                 font=("Poppins", 10, "bold"),
                 fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(anchor="w", pady=(8, 2))
        from datetime import datetime, timedelta
        _default_expiry = (datetime.now() + timedelta(days=90)).strftime("%Y-%m-%d")
        self._e_expiry = ctk.CTkEntry(
            body, height=36,
            fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
            border_color=_BORDER_CLR, corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=12),
        )
        self._e_expiry.insert(0, _default_expiry)
        self._e_expiry.pack(fill="x")

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
            card = tk.Frame(role_box, bg="#E8F4F8", cursor="hand2")
            pad = (0, 4) if col < 2 else (4, 0)
            card.grid(row=0, column=col, sticky="nsew", padx=pad)
            top = tk.Frame(card, bg="#E8F4F8")
            top.pack(fill="x", padx=8, pady=(8, 2))
            icon_l = tk.Label(top, text=cfg["icon"], font=("Poppins", 16),
                              bg="#E8F4F8", fg=TEXT_COLOR)
            icon_l.pack(side="left")
            name_l = tk.Label(top, text=role.capitalize(),
                              font=("Poppins", 10, "bold"),
                              bg="#E8F4F8", fg=TEXT_COLOR)
            name_l.pack(side="left", padx=(6, 0))
            desc_l = tk.Label(card, text=cfg["desc"], font=("Poppins", 8),
                              bg="#E8F4F8", fg=MUTED_TEXT_COLOR,
                              wraplength=120, justify="left")
            desc_l.pack(anchor="w", padx=8, pady=(0, 8))
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

    def _create(self):
        import re
        user    = self._e_user.get().strip()
        pw      = self._e_pass.get()
        confirm = self._e_confirm.get()
        role    = self._role_var.get()
        expiry  = self._e_expiry.get().strip()

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
