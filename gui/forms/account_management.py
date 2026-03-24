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
)


class AccountManagementWindow(tk.Toplevel):
    """Fenêtre de gestion des comptes — admin seulement."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Gestion des comptes utilisateurs")
        self.geometry("720x540")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, False)
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
        # Titre
        title_bar = tk.Frame(self, bg=BUTTON_RED)
        title_bar.pack(fill="x")
        tk.Label(
            title_bar,
            text="  Gestion des comptes utilisateurs",
            font=("Poppins", 14, "bold"),
            fg="white", bg=BUTTON_RED,
        ).pack(side="left", pady=10)

        # Liste des utilisateurs
        list_frame = tk.Frame(self, bg=PANEL_BG_COLOR)
        list_frame.pack(fill="both", expand=True, padx=20, pady=(16, 8))

        cols = ("username", "role", "created_at", "pw_expires")
        self._tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=9)
        headers = {
            "username": "Utilisateur",
            "role": "Rôle",
            "created_at": "Créé le",
            "pw_expires": "Mot de passe",
        }
        widths = {"username": 180, "role": 130, "created_at": 150, "pw_expires": 160}
        for c in cols:
            self._tree.heading(c, text=headers[c])
            self._tree.column(c, width=widths[c], minwidth=80)

        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._tree.tag_configure("expired", foreground=BUTTON_RED)
        self._tree.tag_configure("ok", foreground=BUTTON_GREEN)

        # Boutons
        btn_frame = tk.Frame(self, bg=PANEL_BG_COLOR)
        btn_frame.pack(fill="x", padx=20, pady=(0, 8))

        def _btn(text, cmd, color, hover=None):
            return tk.Button(
                btn_frame, text=text, command=cmd,
                bg=color, fg="white",
                font=("Poppins", 10, "bold"),
                relief="flat", padx=12, pady=6, cursor="hand2",
            )

        _btn("➕ Nouveau compte", self._create_account, BUTTON_GREEN).pack(side="left", padx=(0, 8))
        self._btn_reset_pw = _btn("🔑 Réinitialiser MDP", self._reset_password, BUTTON_ORANGE)
        self._btn_reset_pw.pack(side="left", padx=(0, 8))
        self._btn_delete = _btn("🗑 Supprimer", self._delete_account, BUTTON_RED)
        self._btn_delete.pack(side="left")

        self._tree.bind("<<TreeviewSelect>>", self._on_select)

    # ── Données ───────────────────────────────────────────────────────────────

    def _load_users(self):
        for item in self._tree.get_children():
            self._tree.delete(item)
        for u in get_users():
            expired = u["is_expired"]
            pw_status = "⚠ Expiré" if expired else "✓ Valide"
            tag = "expired" if expired else "ok"
            role_label = ROLES.get(u["role"], u["role"])
            self._tree.insert("", "end", values=(
                u["username"],
                role_label,
                (u["created_at"] or "")[:10],
                pw_status,
            ), tags=(tag,))

    def _get_selected_username(self):
        sel = self._tree.selection()
        if not sel:
            return None
        return self._tree.item(sel[0], "values")[0]

    def _on_select(self, event=None):
        sel = bool(self._tree.selection())
        state = "normal" if sel else "disabled"
        self._btn_reset_pw.configure(state=state)
        self._btn_delete.configure(state=state)

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
            icon="warning",
            parent=self,
        ):
            return
        success, err = delete_user(username)
        if success:
            messagebox.showinfo("Succès", f"Compte « {username} » supprimé.", parent=self)
            self._load_users()
        else:
            messagebox.showerror("Erreur", err, parent=self)


# ── Dialog : créer un compte ──────────────────────────────────────────────────

class _CreateAccountDialog(tk.Toplevel):

    def __init__(self, parent, on_created=None):
        super().__init__(parent)
        self.on_created = on_created
        self.title("Créer un compte")
        self.geometry("420x380")
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
        ctk.CTkLabel(
            self, text="Nouveau compte utilisateur",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(pady=(20, 16))

        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(padx=32, fill="x")

        def _field(label, attr, show=""):
            ctk.CTkLabel(form, text=label, font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=TEXT_COLOR).pack(anchor="w", pady=(8, 2))
            e = ctk.CTkEntry(form, width=340, height=36,
                             fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                             border_color="#C9DDE3", corner_radius=8,
                             font=ctk.CTkFont(size=13), show=show)
            e.pack(anchor="w")
            setattr(self, attr, e)

        _field("Nom d'utilisateur", "_e_user")
        _field("Mot de passe (min. 6 caractères)", "_e_pass", show="•")
        _field("Confirmer le mot de passe", "_e_confirm", show="•")

        # Rôle
        ctk.CTkLabel(form, text="Rôle", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=TEXT_COLOR).pack(anchor="w", pady=(8, 2))
        self._role_var = tk.StringVar(value="agent")
        role_frame = tk.Frame(form, bg=PANEL_BG_COLOR)
        role_frame.pack(anchor="w")
        for role, desc in ROLES.items():
            tk.Radiobutton(
                role_frame, text=f"{role.capitalize()} — {desc}",
                variable=self._role_var, value=role,
                bg=PANEL_BG_COLOR, fg=TEXT_COLOR,
                font=("Poppins", 10), selectcolor=BUTTON_GREEN,
            ).pack(anchor="w")

        self._msg = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12),
                                  text_color=BUTTON_RED)
        self._msg.pack(pady=(6, 0))

        ctk.CTkButton(
            self, text="Créer le compte",
            width=340, height=38,
            fg_color=BUTTON_GREEN, hover_color="#2E7D32",
            corner_radius=8, font=ctk.CTkFont(size=13, weight="bold"),
            command=self._create,
        ).pack(pady=12)

    def _create(self):
        user = self._e_user.get().strip()
        pw = self._e_pass.get()
        confirm = self._e_confirm.get()
        role = self._role_var.get()

        if pw != confirm:
            self._msg.configure(text="Les mots de passe ne correspondent pas.")
            return
        success, err = create_user(user, pw, role)
        if not success:
            self._msg.configure(text=err)
            return
        messagebox.showinfo("Succès", f"Compte « {user} » créé avec le rôle {role}.",
                            parent=self)
        self.destroy()
        if self.on_created:
            self.on_created()


# ── Dialog : réinitialiser MDP (admin) ────────────────────────────────────────

class _AdminResetPasswordDialog(tk.Toplevel):

    def __init__(self, parent, username, on_done=None):
        super().__init__(parent)
        self.username = username
        self.on_done = on_done
        self.title(f"Réinitialiser MDP — {username}")
        self.geometry("400x280")
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
        ctk.CTkLabel(
            self, text=f"Réinitialiser le mot de passe\nde « {self.username} »",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=TEXT_COLOR, justify="center",
        ).pack(pady=(20, 16))

        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(padx=32, fill="x")

        ctk.CTkLabel(form, text="Nouveau mot de passe (min. 6 car.)",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=TEXT_COLOR).pack(anchor="w", pady=(0, 2))
        self._e_new = ctk.CTkEntry(form, width=320, height=36, show="•",
                                    fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                                    border_color="#C9DDE3", corner_radius=8,
                                    font=ctk.CTkFont(size=13))
        self._e_new.pack(anchor="w")

        ctk.CTkLabel(form, text="Confirmer",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=TEXT_COLOR).pack(anchor="w", pady=(10, 2))
        self._e_confirm = ctk.CTkEntry(form, width=320, height=36, show="•",
                                        fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                                        border_color="#C9DDE3", corner_radius=8,
                                        font=ctk.CTkFont(size=13))
        self._e_confirm.pack(anchor="w")

        self._msg = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12),
                                  text_color=BUTTON_RED)
        self._msg.pack(pady=(8, 0))

        ctk.CTkButton(
            self, text="Valider", width=320, height=38,
            fg_color=BUTTON_ORANGE, hover_color="#E65100",
            corner_radius=8, font=ctk.CTkFont(size=13, weight="bold"),
            command=self._reset,
        ).pack(pady=10)

    def _reset(self):
        pw = self._e_new.get()
        confirm = self._e_confirm.get()
        if pw != confirm:
            self._msg.configure(text="Les mots de passe ne correspondent pas.")
            return
        success, err = change_password(self.username, pw)
        if not success:
            self._msg.configure(text=err)
            return
        messagebox.showinfo("Succès",
                            f"Mot de passe de « {self.username} » réinitialisé.",
                            parent=self)
        self.destroy()
        if self.on_done:
            self.on_done()
