"""
Écran de connexion Lahimena Tours.

Flux :
  1. Si aucun compte → assistant création du premier compte admin.
  2. Sinon → formulaire username/password.
  3. Si mot de passe expiré → forcer changement avant accès.
  4. On success → callback on_login_success(user_dict).
"""

import os
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

from config import (
    BUTTON_BLUE,
    BUTTON_GREEN,
    BUTTON_RED,
    INPUT_BG_COLOR,
    MAIN_BG_COLOR,
    MUTED_TEXT_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
)
from utils.auth_handler import (
    authenticate,
    change_password,
    create_user,
    has_users,
    set_current_user,
)

_BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOGO_PATH = os.path.join(_BASE_DIR, "assets", "logo.png")


# ── Fenêtre principale de login ───────────────────────────────────────────────

class LoginWindow(ctk.CTk):
    """
    Fenêtre de connexion autonome.
    Se ferme et appelle on_login_success(user) après authentification réussie.
    """

    def __init__(self, on_login_success):
        super().__init__()
        self.on_login_success = on_login_success
        self.title("Lahimena Tours — Connexion")
        self.geometry("900x600")
        self.resizable(False, False)
        self.configure(fg_color=MAIN_BG_COLOR)

        # Centrer la fenêtre
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - 900) // 2
        y = (sh - 600) // 2
        self.geometry(f"900x600+{x}+{y}")

        self._build_ui()

        # Premier lancement → créer le compte admin
        if not has_users():
            self.after(200, self._show_first_run_dialog)

    # ── Layout principal ──────────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)  # panneau gauche (déco)
        self.grid_columnconfigure(1, weight=1)  # panneau droite (formulaire)
        self.grid_rowconfigure(0, weight=1)

        # ── Panneau gauche — branding ─────────────────────────────────────
        left = ctk.CTkFrame(self, fg_color=BUTTON_RED, corner_radius=0)
        left.grid(row=0, column=0, sticky="nswe")
        left.grid_rowconfigure(0, weight=1)
        left.grid_columnconfigure(0, weight=1)

        branding = ctk.CTkFrame(left, fg_color="transparent")
        branding.grid(row=0, column=0)

        # Logo
        try:
            logo_img = ctk.CTkImage(
                light_image=__import__("PIL.Image", fromlist=["Image"]).open(LOGO_PATH),
                dark_image=__import__("PIL.Image", fromlist=["Image"]).open(LOGO_PATH),
                size=(140, 140),
            )
            ctk.CTkLabel(branding, image=logo_img, text="").pack(pady=(0, 16))
        except Exception:
            ctk.CTkLabel(
                branding,
                text="🦜",
                font=ctk.CTkFont(size=72),
                text_color="white",
            ).pack(pady=(0, 16))

        ctk.CTkLabel(
            branding,
            text="Lahimena",
            font=ctk.CTkFont(size=36, weight="bold"),
            text_color="white",
        ).pack()
        ctk.CTkLabel(
            branding,
            text="Tours Madagascar",
            font=ctk.CTkFont(size=16),
            text_color="#FFCCCC",
        ).pack(pady=(4, 0))

        ctk.CTkLabel(
            left,
            text="Malagasy authentique",
            font=ctk.CTkFont(size=13, slant="italic"),
            text_color="#FFE0E0",
        ).grid(row=0, column=0, sticky="s", pady=(0, 24))

        # ── Panneau droit — formulaire ────────────────────────────────────
        right = ctk.CTkFrame(self, fg_color=PANEL_BG_COLOR, corner_radius=0)
        right.grid(row=0, column=1, sticky="nswe")
        right.grid_rowconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)

        form_box = ctk.CTkFrame(right, fg_color="transparent")
        form_box.grid(row=0, column=0, padx=50)

        ctk.CTkLabel(
            form_box,
            text="Connexion",
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", pady=(0, 4))
        ctk.CTkLabel(
            form_box,
            text="Entrez vos identifiants pour accéder à l'application.",
            font=ctk.CTkFont(size=12),
            text_color=MUTED_TEXT_COLOR,
        ).pack(anchor="w", pady=(0, 28))

        # Username
        ctk.CTkLabel(
            form_box, text="Nom d'utilisateur",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w")
        self.entry_username = ctk.CTkEntry(
            form_box, width=320, height=40,
            placeholder_text="Votre identifiant",
            fg_color=INPUT_BG_COLOR,
            text_color=TEXT_COLOR,
            border_color="#C9DDE3",
            corner_radius=10,
            font=ctk.CTkFont(size=13),
        )
        self.entry_username.pack(anchor="w", pady=(4, 16))

        # Password
        ctk.CTkLabel(
            form_box, text="Mot de passe",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w")
        self.entry_password = ctk.CTkEntry(
            form_box, width=320, height=40,
            placeholder_text="••••••••",
            show="•",
            fg_color=INPUT_BG_COLOR,
            text_color=TEXT_COLOR,
            border_color="#C9DDE3",
            corner_radius=10,
            font=ctk.CTkFont(size=13),
        )
        self.entry_password.pack(anchor="w", pady=(4, 8))
        self.entry_password.bind("<Return>", lambda e: self._do_login())

        # Show/hide password + Remember me
        opts_row = ctk.CTkFrame(form_box, fg_color="transparent")
        opts_row.pack(fill="x", pady=(0, 20))

        self._show_pw_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            opts_row,
            text="Afficher le mot de passe",
            variable=self._show_pw_var,
            command=self._toggle_password_visibility,
            font=ctk.CTkFont(size=12),
            text_color=MUTED_TEXT_COLOR,
            fg_color=BUTTON_BLUE,
            hover_color=BUTTON_BLUE,
            border_color="#C9DDE3",
            width=20, height=20,
        ).pack(side="left")

        # Message d'erreur / info
        self._msg_var = tk.StringVar(value="")
        self._msg_label = ctk.CTkLabel(
            form_box,
            textvariable=self._msg_var,
            font=ctk.CTkFont(size=12),
            text_color=BUTTON_RED,
            wraplength=320,
        )
        self._msg_label.pack(anchor="w", pady=(0, 10))

        # Bouton LOGIN
        self.btn_login = ctk.CTkButton(
            form_box,
            text="LOGIN",
            width=320, height=44,
            fg_color=BUTTON_RED,
            hover_color="#B71C1C",
            corner_radius=10,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._do_login,
        )
        self.btn_login.pack(anchor="w", pady=(0, 16))

        # Mot de passe oublié
        forgot_lbl = ctk.CTkLabel(
            form_box,
            text="Mot de passe oublié ?",
            font=ctk.CTkFont(size=12, underline=True),
            text_color=BUTTON_BLUE,
            cursor="hand2",
        )
        forgot_lbl.pack(anchor="w")
        forgot_lbl.bind("<Button-1>", lambda e: self._show_forgot_password())

        self.entry_username.focus_set()

    # ── Actions ───────────────────────────────────────────────────────────────

    def _toggle_password_visibility(self):
        if self._show_pw_var.get():
            self.entry_password.configure(show="")
        else:
            self.entry_password.configure(show="•")

    def _set_msg(self, text, color=None):
        self._msg_var.set(text)
        if color:
            self._msg_label.configure(text_color=color)

    def _do_login(self):
        username = self.entry_username.get().strip()
        password = self.entry_password.get()

        if not username or not password:
            self._set_msg("Veuillez remplir tous les champs.", BUTTON_RED)
            return

        self.btn_login.configure(state="disabled", text="Vérification…")
        self.update()

        success, user, message = authenticate(username, password)

        self.btn_login.configure(state="normal", text="LOGIN")

        if not success:
            self._set_msg(message, BUTTON_RED)
            self.entry_password.delete(0, tk.END)
            return

        if message == "expired":
            self._set_msg(
                "⚠ Votre mot de passe a expiré (3 mois). Veuillez en choisir un nouveau.",
                "#E65100",
            )
            self.after(200, lambda: self._show_change_password_dialog(user, forced=True))
            return

        self._login_success(user)

    def _login_success(self, user):
        set_current_user(user)
        self.withdraw()
        self.after(100, lambda: (self.destroy(), self.on_login_success(user)))

    # ── Dialogues secondaires ─────────────────────────────────────────────────

    def _show_first_run_dialog(self):
        """Premier lancement — créer le compte administrateur."""
        _FirstRunDialog(self)

    def _show_forgot_password(self):
        """Invite à contacter l'admin — pas de reset automatique."""
        messagebox.showinfo(
            "Mot de passe oublié",
            "Veuillez contacter votre administrateur Lahimena Tours\n"
            "pour réinitialiser votre mot de passe.",
            parent=self,
        )

    def _show_change_password_dialog(self, user, forced=False):
        """Dialog pour changer le mot de passe (forcé si expiré)."""
        _ChangePasswordDialog(self, user["username"], forced=forced,
                              on_success=lambda: self._login_success(user))


# ── Dialog : premier lancement ────────────────────────────────────────────────

class _FirstRunDialog(tk.Toplevel):
    """Crée le premier compte administrateur au premier lancement."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Configuration initiale — Créer le compte administrateur")
        self.geometry("460x420")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # non fermable

        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(
            self,
            text="Bienvenue sur Lahimena Tours",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(pady=(24, 4))

        ctk.CTkLabel(
            self,
            text="Aucun compte n'existe encore.\nCréez le premier compte administrateur.",
            font=ctk.CTkFont(size=13),
            text_color=MUTED_TEXT_COLOR,
            justify="center",
        ).pack(pady=(0, 20))

        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(padx=40, fill="x")

        def _field(label, attr, show=""):
            ctk.CTkLabel(form, text=label, font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=TEXT_COLOR).pack(anchor="w", pady=(8, 2))
            e = ctk.CTkEntry(form, width=360, height=36,
                             fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                             border_color="#C9DDE3", corner_radius=8,
                             font=ctk.CTkFont(size=13), show=show)
            e.pack(anchor="w")
            setattr(self, attr, e)

        _field("Nom d'utilisateur", "_e_user")
        _field("Mot de passe (min. 6 caractères)", "_e_pass", show="•")
        _field("Confirmer le mot de passe", "_e_confirm", show="•")

        self._msg = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12),
                                  text_color=BUTTON_RED)
        self._msg.pack(pady=(8, 0))

        ctk.CTkButton(
            self, text="Créer le compte administrateur",
            width=360, height=40,
            fg_color=BUTTON_GREEN, hover_color="#2E7D32",
            corner_radius=10, font=ctk.CTkFont(size=13, weight="bold"),
            command=self._create,
        ).pack(pady=16)

        self._e_user.focus_set()

    def _create(self):
        username = self._e_user.get().strip()
        password = self._e_pass.get()
        confirm = self._e_confirm.get()

        if password != confirm:
            self._msg.configure(text="Les mots de passe ne correspondent pas.")
            return

        success, err = create_user(username, password, "admin")
        if not success:
            self._msg.configure(text=err)
            return

        messagebox.showinfo(
            "Compte créé",
            f"Compte administrateur « {username} » créé.\n"
            "Vous pouvez maintenant vous connecter.",
            parent=self,
        )
        self.destroy()


# ── Dialog : changer le mot de passe ─────────────────────────────────────────

class _ChangePasswordDialog(tk.Toplevel):
    """
    Dialog de changement de mot de passe.
    forced=True  → titre différent + pas de bouton Annuler.
    """

    def __init__(self, parent, username, forced=False, on_success=None):
        super().__init__(parent)
        self.username = username
        self.forced = forced
        self.on_success = on_success

        title = "Mot de passe expiré — Changement obligatoire" if forced else "Changer le mot de passe"
        self.title(title)
        self.geometry("420x360")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        if forced:
            self.protocol("WM_DELETE_WINDOW", lambda: None)

        self._build_ui()

    def _build_ui(self):
        msg = (
            "Votre mot de passe a expiré (3 mois).\nChoisissez un nouveau mot de passe pour continuer."
            if self.forced
            else f"Changer le mot de passe de « {self.username} »."
        )
        ctk.CTkLabel(
            self, text=msg,
            font=ctk.CTkFont(size=13),
            text_color=TEXT_COLOR, wraplength=360, justify="center",
        ).pack(pady=(24, 16))

        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(padx=32, fill="x")

        def _field(label, attr, show="•"):
            ctk.CTkLabel(form, text=label, font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=TEXT_COLOR).pack(anchor="w", pady=(8, 2))
            e = ctk.CTkEntry(form, width=340, height=36,
                             fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                             border_color="#C9DDE3", corner_radius=8,
                             font=ctk.CTkFont(size=13), show=show)
            e.pack(anchor="w")
            setattr(self, attr, e)

        _field("Nouveau mot de passe (min. 6 caractères)", "_e_new")
        _field("Confirmer le nouveau mot de passe", "_e_confirm")

        self._msg = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12),
                                  text_color=BUTTON_RED)
        self._msg.pack(pady=(8, 0))

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(pady=12)

        ctk.CTkButton(
            btns, text="Valider",
            width=160, height=38,
            fg_color=BUTTON_GREEN, hover_color="#2E7D32",
            corner_radius=8, font=ctk.CTkFont(size=13, weight="bold"),
            command=self._change,
        ).pack(side="left", padx=(0, 10))

        if not self.forced:
            ctk.CTkButton(
                btns, text="Annuler",
                width=120, height=38,
                fg_color="#AAAAAA", hover_color="#888888",
                corner_radius=8,
                command=self.destroy,
            ).pack(side="left")

        self._e_new.focus_set()

    def _change(self):
        new_pw = self._e_new.get()
        confirm = self._e_confirm.get()
        if new_pw != confirm:
            self._msg.configure(text="Les mots de passe ne correspondent pas.")
            return
        success, err = change_password(self.username, new_pw)
        if not success:
            self._msg.configure(text=err)
            return
        messagebox.showinfo("Succès", "Mot de passe modifié avec succès !", parent=self)
        self.destroy()
        if self.on_success:
            self.on_success()
