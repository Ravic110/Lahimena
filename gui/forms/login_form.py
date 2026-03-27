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
        self.geometry("580x720")
        self.resizable(False, False)

        # Centrer la fenêtre
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - 580) // 2
        y = (sh - 720) // 2
        self.geometry(f"580x720+{x}+{y}")

        self._build_ui()

        # Premier lancement → créer le compte admin
        if not has_users():
            self.after(200, self._show_first_run_dialog)

    # ── Layout principal ──────────────────────────────────────────────────────

    def _build_ui(self):
        _TEAL        = "#0D7A87"
        _TEAL_DARK   = "#0A606B"
        _RED         = "#C62828"
        _RED_DK      = "#A31515"
        _BG          = "#F0F2F5"
        _BTN_DK      = "#1C2233"
        _BTN_HOVER   = "#2E3A55"
        _TAB_W       = 62
        _CARD_W      = 360
        _CORNER_R    = 20
        _BORDER_NRM  = "#7BB8C2"
        _BORDER_FOC  = "#FFFFFF"
        _PH_COLOR    = "#8BBCC4"
        _TEXT_COLOR  = "#1C2A2E"
        _RED_LEFT_SHIFT = 50
        _RED_VERTICAL_SHIFT = 30

        self.configure(fg_color=_BG)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        outer = ctk.CTkFrame(self, fg_color="transparent")
        outer.grid(row=0, column=0)

        # ── Logo ──────────────────────────────────────────────────────────
        logo_area = ctk.CTkFrame(outer, fg_color="transparent")
        logo_area.pack(pady=(22, 18))

        try:
            from PIL import Image
            logo_img = ctk.CTkImage(
                light_image=Image.open(LOGO_PATH),
                size=(100, 100),
            )
            self._logo_img_ref = logo_img
            ctk.CTkLabel(logo_area, image=logo_img, text="").pack()
        except Exception:
            ctk.CTkLabel(logo_area, text="🦜",
                         font=ctk.CTkFont(size=60),
                         text_color=_RED).pack()

        ctk.CTkLabel(logo_area, text="Lahimena",
                     font=ctk.CTkFont(family="Poppins", size=38, weight="bold"),
                     text_color=_RED).pack(pady=(6, 0))
        ctk.CTkLabel(logo_area, text="Tours Madagascar",
                     font=ctk.CTkFont(family="Poppins", size=12),
                     text_color="#AAAAAA").pack(pady=(2, 0))

        # ── Carte [onglet-rouge | carte-teal | onglet-rouge] ──────────────
        card_row = tk.Frame(outer, bg=_BG)
        card_row.pack(pady=(0, 22))

        # Canvas pour les onglets latéraux avec coins arrondis
        def _draw_tab(canvas, side):
            canvas.delete("all")
            w = canvas.winfo_width()
            h = canvas.winfo_height()
            if w < 2 or h < 2:
                return
            r = _CORNER_R

            # Polygone lisse : coins arrondis d'un seul côté, droits de l'autre
            if side == "left":
                pts = [
                    w, 0,  w, 0,   # coin haut-droit : dur (répété)
                    r, 0,           # haut, approche arc gauche
                    0, 0,           # point de contrôle arc haut-gauche
                    0, r,           # gauche, après arc haut
                    0, h - r,       # gauche, avant arc bas
                    0, h,           # point de contrôle arc bas-gauche
                    r, h,           # bas, après arc gauche
                    w, h,  w, h,   # coin bas-droit : dur (répété)
                ]
            else:
                pts = [
                    0, 0,  0, 0,   # coin haut-gauche : dur (répété)
                    w - r, 0,       # haut, approche arc droit
                    w, 0,           # point de contrôle arc haut-droit
                    w, r,           # droit, après arc haut
                    w, h - r,       # droit, avant arc bas
                    w, h,           # point de contrôle arc bas-droit
                    w - r, h,       # bas, après arc droit
                    0, h,  0, h,   # coin bas-gauche : dur (répété)
                ]
            canvas.create_polygon(pts, smooth=True,
                                  fill=_RED, outline=_RED, width=0)

        # Le canvas gauche est élargi de _RED_LEFT_SHIFT pour déborder à gauche
        left_cv = tk.Canvas(card_row, width=_TAB_W + _RED_LEFT_SHIFT, bg=_BG,
                            highlightthickness=0, bd=0)
        left_cv.pack(side="left", fill="y", pady=_RED_VERTICAL_SHIFT)
        left_cv.bind("<Configure>", lambda e, c=left_cv: _draw_tab(c, "left"))

        # Carte sans pady → sa hauteur naturelle dépasse les canvas de 30px en haut et en bas
        card = tk.Frame(card_row, bg=_TEAL)
        card.pack(side="left")

        right_cv = tk.Canvas(card_row, width=_TAB_W, bg=_BG,
                             highlightthickness=0, bd=0)
        right_cv.pack(side="left", fill="y", pady=_RED_VERTICAL_SHIFT)
        right_cv.bind("<Configure>", lambda e, c=right_cv: _draw_tab(c, "right"))

        # ── Formulaire ────────────────────────────────────────────────────
        form = tk.Frame(card, bg=_TEAL)
        form.pack(padx=32, pady=(90, 90))

        def _field_row(icon_char, hint, attr, show="", with_eye=False):

            box = ctk.CTkFrame(form,
                               fg_color="#F4F9FA",
                               bg_color=_TEAL,
                               corner_radius=8,
                               border_width=1,
                               border_color=_BORDER_NRM)
            box.pack(fill="x", pady=(0, 16))

            ctk.CTkLabel(box, text=icon_char,
                         fg_color="transparent",
                         font=ctk.CTkFont(size=13),
                         text_color=_PH_COLOR).pack(side="left", padx=(14, 0))

            e = tk.Entry(
                box,
                bg="#F4F9FA", fg=_PH_COLOR,
                font=("Poppins", 11),
                bd=0, highlightthickness=0,
                insertbackground=_TEXT_COLOR,
                show="",
            )
            e.pack(side="left", fill="x", expand=True, ipady=7,
                   padx=(8, 0 if with_eye else 14))

            e._hint      = hint
            e._real_show = show
            e._is_ph     = True
            e.insert(0, hint)

            def _focus_in(ev, entry=e):
                if entry._is_ph:
                    entry.delete(0, "end")
                    entry.config(fg=_TEXT_COLOR, show=entry._real_show,
                                 bg="white")
                    entry._is_ph = False
                box.configure(border_color=_BORDER_FOC, fg_color="white")

            def _focus_out(ev, entry=e):
                if not entry.get():
                    entry.config(show="", fg=_PH_COLOR, bg="#F4F9FA")
                    entry.insert(0, entry._hint)
                    entry._is_ph = True
                box.configure(border_color=_BORDER_NRM, fg_color="#F4F9FA")

            e.bind("<FocusIn>",  _focus_in)
            e.bind("<FocusOut>", _focus_out)

            if with_eye:
                _vis = [False]
                _eye_ref = [None]

                def _toggle(_vis=_vis, entry=e):
                    if entry._is_ph:
                        return
                    _vis[0] = not _vis[0]
                    entry.config(show="" if _vis[0] else "•")
                    _eye_ref[0].configure(text="🙈" if _vis[0] else "👁")

                eye_lbl = ctk.CTkLabel(
                    box, text="👁",
                    fg_color="transparent",
                    font=ctk.CTkFont(size=14),
                    text_color="#AAAAAA",
                    cursor="hand2",
                )
                eye_lbl.pack(side="left", padx=(2, 12))
                eye_lbl.bind("<Button-1>", lambda ev: _toggle())
                _eye_ref[0] = eye_lbl

            setattr(self, attr, e)
            return e

        _field_row("👤", "Nom d'utilisateur", "entry_username")
        _field_row("🔒", "Mot de passe", "entry_password",
                   show="•", with_eye=True)
        self.entry_password.bind("<Return>", lambda e: self._do_login())

        # ── Ligne options ─────────────────────────────────────────────────
        opts = ctk.CTkFrame(form, fg_color="transparent", bg_color=_TEAL)
        opts.pack(fill="x", pady=(4, 14))

        self._remember_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            opts,
            text="Se souvenir de moi",
            variable=self._remember_var,
            onvalue=True, offvalue=False,
            fg_color=_RED,
            hover_color=_RED_DK,
            border_color="#A8D8DF",
            checkmark_color="white",
            text_color="#A8D8DF",
            font=ctk.CTkFont(family="Poppins", size=9),
            bg_color="transparent",
            corner_radius=4,
            checkbox_width=16, checkbox_height=16,
        ).pack(side="left")

        forgot = tk.Label(opts, text="Mot de passe oublié ?",
                          bg=_TEAL, fg="#A8D8DF",
                          font=("Poppins", 9, "underline"),
                          cursor="hand2")
        forgot.pack(side="right")
        forgot.bind("<Button-1>", lambda e: self._show_forgot_password())

        # ── Message d'erreur ──────────────────────────────────────────────
        self._msg_var = tk.StringVar(value="")
        self._msg_label = ctk.CTkLabel(
            form,
            textvariable=self._msg_var,
            font=ctk.CTkFont(family="Poppins", size=10),
            text_color="#FFB3B3",
            fg_color="transparent",
            wraplength=_CARD_W - 10,
        )
        self._msg_label.pack(pady=(0, 8))

        # ── Bouton SE CONNECTER ───────────────────────────────────────────
        self.btn_login = ctk.CTkButton(
            form,
            text="SE CONNECTER",
            width=_CARD_W - 4,
            height=44,
            fg_color=_BTN_DK,
            hover_color=_BTN_HOVER,
            corner_radius=8,
            font=ctk.CTkFont(family="Poppins", size=13, weight="bold"),
            text_color="white",
            command=self._do_login,
        )
        self.btn_login.pack()

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

    def _field_value(self, entry):
        """Retourne la valeur réelle du champ (vide si placeholder actif)."""
        return "" if getattr(entry, "_is_ph", False) else entry.get()

    def _reset_password_field(self):
        """Vide le champ mot de passe et restaure le placeholder."""
        e = self.entry_password
        e.delete(0, "end")
        e.config(show="")
        e.insert(0, e._hint)
        e.config(fg="#BBBBBB")
        e._is_ph = True

    def _do_login(self):
        username = self._field_value(self.entry_username).strip()
        password = self._field_value(self.entry_password)

        if not username or not password:
            self._set_msg("Veuillez remplir tous les champs.", BUTTON_RED)
            return

        self.btn_login.configure(state="disabled", text="Vérification…")
        self.update()

        success, user, message = authenticate(username, password)

        self.btn_login.configure(state="normal", text="SE CONNECTER")

        if not success:
            if message == "suspended":
                self._set_msg("🚫 Ce compte est suspendu. Contactez votre administrateur.", BUTTON_RED)
            elif message == "access_expired":
                self._set_msg("⏱ La durée d'accès de ce compte a expiré. Contactez votre administrateur.", BUTTON_RED)
            else:
                self._set_msg(message, BUTTON_RED)
            self._reset_password_field()
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
