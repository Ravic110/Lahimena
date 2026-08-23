"""
Lahimena Tours Devis Generation Application

Main entry point for the application
"""

import os
import sys
import tkinter as _tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from gui.ctk_patch import appliquer as _appliquer_correctifs_ctk

# Avant tout import de gui.* : importer ces modules instancie des widgets
# CustomTkinter, qui doivent deja etre traces par les correctifs.
_appliquer_correctifs_ctk()

from config import (  # noqa: E402
    APP_GEOMETRY,
    APP_TITLE,
    APPEARANCE_MODE,
    BUTTON_GREEN,
    BUTTON_GREEN_HOVER,
    BUTTON_RED,
    DEFAULT_COLOR_THEME,
    INPUT_BG_COLOR,
    MAIN_BG_COLOR,
    MUTED_TEXT_COLOR,
    TEXT_COLOR,
)
from gui.main_content import MainContent  # noqa: E402
from gui.sidebar import Sidebar  # noqa: E402
from utils.logger import logger  # noqa: E402


def _launch_main_app(user):
    """Lance la fenêtre principale après authentification."""
    app = ctk.CTk()
    app.title(APP_TITLE)
    app.geometry(APP_GEOMETRY)
    app.configure(fg_color=MAIN_BG_COLOR)

    def _apply_ui_theme_inner():
        cursor_color = TEXT_COLOR
        app.option_add("*Entry.insertBackground", cursor_color)
        app.option_add("*Text.insertBackground", cursor_color)
        app.option_add("*Entry.selectBackground", BUTTON_GREEN)
        app.option_add("*Text.selectBackground", BUTTON_GREEN)

        style = ttk.Style(app)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        bg_main = MAIN_BG_COLOR
        bg_input = INPUT_BG_COLOR
        fg_main = TEXT_COLOR
        fg_muted = MUTED_TEXT_COLOR
        accent = BUTTON_GREEN
        accent_hover = BUTTON_RED
        border = "#C9DDE3"

        style.configure(
            ".", background=bg_main, foreground=fg_main, fieldbackground=bg_input
        )
        style.configure(
            "TCombobox",
            foreground=fg_main,
            fieldbackground=bg_input,
            background=bg_input,
            bordercolor=border,
            lightcolor=border,
            darkcolor=border,
            arrowcolor=fg_muted,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", bg_input)],
            foreground=[("readonly", fg_main)],
        )
        style.configure(
            "Treeview",
            background=bg_input,
            fieldbackground=bg_input,
            foreground=fg_main,
            bordercolor=border,
            rowheight=26,
        )
        style.map(
            "Treeview",
            background=[("selected", accent)],
            foreground=[("selected", "#FFFFFF")],
        )
        style.configure(
            "Treeview.Heading",
            background=bg_main,
            foreground=fg_main,
            bordercolor=border,
            relief="flat",
            font=("Poppins", 10, "bold"),
        )
        style.map(
            "Treeview.Heading",
            background=[("active", accent_hover)],
            foreground=[("active", "#FFFFFF")],
        )
        style.configure(
            "Vertical.TScrollbar",
            background=bg_input,
            troughcolor=bg_main,
            bordercolor=border,
            arrowcolor=fg_muted,
        )
        style.configure(
            "Horizontal.TScrollbar",
            background=bg_input,
            troughcolor=bg_main,
            bordercolor=border,
            arrowcolor=fg_muted,
        )

    _apply_ui_theme_inner()

    app.grid_columnconfigure(0, weight=0)
    app.grid_columnconfigure(1, weight=1)
    app.grid_rowconfigure(0, weight=1)

    main_content = MainContent(app)
    _sidebar = Sidebar(app, main_content.update_content)

    # Redirect comptable directly to financial section
    if user.get("role") == "comptable":
        main_content.update_content("financial_home")

    logger.info(
        f"Application démarrée — utilisateur : {user['username']} ({user['role']})"
    )
    app.mainloop()


def _amorcer_les_donnees():
    """Installe les classeurs manquants et signale les references vides.

    Sans ce garde-fou, un poste neuf demarre sur des classeurs absents : toutes
    les lectures renvoient une liste vide et rien ne distingue "aucun client
    enregistre" de "le fichier de donnees n'existe pas".
    """
    from utils.storage.bootstrap import describe_state, ensure_workbooks

    try:
        installes = ensure_workbooks()
    except Exception as exc:
        logger.error(f"Amorcage des classeurs impossible : {exc}", exc_info=True)
        return

    vides = []
    for nom, infos in describe_state().items():
        if infos["references_vides"]:
            vides.append(f"{nom} : {', '.join(infos['references_vides'])}")

    if installes:
        logger.info(f"Classeurs initialises : {', '.join(installes)}")
    if vides:
        logger.warning(f"Donnees de reference absentes — {' ; '.join(vides)}")

    if not installes and not vides:
        return

    lignes = []
    if installes:
        lignes.append(
            "Les classeurs de donnees ont ete crees a partir des gabarits :\n"
            + "\n".join(f"  • {os.path.basename(c)}" for c in installes)
        )
    if vides:
        lignes.append(
            "Donnees de reference absentes :\n"
            + "\n".join(f"  • {v}" for v in vides)
            + "\n\nLes cotations ne pourront pas etre chiffrees tant que ces "
            "feuilles sont vides."
        )

    try:
        messagebox.showwarning("Donnees de l'application", "\n\n".join(lignes))
    except Exception:  # pragma: no cover - pas d'affichage disponible
        print("\n\n".join(lignes))


def main():
    """Main application entry point"""
    try:
        logger.info(f"Starting {APP_TITLE}")

        # Set appearance theme
        ctk.set_appearance_mode(APPEARANCE_MODE)
        ctk.set_default_color_theme(DEFAULT_COLOR_THEME)
        logger.debug(f"Theme set: {APPEARANCE_MODE} mode, {DEFAULT_COLOR_THEME} color")

        _amorcer_les_donnees()

        # Les tables de reference se chargent pendant la saisie des
        # identifiants. Sans cela, le premier ecran de cotation ouvert payait
        # ~1,7 s pour le catalogue hotels et ~1,5 s pour l'appel reseau des
        # taux de change, soit plus de 4 secondes avant affichage.
        from utils.storage.prewarm import prechauffer

        prechauffer()

        # ── Écran de login ────────────────────────────────────────────────
        from gui.forms.login_form import LoginWindow

        logger.info("Application started successfully")
        login = LoginWindow(on_login_success=_launch_main_app)
        login.mainloop()
        # _launch_main_app lance sa propre boucle après le login

    except Exception as e:
        error_msg = f"Application error: {e}"
        logger.error(error_msg, exc_info=True)
        # Show error dialog to user
        try:
            messagebox.showerror(
                "❌ Erreur Application",
                f"Une erreur est survenue:\n\n{str(e)}\n\nVoir les logs pour plus de détails.",
            )
        except Exception as dialog_error:
            print(f"CRITICAL ERROR: {error_msg}")
            print(f"Dialog error: {dialog_error}")
        raise


if __name__ == "__main__":
    main()
