"""
Lahimena Tours Devis Generation Application

Main entry point for the application
"""

from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
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
from gui.main_content import MainContent
from gui.sidebar import Sidebar
from utils.logger import logger


def main():
    """Main application entry point"""
    try:
        logger.info(f"Starting {APP_TITLE}")

        # Set appearance theme
        ctk.set_appearance_mode(APPEARANCE_MODE)
        ctk.set_default_color_theme(DEFAULT_COLOR_THEME)
        logger.debug(f"Theme set: {APPEARANCE_MODE} mode, {DEFAULT_COLOR_THEME} color")

        # Initialize main application window
        app = ctk.CTk()
        app.title(APP_TITLE)
        app.geometry(APP_GEOMETRY)
        app.configure(fg_color=MAIN_BG_COLOR)
        logger.debug(f"Main window created: {APP_GEOMETRY}")

        def _apply_ui_theme():
            cursor_color = TEXT_COLOR
            app.option_add("*Entry.insertBackground", cursor_color)
            app.option_add("*Text.insertBackground", cursor_color)
            app.option_add("*Entry.selectBackground", BUTTON_GREEN)
            app.option_add("*Text.selectBackground", BUTTON_GREEN)

            # Harmonize ttk widgets (Combobox, Treeview, headings, scrollbar).
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
                ".",
                background=bg_main,
                foreground=fg_main,
                fieldbackground=bg_input,
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

        # Keep insertion cursor visible for both dark and light themes.
        _apply_ui_theme()

        # Configure main grid layout
        app.grid_columnconfigure(0, weight=0)  # Sidebar
        app.grid_columnconfigure(1, weight=1)  # Main content area
        app.grid_rowconfigure(0, weight=1)

        # Initialize main content area
        main_content = MainContent(app)
        logger.debug("Main content initialized")

        # Initialize sidebar with callback to main content
        _sidebar = Sidebar(app, main_content.update_content)
        logger.debug("Sidebar initialized")

        logger.info("Application started successfully")

        # Start the application loop
        app.mainloop()

        logger.info("Application closed")

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
