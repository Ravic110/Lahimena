"""
Lahimena Tours Devis Generation Application

Main entry point for the application
"""

from tkinter import messagebox

import customtkinter as ctk

from config import (
    APP_GEOMETRY,
    APP_TITLE,
    APPEARANCE_MODE,
    DEFAULT_COLOR_THEME,
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
        logger.debug(f"Main window created: {APP_GEOMETRY}")

        # Configure main grid layout
        app.grid_columnconfigure(0, weight=0)  # Sidebar
        app.grid_columnconfigure(1, weight=1)  # Main content area
        app.grid_rowconfigure(0, weight=1)

        # Initialize main content area
        main_content = MainContent(app)
        logger.debug("Main content initialized")

        # Initialize sidebar with callback to main content
        _sidebar = Sidebar(app, main_content.update_content, main_content.refresh)
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
