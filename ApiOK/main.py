"""
Lahimena Tours Devis Generation Application

Main entry point for the application
"""

import customtkinter as ctk
from config import *
from gui.sidebar import Sidebar
from gui.main_content import MainContent


def main():
    """Main application entry point"""
    # Set appearance theme
    ctk.set_appearance_mode(APPEARANCE_MODE)
    ctk.set_default_color_theme(DEFAULT_COLOR_THEME)

    # Initialize main application window
    app = ctk.CTk()
    app.title(APP_TITLE)
    app.geometry(APP_GEOMETRY)

    # Configure main grid layout
    app.grid_columnconfigure(0, weight=0)  # Sidebar
    app.grid_columnconfigure(1, weight=1)  # Main content area
    app.grid_rowconfigure(0, weight=1)

    # Initialize main content area
    main_content = MainContent(app)

    # Initialize sidebar with callback to main content
    sidebar = Sidebar(app, main_content.update_content)

    # Start the application loop
    app.mainloop()


if __name__ == "__main__":
    main()