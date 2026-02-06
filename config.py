"""
Configuration constants for Lahimena Tours application
"""

import os

# Application settings
APP_TITLE = "Lahimena Tours Devis Generation"
APP_GEOMETRY = "1200x500"

# Theme settings
APPEARANCE_MODE = "dark"
DEFAULT_COLOR_THEME = "blue"

# Colors
SIDEBAR_BG_COLOR = "#FFFFFF"
MAIN_BG_COLOR = "#F5F5F5"
INPUT_BG_COLOR = "#FFFFFF"
TEXT_COLOR = "#000000"
BUTTON_GREEN = "#059669"
BUTTON_GREEN_HOVER = "#3DC096"
BUTTON_BLUE = "#3B82F6"
BUTTON_RED = "#D31F25"

# Fonts
TITLE_FONT = ("Arial", 16, "bold")
LABEL_FONT = ("Arial", 10, "bold")
ENTRY_FONT = ("Arial", 11)
BUTTON_FONT = ("Arial", 11, "bold")

# File paths (using absolute paths)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "assets", "logo.png")
CLIENT_EXCEL_PATH = os.path.join(BASE_DIR, "data.xlsx")
HOTEL_EXCEL_PATH = os.path.join(BASE_DIR, "data-hotel.xlsx")
DEVIS_FOLDER = os.path.join(BASE_DIR, "devis")
CLIENT_SHEET_NAME = "DEMANDE_CLIENT"
HOTEL_SHEET_NAME = "BDD_HOTEL"
COTATION_H_SHEET_NAME = "COTATION_H"

# Form constants
PERIODES = ["Haute saison", "Moyenne saison", "Basse saison"]
RESTAURATIONS = ["Sans restauration", "Petit déjeuner", "Demi-pension",
                "Pension complète", "All inclusive soft", "All inclusive",
                "Ultra all inclusive"]
TYPE_HEBERGEMENTS = ["Bivouac", "Hôtel 3*", "Hôtel 2*", "Hôtel 4*",
                    "Hôtel 5*", "Maison d'hôte", "Auberge de vacances",
                    "Bâteau", "Gîte"]
TYPE_CHAMBRES = ["Single", "Double/twin", "Triple", "Familliale"]
AGES_ENFANTS = ["0 à 2 ans", "2 à 6 ans", "6 à 12 ans"]
FORFAITS = ["À la carte", "Sur mesure"]
CIRCUITS = ["Circuit nature", "Circuit aventure", "Circuit tourisme solidaire",
           "Circuit culturel", "Circuits trek et randonnée", "Circuits détente",
           "Voyage de noce"]

# Phone codes
PHONE_CODES = ["+261", "+33", "+32", "+34"]
DEFAULT_PHONE_CODE = "+261"