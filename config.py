"""
Configuration constants for Lahimena Tours application
"""

import json
import os
import sys

# load optional configuration file to override defaults
_config_path = os.path.join(os.path.dirname(__file__), "config.json")

# internal config dictionary
_cfg = {}


def load_config(path=None):
    """Load configuration from JSON file and update _cfg."""
    global _cfg
    if path is None:
        path = _config_path
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                _cfg = json.load(f)
        except Exception:
            _cfg = {}
    else:
        _cfg = {}
    # propagate overrides to existing globals if they already exist
    for key, val in _cfg.items():
        globals()[key] = val


# initialize at import
load_config()

# Application settings
APP_TITLE = "Lahimena Tours Devis Generation"
APP_GEOMETRY = "1200x500"

# override from JSON config if provided
if "APP_TITLE" in _cfg:
    APP_TITLE = _cfg["APP_TITLE"]
if "APP_GEOMETRY" in _cfg:
    APP_GEOMETRY = _cfg["APP_GEOMETRY"]

# Theme settings
APPEARANCE_MODE = _cfg.get("APPEARANCE_MODE", "dark")
DEFAULT_COLOR_THEME = _cfg.get("DEFAULT_COLOR_THEME", "blue")
CURRENT_THEME = _cfg.get("CURRENT_THEME", "dark")

THEMES = {
    "dark": {
        "SIDEBAR_BG_COLOR": "#071026",
        "MAIN_BG_COLOR": "#0B1226",
        "PANEL_BG_COLOR": "#0F1B2B",
        "CARD_BG_COLOR": "#0F1724",
        "CARD_HOVER_BG_COLOR": "#172033",
        "INPUT_BG_COLOR": "#0B1220",
        "READONLY_BG_COLOR": "#08101A",
        "TEXT_COLOR": "#FFFFFF",
        "MUTED_TEXT_COLOR": "#9AA8B8",
        "ACCENT_TEXT_COLOR": "#A0D2FF",
        "BUTTON_GREEN": "#059669",
        "BUTTON_GREEN_HOVER": "#047857",
        "BUTTON_BLUE": "#0284C7",
        "BUTTON_RED": "#DC2626",
        "BUTTON_ORANGE": "#F59E0B",
        "BUTTON_GRAY": "#64748B",
    },
    "light": {
        "SIDEBAR_BG_COLOR": "#FFFFFF",
        "MAIN_BG_COLOR": "#F5F5F5",
        "PANEL_BG_COLOR": "#FFFFFF",
        "CARD_BG_COLOR": "#FFFFFF",
        "CARD_HOVER_BG_COLOR": "#E5E7EB",
        "INPUT_BG_COLOR": "#FFFFFF",
        "READONLY_BG_COLOR": "#E8F5E9",
        "TEXT_COLOR": "#111827",
        "MUTED_TEXT_COLOR": "#6B7280",
        "ACCENT_TEXT_COLOR": "#2563EB",
        "BUTTON_GREEN": "#16A34A",
        "BUTTON_GREEN_HOVER": "#15803D",
        "BUTTON_BLUE": "#2563EB",
        "BUTTON_RED": "#DC2626",
        "BUTTON_ORANGE": "#F59E0B",
        "BUTTON_GRAY": "#64748B",
    },
}

# Colors (active theme)
SIDEBAR_BG_COLOR = THEMES[CURRENT_THEME]["SIDEBAR_BG_COLOR"]
MAIN_BG_COLOR = THEMES[CURRENT_THEME]["MAIN_BG_COLOR"]
PANEL_BG_COLOR = THEMES[CURRENT_THEME]["PANEL_BG_COLOR"]
CARD_BG_COLOR = THEMES[CURRENT_THEME]["CARD_BG_COLOR"]
CARD_HOVER_BG_COLOR = THEMES[CURRENT_THEME]["CARD_HOVER_BG_COLOR"]
INPUT_BG_COLOR = THEMES[CURRENT_THEME]["INPUT_BG_COLOR"]
READONLY_BG_COLOR = THEMES[CURRENT_THEME]["READONLY_BG_COLOR"]
TEXT_COLOR = THEMES[CURRENT_THEME]["TEXT_COLOR"]
MUTED_TEXT_COLOR = THEMES[CURRENT_THEME]["MUTED_TEXT_COLOR"]
ACCENT_TEXT_COLOR = THEMES[CURRENT_THEME]["ACCENT_TEXT_COLOR"]
BUTTON_GREEN = THEMES[CURRENT_THEME]["BUTTON_GREEN"]
BUTTON_GREEN_HOVER = THEMES[CURRENT_THEME]["BUTTON_GREEN_HOVER"]
BUTTON_BLUE = THEMES[CURRENT_THEME]["BUTTON_BLUE"]
BUTTON_RED = THEMES[CURRENT_THEME]["BUTTON_RED"]
BUTTON_ORANGE = THEMES[CURRENT_THEME]["BUTTON_ORANGE"]
BUTTON_GRAY = THEMES[CURRENT_THEME]["BUTTON_GRAY"]

THEME_KEYS = (
    "CURRENT_THEME",
    "APPEARANCE_MODE",
    "SIDEBAR_BG_COLOR",
    "MAIN_BG_COLOR",
    "PANEL_BG_COLOR",
    "CARD_BG_COLOR",
    "CARD_HOVER_BG_COLOR",
    "INPUT_BG_COLOR",
    "READONLY_BG_COLOR",
    "TEXT_COLOR",
    "MUTED_TEXT_COLOR",
    "ACCENT_TEXT_COLOR",
    "BUTTON_GREEN",
    "BUTTON_GREEN_HOVER",
    "BUTTON_BLUE",
    "BUTTON_RED",
    "BUTTON_ORANGE",
    "BUTTON_GRAY",
)


def _propagate_theme():
    """Update modules that imported config symbols via 'from config import *'."""
    for module in list(sys.modules.values()):
        if module is None:
            continue
        for key in THEME_KEYS:
            if hasattr(module, key):
                setattr(module, key, globals()[key])


def apply_theme(theme_name):
    """Apply a named theme and update exported color constants."""
    global CURRENT_THEME, APPEARANCE_MODE
    global SIDEBAR_BG_COLOR, MAIN_BG_COLOR, PANEL_BG_COLOR
    global CARD_BG_COLOR, CARD_HOVER_BG_COLOR
    global INPUT_BG_COLOR, READONLY_BG_COLOR
    global TEXT_COLOR, MUTED_TEXT_COLOR, ACCENT_TEXT_COLOR
    global BUTTON_GREEN, BUTTON_GREEN_HOVER, BUTTON_BLUE
    global BUTTON_RED, BUTTON_ORANGE, BUTTON_GRAY

    if theme_name not in THEMES:
        return

    CURRENT_THEME = theme_name
    APPEARANCE_MODE = "dark" if theme_name == "dark" else "light"
    theme = THEMES[theme_name]

    SIDEBAR_BG_COLOR = theme["SIDEBAR_BG_COLOR"]
    MAIN_BG_COLOR = theme["MAIN_BG_COLOR"]
    PANEL_BG_COLOR = theme["PANEL_BG_COLOR"]
    CARD_BG_COLOR = theme["CARD_BG_COLOR"]
    CARD_HOVER_BG_COLOR = theme["CARD_HOVER_BG_COLOR"]
    INPUT_BG_COLOR = theme["INPUT_BG_COLOR"]
    READONLY_BG_COLOR = theme["READONLY_BG_COLOR"]
    TEXT_COLOR = theme["TEXT_COLOR"]
    MUTED_TEXT_COLOR = theme["MUTED_TEXT_COLOR"]
    ACCENT_TEXT_COLOR = theme["ACCENT_TEXT_COLOR"]
    BUTTON_GREEN = theme["BUTTON_GREEN"]
    BUTTON_GREEN_HOVER = theme["BUTTON_GREEN_HOVER"]
    BUTTON_BLUE = theme["BUTTON_BLUE"]
    BUTTON_RED = theme["BUTTON_RED"]
    BUTTON_ORANGE = theme["BUTTON_ORANGE"]
    BUTTON_GRAY = theme["BUTTON_GRAY"]
    _propagate_theme()


# Fonts
TITLE_FONT = ("Arial", 16, "bold")
LABEL_FONT = ("Arial", 10, "bold")
ENTRY_FONT = ("Arial", 11)
BUTTON_FONT = ("Arial", 11, "bold")

# File paths (using absolute paths)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "assets", "logo.png")
CLIENT_EXCEL_PATH = os.path.join(BASE_DIR, _cfg.get("CLIENT_EXCEL_PATH", "data.xlsx"))
HOTEL_EXCEL_PATH = os.path.join(
    BASE_DIR, _cfg.get("HOTEL_EXCEL_PATH", "data-hotel.xlsx")
)
FINANCIAL_EXCEL_PATH = CLIENT_EXCEL_PATH
DEVIS_FOLDER = os.path.join(BASE_DIR, "devis")
CLIENT_SHEET_NAME = "DEMANDE_CLIENT"
CLIENT_INFOS_SHEET_NAME = "INFOS_CLIENTS"
HOTEL_SHEET_NAME = "BDD_HOTEL"
COTATION_H_SHEET_NAME = "COTATION_H"
COTATION_FRAIS_COL_SHEET_NAME = "COTATION_FRAIS_COL"
FRAIS_COLLECTIFS_SHEET_NAME = "Frais collectifs"

# Form constants
PERIODES = ["Haute saison", "Moyenne saison", "Basse saison"]
RESTAURATIONS = [
    "Sans restauration",
    "Petit déjeuner",
    "Demi-pension",
    "Pension complète",
    "All inclusive soft",
    "All inclusive",
    "Ultra all inclusive",
]
TYPE_HEBERGEMENTS = [
    "Bivouac",
    "Hôtel 3*",
    "Hôtel 2*",
    "Hôtel 4*",
    "Hôtel 5*",
    "Maison d'hôte",
    "Auberge de vacances",
    "Bâteau",
    "Gîte",
]
TYPE_CHAMBRES = ["Single", "Double/twin", "Triple", "Familliale"]
AGES_ENFANTS = ["0 à 2 ans", "2 à 6 ans", "6 à 12 ans"]
FORFAITS = ["À la carte", "Sur mesure"]
CIRCUITS = [
    "Circuits nature",
    "Circuits aventure",
    "Circuit Tourisme Solidaire",
    "Circuit Culturel",
    "Circuits trek et randonnées",
    "Circuits détente",
    "Voyages de noce",
]
HOTEL_ARRIVAL_TYPES = ["1*", "2*", "3*", "4*", "5*", "Luxe", "Eco"]

# Phone codes
PHONE_CODES = ["+261", "+33", "+32", "+34"]
DEFAULT_PHONE_CODE = "+261"
