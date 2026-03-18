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
_ALLOWED_CONFIG_KEYS = {
    "APP_TITLE",
    "APP_GEOMETRY",
    "APPEARANCE_MODE",
    "DEFAULT_COLOR_THEME",
    "CURRENT_THEME",
    "CLIENT_EXCEL_PATH",
    "HOTEL_EXCEL_PATH",
    "COMPANY_NAME",
    "COMPANY_TAGLINE",
    "COMPANY_PHONE",
    "PDF_FOOTER_TEXT",
}


def _sanitize_config(raw_cfg):
    """Return only allowed config keys from a user-provided dict."""
    if not isinstance(raw_cfg, dict):
        return {}
    return {k: v for k, v in raw_cfg.items() if k in _ALLOWED_CONFIG_KEYS}


def _propagate_config_overrides():
    """Propagate sanitized config values to already-defined globals."""
    for key, val in _cfg.items():
        if key in globals():
            globals()[key] = val

    # Recompute derived absolute paths when base constants are already loaded.
    if "BASE_DIR" in globals():
        globals()["CLIENT_EXCEL_PATH"] = os.path.join(
            BASE_DIR, _cfg.get("CLIENT_EXCEL_PATH", "data.xlsx")
        )
        globals()["HOTEL_EXCEL_PATH"] = os.path.join(
            BASE_DIR, _cfg.get("HOTEL_EXCEL_PATH", "data-hotel.xlsx")
        )
        globals()["FINANCIAL_EXCEL_PATH"] = globals()["CLIENT_EXCEL_PATH"]


def load_config(path=None):
    """Load configuration from JSON file and update _cfg."""
    global _cfg
    if path is None:
        path = _config_path
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                _cfg = _sanitize_config(json.load(f))
        except Exception:
            _cfg = {}
    else:
        _cfg = {}
    _propagate_config_overrides()


# initialize at import
load_config()

# Application settings
APP_TITLE = "Lahimena Tours Devis Generation"
APP_GEOMETRY = "1250x720"
COMPANY_NAME = _cfg.get("COMPANY_NAME", "Lahimena Tours")
COMPANY_TAGLINE = _cfg.get("COMPANY_TAGLINE", "Madagascar - Tours & Travel")
COMPANY_PHONE = _cfg.get("COMPANY_PHONE", "+261-34-00-000-00")
PDF_FOOTER_TEXT = _cfg.get(
    "PDF_FOOTER_TEXT", f"{COMPANY_NAME} | Madagascar | Tel: {COMPANY_PHONE}"
)

# override from JSON config if provided
if "APP_TITLE" in _cfg:
    APP_TITLE = _cfg["APP_TITLE"]
if "APP_GEOMETRY" in _cfg:
    APP_GEOMETRY = _cfg["APP_GEOMETRY"]

# Theme settings (fixed light theme for the new UI design)
APPEARANCE_MODE = "light"
DEFAULT_COLOR_THEME = _cfg.get("DEFAULT_COLOR_THEME", "blue")
CURRENT_THEME = "light"

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
        "MAIN_BG_COLOR": "#EAF4F7",
        "PANEL_BG_COLOR": "#DCEFF3",
        "CARD_BG_COLOR": "#F5FBFD",
        "CARD_HOVER_BG_COLOR": "#EFF7FA",
        "INPUT_BG_COLOR": "#FFFFFF",
        "READONLY_BG_COLOR": "#E3F0F4",
        "TEXT_COLOR": "#2D4E57",
        "MUTED_TEXT_COLOR": "#6C8A92",
        "ACCENT_TEXT_COLOR": "#C62828",
        "BUTTON_GREEN": "#0E7C86",
        "BUTTON_GREEN_HOVER": "#0B6A72",
        "BUTTON_BLUE": "#0F7D8A",
        "BUTTON_RED": "#C62828",
        "BUTTON_ORANGE": "#F4A623",
        "BUTTON_GRAY": "#8AA1A8",
    },
}

if CURRENT_THEME not in THEMES:
    CURRENT_THEME = "dark"

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
TITLE_FONT = ("Poppins", 16, "bold")
LABEL_FONT = ("Poppins", 10, "bold")
ENTRY_FONT = ("Poppins", 11)
BUTTON_FONT = ("Poppins", 11, "bold")

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
VISITE_EXCURSION_SOURCE_SHEET_NAME = "Visite_excursion"
VISITE_EXCURSION_SHEET_NAME = "VISITE_EXCURSION"
AVION_SOURCE_SHEET_NAME = "avion"
AVION_SHEET_NAME = "AVION"
TRANSPORT_SOURCE_SHEET_NAME = "TRANSPORT"
TRANSPORT_SHEET_NAME = "TRANSPORT"
KM_MADA_SHEET_NAME = "KM_MADA"
PARAMETRAGE_SHEET_NAME = "PARAMETRE"
INVOICE_SHEET_NAME = "FACTURES_CLIENTS"
FINANCIAL_STATE_SHEET_NAME = "ETAT_FINANCIER_AUTO"

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

# Country → dial code mapping (used in mobile field)
COUNTRY_PHONE_MAP = {
    "Madagascar": "+261",
    "France": "+33",
    "Réunion": "+262",
    "Belgique": "+32",
    "Suisse": "+41",
    "Luxembourg": "+352",
    "Italie": "+39",
    "Espagne": "+34",
    "Allemagne": "+49",
    "Autriche": "+43",
    "Pays-Bas": "+31",
    "Portugal": "+351",
    "Royaume-Uni": "+44",
    "Irlande": "+353",
    "Canada": "+1",
    "États-Unis": "+1",
    "Australie": "+61",
    "Nouvelle-Zélande": "+64",
    "Afrique du Sud": "+27",
    "Maurice": "+230",
    "Comores": "+269",
    "Seychelles": "+248",
    "Maroc": "+212",
    "Algérie": "+213",
    "Tunisie": "+216",
    "Sénégal": "+221",
    "Côte d'Ivoire": "+225",
    "Cameroun": "+237",
    "Kenya": "+254",
    "Tanzanie": "+255",
    "Mozambique": "+258",
    "Japon": "+81",
    "Chine": "+86",
    "Inde": "+91",
    "Brésil": "+55",
}
DEFAULT_COUNTRY = "Madagascar"
# Reverse lookup: code → country name (first match)
CODE_TO_COUNTRY = {}
for _c, _code in COUNTRY_PHONE_MAP.items():
    if _code not in CODE_TO_COUNTRY:
        CODE_TO_COUNTRY[_code] = _c
