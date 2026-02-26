"""
Configuration centralisée de l'application comptable
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)

DEFAULT_DATA_PATH = os.environ.get("TSARAKONTA_DATA_PATH")
DEFAULT_PCG_PATH = os.environ.get("TSARAKONTA_PCG_PATH")

if not DEFAULT_DATA_PATH:
    DEFAULT_DATA_PATH = os.path.join(PROJECT_ROOT, "data.xlsx")
if not DEFAULT_PCG_PATH:
    DEFAULT_PCG_PATH = os.path.join(BASE_DIR, "pcg.xlsx")

CONFIG = {
    "fichier_defaut": DEFAULT_DATA_PATH,
    "fichier_pcg": DEFAULT_PCG_PATH,
    "feuille_journal": "Journal",
    "feuille_pcg": "pcg",
    "colonnes_journal": [
        "Date",
        "Libellé",
        "DateValeur",
        "MontantDébit",
        "MontantCrédit",
        "CompteDébit",
        "CompteCrédit",
        "Année",
    ],
    "largeurs_colonnes": {
        "Libellé": 200,
        "DateValeur": 100,
        "Date": 100,
        "MontantDébit": 120,
        "MontantCrédit": 120,
        "CompteDébit": 100,
        "CompteCrédit": 100,
    },
    "etats_financiers": [
        "Bilan actif",
        "Bilan passif",
        "Compte de résultat par nature",
        "Compte de résultat par fonction",
        "Tableau de flux de trésorerie (méthode directe)",
        "Tableau de flux de trésorerie (méthode indirecte)",
        "Etat de variation des capitaux propres",
    ],
}
