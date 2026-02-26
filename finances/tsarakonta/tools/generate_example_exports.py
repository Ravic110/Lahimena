"""
Génère un export d'exemple (Excel + PDF) pour le Compte de résultat par nature
Utilise : CONFIG['fichier_defaut'] et 'Journal' + 'pcg.xlsx' pour PCG
Le PDF et Excel seront créés dans EtatFiFolder
"""

import os
import sys

# ensure project root is importable when running this script
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import pandas as pd

from finances.tsarakonta.config import CONFIG
from finances.tsarakonta.models.data import DataManager, PCGManager

from .ui.etat_resultat import build_result_df, export_pdf_from_df

BASE = os.path.dirname(os.path.dirname(__file__))
STATE_FOLDER = os.path.join(BASE, "EtatFiFolder")
os.makedirs(STATE_FOLDER, exist_ok=True)
MAPPING_CSV = os.path.join(STATE_FOLDER, "mapping_resultat_nature.csv")
SETTINGS_JSON = os.path.join(STATE_FOLDER, "settings.json")

# Charger journal
fichier = CONFIG.get("fichier_defaut")
journal = DataManager.charger_feuille(fichier, CONFIG.get("feuille_journal"))
if journal is None:
    # créer un dataframe d'exemple minimal
    journal = pd.DataFrame(columns=CONFIG["colonnes_journal"])
    journal.loc[0] = [
        "01/01/2025",
        "Vente",
        "01/01/2025",
        0.0,
        100000000.0,
        "701",
        "",
        "2025",
    ]
    journal.loc[1] = [
        "01/01/2025",
        "Achat",
        "01/01/2025",
        50000000.0,
        0.0,
        "601",
        "",
        "2025",
    ]
    journal.loc[2] = [
        "01/01/2024",
        "Vente",
        "01/01/2024",
        0.0,
        80000000.0,
        "701",
        "",
        "2024",
    ]
    journal.loc[3] = [
        "01/01/2024",
        "Achat",
        "01/01/2024",
        25000000.0,
        0.0,
        "601",
        "",
        "2024",
    ]

# Charger PCG
pcg_comptes, pcg_numeros, pcg_dict = PCGManager.charger_pcg(fichier)

# Assurer mapping
if not os.path.exists(MAPPING_CSV):
    # create will be done by the module when invoked from UI; create a small default here
    default = [
        (1, "Période d'exercice", "", "title"),
        (2, "Chiffres d'affaires", "70", "produit"),
        (3, "Achat consommés", "60", "charge"),
        (4, "Charges de personnel", "64", "charge"),
        (5, "TOTAL PRODUITS DES ACTIVITES ORDINAIRES", "70", "subtotal"),
        (6, "TOTAL CHARGES DES ACTIVITES ORDINAIRES", "60;64", "subtotal"),
        (7, "RESULTAT NET DE L'EXERCICE", "", "subtotal"),
    ]
    import pandas as pd

    pd.DataFrame(default, columns=["ordre", "rubrique", "prefixes", "type"]).to_csv(
        MAPPING_CSV, index=False, encoding="utf-8"
    )

# Build result df
result_df = build_result_df(pcg_dict, journal, MAPPING_CSV, selection="Both")

# Save Excel
excel_path = os.path.join(
    STATE_FOLDER,
    f"example_compte_resultat_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
)
result_df.to_excel(excel_path, index=False)
print("Excel exporté:", excel_path)

# Load settings if any
prefs = {}
try:
    import json

    if os.path.exists(SETTINGS_JSON):
        with open(SETTINGS_JSON, "r", encoding="utf-8") as f:
            prefs = json.load(f)
except Exception:
    prefs = {}

# Save PDF
pdf_path = os.path.join(
    STATE_FOLDER,
    f"example_compte_resultat_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.pdf",
)
export_pdf_from_df(
    result_df,
    pdf_path,
    title=prefs.get("nom_societe", "Compte de résultat par nature"),
    header_prefs=prefs,
    arrete="2025",
)
print("PDF exporté:", pdf_path)
