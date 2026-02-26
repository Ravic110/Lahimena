import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import pandas as pd

from finances.tsarakonta.config import CONFIG
from finances.tsarakonta.models.data import DataManager, PCGManager

from .ui.etat_resultat import build_result_df, export_pdf_from_df

BASE = os.path.dirname(os.path.dirname(__file__))
STATE_FOLDER = os.path.join(BASE, "EtatFiFolder")
os.makedirs(STATE_FOLDER, exist_ok=True)
MAPPING_CSV = os.path.join(STATE_FOLDER, "mapping_resultat_nature.csv")

# Charger journal
fichier = CONFIG.get("fichier_defaut")
journal = DataManager.charger_feuille(fichier, CONFIG.get("feuille_journal"))
if journal is None:
    print("Journal non trouvé, création d'exemple temporaire")
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

# Build result
result_df = build_result_df(pcg_dict, journal, MAPPING_CSV, selection="Both")

# Save CompteResultatNature.csv
csv1 = os.path.join(STATE_FOLDER, "CompteResultatNature.csv")
result_df.to_csv(csv1, index=False, encoding="utf-8")
print("Wrote", csv1)

# Also save ompteResultatNature.csv (as requested)
csv2 = os.path.join(STATE_FOLDER, "ompteResultatNature.csv")
result_df.to_csv(csv2, index=False, encoding="utf-8")
print("Wrote", csv2)

# Save Excel and PDF
excel_path = os.path.join(STATE_FOLDER, "CompteResultatNature.xlsx")
result_df.to_excel(excel_path, index=False)
print("Excel:", excel_path)

pdf_path = os.path.join(STATE_FOLDER, "CompteResultatNature.pdf")
export_pdf_from_df(
    result_df,
    pdf_path,
    title="Compte de résultat par nature",
    header_prefs={},
    arrete="2025",
)
print("PDF:", pdf_path)
