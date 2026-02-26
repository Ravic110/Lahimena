"""
SCHEMA DE FACTORISATION
Conversion de apifactorise.py en architecture modulaire

AVANT : Un seul fichier monolithique
apifactorise.py (400+ lignes)
├── CONFIG (dict)
├── DataManager (classe)
├── PCGManager (classe)
├── DialogueLigne (classe)
└── ComptabiliteApp (classe)

APRES : Architecture bien organisée
apifactorise_refactored/
├── config.py (33 lignes)               ◄── Configuration extraite
├── main.py (30 lignes)                 ◄── Point d'entrée simple
├── models/
│   ├── __init__.py
│   └── data.py (100 lignes)           ◄── Gestionnaires de données
├── ui/
│   ├── __init__.py
│   ├── main.py (230 lignes)           ◄── Interface principale
│   └── dialogs.py (170 lignes)        ◄── Dialogues modaux
└── utils/
    ├── __init__.py
    └── formatters.py (45 lignes)      ◄── Utilitaires réutilisables
"""

# ========================================
# CORRESPONDANCES CLASSES/FONCTIONS
# ========================================

MAPPING = {
    "ORIGINAL (apifactorise.py)": {
        "CONFIG": "→ config.py",
        "DataManager": "→ models/data.py",
        "PCGManager": "→ models/data.py",
        "DialogueLigne": "→ ui/dialogs.py",
        "ComptabiliteApp": "→ ui/main.py",
    }
}

# ========================================
# BENEFICES DE LA REFACTORISATION
# ========================================

BENEFICES = {
    "1. Maintenabilité": [
        "- Chaque module a une responsabilité unique (SRP)",
        "- Code plus facile à comprendre et modifier",
        "- Bugfixes localisés à un seul endroit",
    ],
    "2. Réutilisabilité": [
        "- DataManager et PCGManager peuvent être importés elsewhere",
        "- Formatters peuvent être utilisés dans d'autres modules",
        "- Dialogs peuvent être réutilisées ou modifiées",
    ],
    "3. Testabilité": [
        "- Chaque classe peut être testée indépendamment",
        "- Mocking et fixtures plus simples",
        "- Tests d'intégration plus faciles",
    ],
    "4. Extensibilité": [
        "- Ajouter de nouveaux états financiers : facile",
        "- Ajouter de nouveaux dialogues : dans ui/",
        "- Ajouter de nouvelles sources de données : dans models/",
    ],
    "5. Scalabilité": [
        "- Code prêt pour une équipe de développeurs",
        "- Fusion/conflits git minimisés",
        "- Structure compatible avec des frameworks modernes",
    ],
}

# ========================================
# GUIDE D'UTILISATION
# ========================================

USAGE_EXAMPLES = """
# 1. DEMARRER L'APPLICATION
python main.py

# 2. UTILISER DIRECTEMENT EN IMPORT
from models.data import DataManager
from config import CONFIG

df = DataManager.charger_feuille("EtatFidata.xlsx", "Journal")

# 3. UTILISER LES FORMATTERS
from utils.formatters import format_montant, parse_montant

montant = parse_montant("1234,56")  # 1234.56
affiche = format_montant(montant)   # "1 234.56"

# 4. CREER UN DIALOGUE PERSONNALISE
from .ui.dialogs import DialogueLigne
dialogue = DialogueLigne(parent, df, pcg_comptes, pcg_numeros, ajout=True)

# 5. ACCEDER A LA CONFIG
from config import CONFIG
fichier = CONFIG['fichier_defaut']
colonnes = CONFIG['colonnes_journal']
"""

# ========================================
# CHECKLIST DE MIGRATION
# ========================================

MIGRATION_CHECKLIST = [
    "✅ Configuration centralisée → config.py",
    "✅ Classes DataManager/PCGManager → models/data.py",
    "✅ Classe DialogueLigne → ui/dialogs.py",
    "✅ Classe ComptabiliteApp → ui/main.py",
    "✅ Fonctions formatage → utils/formatters.py",
    "✅ Point d'entrée simplifié → main.py",
    "✅ Documentation → README.md",
    "✅ Imports __init__.py configurés",
    "✅ Chemins sys.path pour imports relatifs",
    "✅ Tests de compatibilité réussis",
]

if __name__ == "__main__":
    print("SCHEMA DE FACTORISATION - apifactorise.py")
    print("=" * 50)
    for section, content in [
        ("STRUCTURE", MAPPING),
        ("BENEFICES", BENEFICES),
        ("CHECKLIST", MIGRATION_CHECKLIST),
    ]:
        print(f"\n📋 {section}")
        print("-" * 50)
        if isinstance(content, dict):
            for k, v in content.items():
                if isinstance(v, dict):
                    print(f"  {k}:")
                    for mk, mv in v.items():
                        print(f"    {mk} {mv}")
                elif isinstance(v, list):
                    print(f"  {k}:")
                    for item in v:
                        print(f"    {item}")
        elif isinstance(content, list):
            for item in content:
                print(f"  {item}")

    print("\n\n📚 EXEMPLES D'UTILISATION")
    print("-" * 50)
    print(USAGE_EXAMPLES)
