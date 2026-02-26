"""
STRUCTURE VISUELLE - apifactorise_refactored

Cet archivage montre la hiérarchie complète et les dépendances
"""

import os


def afficher_arborescence(repertoire, prefix="", is_last=True):
    """Affiche l'arborescence du répertoire"""
    items = []
    ignored = {"__pycache__", ".git", ".pytest_cache", ".venv", "venv"}

    try:
        items = [x for x in sorted(os.listdir(repertoire)) if x not in ignored]
    except PermissionError:
        return

    for i, item in enumerate(items):
        path = os.path.join(repertoire, item)
        is_last_item = i == len(items) - 1

        current_prefix = "└── " if is_last_item else "├── "
        print(f"{prefix}{current_prefix}{item}")

        if os.path.isdir(path) and not item.startswith("."):
            extension = "    " if is_last_item else "│   "
            afficher_arborescence(path, prefix + extension, is_last_item)


def main():
    """Affiche la structure complète"""

    print("""
╔════════════════════════════════════════════════════════════════════════╗
║           STRUCTURE VISUELLE DE apifactorise_refactored               ║
╚════════════════════════════════════════════════════════════════════════╝

📁 Répertoire: apifactorise_refactored/
   │
""")

    # Récupérer le chemin
    script_dir = os.path.dirname(os.path.abspath(__file__))

    print("   Contenu:\n")

    # Afficher l'arborescence
    afficher_arborescence(script_dir, "   ")

    print("""

═══════════════════════════════════════════════════════════════════════════

📊 ARCHITECTURE LOGIQUE
═══════════════════════════════════════════════════════════════════════════

                           ┌──────────────┐
                           │  UTILISATEUR │
                           └──────┬───────┘
                                  │
                         ┌────────▼────────┐
                         │   Tkinter GUI   │
                         │    (main.py)    │
                         └────────┬────────┘
                                  │
              ┌───────────────────┼──────────────────┐
              │                   │                  │
        ┌─────▼──────┐    ┌──────▼──────┐    ┌─────▼──────┐
        │   UI/Main   │    │ UI/Dialogs  │    │ Menu/States│
        │ComptabiliteA├───┤DialogueLigne├───┤(Fichier,   │
        │   pp        │    │             │    │Etat Fin)   │
        └─────┬──────┘    └──────┬──────┘    └─────┬──────┘
              │                   │                 │
              └───────────────────┼─────────────────┘
                                  │
                         ┌────────▼────────┐
                         │ MODELS/DATA.PY  │
                         │ DataManager     │
                         │ PCGManager      │
                         └────────┬────────┘
                                  │
                    ┌─────────────┼─────────────┐
                    │                           │
            ┌──────▼──────────┐      ┌─────────▼────┐
            │ FICHIERS EXCEL  │      │ CONFIGURATION│
            │ EtatFidata.xlsx │      │  config.py   │
            └─────────────────┘      └──────────────┘

                   ┌────┐
                   │ UI │  ← Dépend de:
                   └─┬──┘
      config.py ◄───┤
      models.data◄──┤
      utils.formatters◄────┐
                            │
                   ┌────────▼──────┐
                   │UTILS/FORMATTERS│
                   │ formatage      │
                   │ parsing        │
                   └────────────────┘


═══════════════════════════════════════════════════════════════════════════

🔄 FLUX DE DONNÉES
═══════════════════════════════════════════════════════════════════════════

1. CHARGEMENT DES DONNÉES:
   
   Utilisateur lance main.py
   ↓
   ComptabiliteApp.__init__()
   ↓
   charger_pcg() ──→ PCGManager.charger_pcg()
   ↓
   charger_donnees() ──→ DataManager.charger_feuille()
   ↓
   afficher_donnees() ──→ Boucle sur DataFrame
   ↓
   Affichage dans Treeview


2. AJOUT D'UNE ÉCRITURE:
   
   Utilisateur clique "Ajouter"
   ↓
   ajouter_ligne() → ouvrir_dialogue(ajout=True)
   ↓
   DialogueLigne.creer_widgets()
   ↓
   Utilisateur complète les champs
   ↓
   Valider → preparer_ligne()
   ↓
   parse_montant() (utils)
   ↓
   Ajouter au DataFrame
   ↓
   DataManager.sauvegarder_df()
   ↓
   Retour à l'affichage des données


3. MODIFICATION:
   
   Utilisateur sélectionne + clique "Modifier"
   ↓
   modifier_ligne() → ouvrir_dialogue(ajout=False, index=..., donnees=...)
   ↓
   DialogueLigne avec pré-remplissage
   ↓
   format_montant() pour affichage
   ↓
   Modification et validertion
   ↓
   parse_montant()
   ↓
   Mise à jour ligne dans DataFrame
   ↓
   DataManager.sauvegarder_df()


═══════════════════════════════════════════════════════════════════════════

📦 DÉPENDANCES ENTRE MODULES
═══════════════════════════════════════════════════════════════════════════

config.py
├─ Pas de dépendances externes internes
└─ Utilisé par: models, ui, utils

models/__init__.py
└─ Importe: data.py

models/data.py
├─ Dépend de: config, messagebox (tkinter)
├─ Utilisé par: ui.main

utils/__init__.py
└─ Importe: formatters.py

utils/formatters.py
├─ Pas de dépendances
├─ Utilisé par: ui.dialogs, ui.main

ui/__init__.py
├─ Importe: main.py, dialogs.py
└─ Export: ComptabiliteApp, DialogueLigne

ui/main.py (ComptabiliteApp)
├─ Importe: config, models.data, utils.formatters, ui.dialogs
├─ Utilisé par: main.py (point d'entrée)

ui/dialogs.py (DialogueLigne)
├─ Importe: config, utils.formatters, pandas
└─ Utilisé par: ui.main (ComptabiliteApp)

main.py (point d'entrée)
└─ Importe et utilise: ui.main (ComptabiliteApp)


═══════════════════════════════════════════════════════════════════════════

✅ PRINCIPES DE CONCEPTION APPLIQUÉS
═══════════════════════════════════════════════════════════════════════════

1. SEPARATION OF CONCERNS (SoC)
   ├─ Configuration séparée de la logique
   ├─ Donnees séparée de la présentation
   └─ UI séparée de la business logic

2. DRY (Don't Repeat Yourself)
   ├─ Fonctions de formatage centralisées dans utils
   └─ Config unique pour éviter duplication

3. SINGLE RESPONSIBILITY PRINCIPLE (SRP)
   ├─ DataManager : charge/sauvegarde
   ├─ PCGManager : charge le PCG
   ├─ ComptabiliteApp : interface
   ├─ DialogueLigne : dialogue d'édition
   └─ Formatters : utilitaires

4. OPEN/CLOSED PRINCIPLE
   ├─ Facile d'ajouter de nouveaux modules
   ├─ Extensible sans modification du core
   └─ Inheritance possible (ex: CustomApp extends ComptabiliteApp)

5. INVERSION OF DEPENDENCIES
   ├─ Import explicit plutôt que wildcard
   ├─ Dépendances claires
   └─ Facile à tester

═══════════════════════════════════════════════════════════════════════════

📝 FICHIERS DÉTAILS
═══════════════════════════════════════════════════════════════════════════

File                    Size(lines)  Purpose                  Imports
─────────────────────────────────────────────────────────────────────────
config.py              ~33          Configuration            None
main.py                ~30          Point d'entrée           ui.main
models/__init__.py     ~3           Export                   data
models/data.py         ~100         Gestion données         config, tk
ui/__init__.py         ~3           Export                   main, dialogs
ui/main.py             ~230         Interface princip.       config, models, utils
ui/dialogs.py          ~170         Dialogue d'édition       config, utils, pd
utils/__init__.py      ~3           Export                   formatters
utils/formatters.py    ~45          Formatage/Parsing        None

═══════════════════════════════════════════════════════════════════════════
""")


if __name__ == "__main__":
    main()

    print("\n" + "=" * 75)
    print("Pour plus de détails:")
    print("  • README.md - Architecture complète")
    print("  • EXAMPLES.py - 8 exemples pratiques")
    print("  • TEST_IMPORTS.py - Tests de validation")
    print("  • QUICKSTART.py - Menu interactif")
    print("=" * 75 + "\n")
