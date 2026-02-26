"""
╔════════════════════════════════════════════════════════════════════════╗
║                    GUIDE DE DÉMARRAGE RAPIDE                           ║
║              apifactorise_refactored - Quickstart                       ║
╚════════════════════════════════════════════════════════════════════════╝
"""

# ════════════════════════════════════════════════════════════════════════
# 🚀 OPTION 1 : LANCER L'APPLICATION GRAPHIQUE
# ════════════════════════════════════════════════════════════════════════


def option_1_app():
    """Lance l'interface graphique complète"""
    print("""
    OPTION 1 - Application Graphique
    ════════════════════════════════════════════════════════════════════════
    
    Commande:
        python main.py
    
    Résultat:
        ✅ Fenêtre de l'application comptable s'ouvre
        ✅ Interface complète avec menu, tableau, boutons d'action
        ✅ Chargement automatique du fichier EtatFidata.xlsx
        ✅ Gestion complète du journal comptable
    
    Fonctionnalités disponibles:
        • Ajouter / Modifier / Supprimer des écritures
        • Rechercher dans le journal
        • Voir la balance des comptes
        • États financiers (en développement)
        • Charger un autre fichier Excel
    """)


# ════════════════════════════════════════════════════════════════════════
# 📚 OPTION 2 : UTILISER LES MODULES DANS UN SCRIPT
# ════════════════════════════════════════════════════════════════════════


def option_2_script():
    """Exemple simple d'utilisation en script"""
    print("""
    OPTION 2 - Utilisation en Script
    ════════════════════════════════════════════════════════════════════════
    
    Exemple simple:
    
        from models.data import DataManager
        from config import CONFIG
        
        # Charger les données
        df = DataManager.charger_feuille(
            CONFIG['fichier_defaut'],
            CONFIG['feuille_journal']
        )
        
        # Afficher les 5 premières lignes
        print(df.head())
        
        # Sommes
        print(f"Total Débit: {df['MontantDébit'].sum()}")
        print(f"Total Crédit: {df['MontantCrédit'].sum()}")
    
    Résultat:
        ✅ Chargement des données sans interface graphique
        ✅ Manipulation libre en Python
        ✅ Création de rapports personnalisés
    """)


# ════════════════════════════════════════════════════════════════════════
# 🧪 OPTION 3 : TESTER LA STRUCTURE
# ════════════════════════════════════════════════════════════════════════


def option_3_test():
    """Valider que tout fonctionne"""
    print("""
    OPTION 3 - Tests d'Importation
    ════════════════════════════════════════════════════════════════════════
    
    Commande:
        python TEST_IMPORTS.py
    
    Résultat:
        ✅ Vérification de tous les imports
        ✅ Test de la configuration
        ✅ Test des utilitaires
        ✅ Rapport détaillé de la structure
    
    Exemple de sortie:
        ✅ CONFIG - Importé avec succès
        ✅ DataManager - Importé avec succès
        ✅ PCGManager - Importé avec succès
        ✅ DialogueLigne - Importé avec succès
        ✅ ComptabiliteApp - Importé avec succès
        ...
        Résultats: 8 ✅ | 0 ❌
    """)


# ════════════════════════════════════════════════════════════════════════
# 📖 OPTION 4 : VOIR LES EXEMPLES
# ════════════════════════════════════════════════════════════════════════


def option_4_examples():
    """Découvrir les 8 exemples pratiques"""
    print("""
    OPTION 4 - Exemples Pratiques
    ════════════════════════════════════════════════════════════════════════
    
    Commande:
        python EXAMPLES.py
    
    Vous pouvez choisir parmi 8 exemples:
    
        1. Accéder à la configuration
           → Voir tous les paramètres CONFIG
        
        2. Charger des données
           → Lire un fichier Excel sans UI
        
        3. Sauvegarder des modifications
           → Ajouter une ligne au journal et sauvegarder
        
        4. Formater des montants
           → Convertir entre formats (1234,56 ↔ 1 234,56)
        
        5. Lancer l'application
           → Ouvrir l'interface graphique
        
        6. Analyser le journal
           → Calculer totaux, balance, statistiques
        
        7. Créer un rapport
           → Générer un rapport par année
        
        8. Customiser l'app
           → Étendre ComptabiliteApp pour vos besoins
    
    Choisissez "tous" pour exécuter les 6 premiers exemples
    """)


# ════════════════════════════════════════════════════════════════════════
# 📂 OPTION 5 : EXPLORER LA STRUCTURE
# ════════════════════════════════════════════════════════════════════════


def option_5_explore():
    """Comprendre l'organisation du code"""
    print("""
    OPTION 5 - Exploration de la Structure
    ════════════════════════════════════════════════════════════════════════
    
    📁 apifactorise_refactored/
    
    🔧 Configuration:
       config.py
       → Tous les paramètres en un endroit
       → Fichiers, colonnes, largeurs, états financiers
    
    📊 Données:
       models/data.py
       → DataManager: Lire/écrire Excel
       → PCGManager: Charger le Plan Comptable
    
    🖥️  Interface:
       ui/main.py → Application principale
       ui/dialogs.py → Dialogs d'ajout/modificación
    
    🛠️  Utilitaires:
       utils/formatters.py
       → format_montant("1234,56")
       → parse_montant(1234.56)
       → extraire_numero_compte("401 - Client")
    
    📚 Documentation:
       README.md → Architecture complète
       EXAMPLES.py → 8 exemples pratiques
       SYNTHESE.py → Résumé de la refactorisation
       TEST_IMPORTS.py → Tests de validation
    """)


# ════════════════════════════════════════════════════════════════════════
# 💡 OPTION 6 : IMPORTER UN MODULE SPÉCIFIQUE
# ════════════════════════════════════════════════════════════════════════


def option_6_import():
    """Montrer comment importer chaque module"""
    print("""
    OPTION 6 - Guide d'Importation
    ════════════════════════════════════════════════════════════════════════
    
    🔧 Configuration:
       from config import CONFIG
       fichier = CONFIG['fichier_defaut']
    
    📊 Gestion de Données:
       from models.data import DataManager, PCGManager
       df = DataManager.charger_feuille("file.xlsx", "Sheet")
       comptes, nums, dict_c = PCGManager.charger_pcg("file.xlsx")
    
    💬 Dialogues:
       from .ui.dialogs import DialogueLigne
       dialogue = DialogueLigne(parent, df, comptes, numeros, ajout=True)
    
    🖥️  Interface Principale:
       from .ui.main import ComptabiliteApp
       app = ComptabiliteApp(root)
    
    🛠️  Utilitaires:
       from utils.formatters import format_montant, parse_montant
       montant = parse_montant("1234,56")
       texte = format_montant(montant)
    """)


# ════════════════════════════════════════════════════════════════════════
# 📋 OPTION 7 : VER LE CONTENU D'UN FICHIER
# ════════════════════════════════════════════════════════════════════════


def option_7_details():
    """Montrer les détails de chaque module"""
    print("""
    OPTION 7 - Contenu de Chaque Module
    ════════════════════════════════════════════════════════════════════════
    
    📄 config.py (33 lignes)
       ├─ fichier_defaut: "EtatFidata.xlsx"
       ├─ feuille_journal: "Journal"
       ├─ feuille_pcg: "PCG"
       ├─ colonnes_journal: [Date, Libellé, ...]
       ├─ largeurs_colonnes: {colonne: largeur}
       └─ etats_financiers: [Bilan, CR, etc]
    
    📄 models/data.py (100 lignes)
       ├─ DataManager
       │  ├─ charger_feuille(fichier, feuille)
       │  └─ sauvegarder_df(df, fichier, feuille)
       └─ PCGManager
          └─ charger_pcg(fichier)
    
    📄 ui/main.py (230 lignes)
       └─ ComptabiliteApp (interface complète)
          ├─ creer_menu()
          ├─ creer_interface()
          ├─ charger_donnees()
          ├─ afficher_donnees()
          ├─ filtrer_donnees()
          ├─ ajouter_ligne() / modifier / supprimer
          └─ afficher_balance()
    
    📄 ui/dialogs.py (170 lignes)
       └─ DialogueLigne (modal pour éditer une écriture)
          ├─ creer_widgets()
          ├─ valider()
          └─ Événements: on_key, on_select
    
    📄 utils/formatters.py (45 lignes)
       ├─ format_montant(montant)
       ├─ parse_montant(value_str)
       └─ extraire_numero_compte(compte_display)
    
    📄 main.py (30 lignes)
       └─ Point d'entrée pour lancer l'app
    """)


# ════════════════════════════════════════════════════════════════════════
# MENU PRINCIPAL
# ════════════════════════════════════════════════════════════════════════


def afficher_menu():
    """Affiche le menu principal"""
    print("""
╔═════════════════════════════════════════════════════════════════════════╗
║                    GUIDE DE DÉMARRAGE RAPIDE                           ║
║              apifactorise_refactored - Quickstart                       ║
║                     Choisissez une option (1-7)                        ║
╚═════════════════════════════════════════════════════════════════════════╝

  1️⃣  Lancer l'application graphique (GUI)
  2️⃣  Utiliser les modules dans un script
  3️⃣  Tester la structure (import tests)
  4️⃣  Voir les 8 exemples pratiques
  5️⃣  Explorer la structure des répertoires
  6️⃣  Guide d'importation des modules
  7️⃣  Détails du contenu de chaque fichier
  
  0️⃣  Quitter
""")


def main():
    """Menu interactif"""
    options = {
        "1": ("Lancer l'application graphique", option_1_app),
        "2": ("Utiliser en script", option_2_script),
        "3": ("Tester la structure", option_3_test),
        "4": ("Voir les exemples", option_4_examples),
        "5": ("Explorer la structure", option_5_explore),
        "6": ("Guide d'importation", option_6_import),
        "7": ("Détails des modules", option_7_details),
    }

    while True:
        afficher_menu()
        choix = input("\nQuel option (0-7)? ").strip()

        if choix == "0":
            print("\n👋 Au revoir!\n")
            break
        elif choix in options:
            titre, func = options[choix]
            print()
            func()
            input("\nAppuyez sur Entrée pour continuer... ")
            print("\n" + "=" * 75 + "\n")
        else:
            print("\n❌ Option invalide! Veuillez choisir 0-7.\n")


if __name__ == "__main__":
    main()
