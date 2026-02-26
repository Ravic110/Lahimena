"""
EXAMPLES - Comment utiliser la structure refactorisée

Cet exemple montre différents cas d'usage
"""

# ========================================
# EXEMPLE 1: Accéder à la configuration
# ========================================


def exemple_config():
    """Accès simple à la configuration"""
    from config import CONFIG

    print("Fichier par défaut:", CONFIG["fichier_defaut"])
    print("Feuille Journal:", CONFIG["feuille_journal"])
    print("Colonnes:", CONFIG["colonnes_journal"])
    print("Largeurs:", CONFIG["largeurs_colonnes"])


# ========================================
# EXEMPLE 2: Charger des données Excel
# ========================================


def exemple_charger_donnees():
    """Chargement simple de données sans interface"""
    from config import CONFIG
    from models.data import DataManager, PCGManager

    # Charger le journal
    df = DataManager.charger_feuille(
        CONFIG["fichier_defaut"], CONFIG["feuille_journal"]
    )

    if df is not None:
        print(f"Nombre d'écritures: {len(df)}")
        print(df.head())

    # Charger le PCG
    comptes, numeros, dict_comptes = PCGManager.charger_pcg(CONFIG["fichier_defaut"])
    print(f"Nombre de comptes: {len(comptes)}")


# ========================================
# EXEMPLE 3: Sauvegarder des modifications
# ========================================


def exemple_sauvegarder():
    """Modification et sauvegarde de données"""
    import pandas as pd

    from config import CONFIG
    from models.data import DataManager

    # Charger les données
    df = DataManager.charger_feuille(
        CONFIG["fichier_defaut"], CONFIG["feuille_journal"]
    )

    if df is not None:
        # Ajouter une ligne
        nouvelle_ligne = {
            "Date": "13/02/2026",
            "Libellé": "Facture client",
            "DateValeur": "13/02/2026",
            "MontantDébit": 1000.0,
            "MontantCrédit": 0.0,
            "CompteDébit": "411",
            "CompteCrédit": "701",
            "Année": "2026",
        }
        df = pd.concat([df, pd.DataFrame([nouvelle_ligne])], ignore_index=True)

        # Sauvegarder
        success = DataManager.sauvegarder_df(
            df, CONFIG["fichier_defaut"], CONFIG["feuille_journal"]
        )
        print("Sauvegarde:", "OK" if success else "ÉCHOUÉE")


# ========================================
# EXEMPLE 4: Formater des montants
# ========================================


def exemple_formatage():
    """Utilisation des fonctions de formatage"""
    from utils.formatters import extraire_numero_compte, format_montant, parse_montant

    # Parser un montant depuis l'UI
    montant_str = "15 234,56"
    montant = parse_montant(montant_str)
    print(f"Parsé: {montant_str} → {montant}")

    # Formater pour l'affichage
    affiche = format_montant(montant)
    print(f"Formaté: {montant} → {affiche}")

    # Extraire numéro de compte
    compte_display = "401 - Client"
    numero = extraire_numero_compte(compte_display)
    print(f"Numéro de compte: {numero}")


# ========================================
# EXEMPLE 5: Utiliser l'application complète
# ========================================


def exemple_application():
    """Lancer l'application graphique"""
    import tkinter as tk

    from .ui.main import ComptabiliteApp

    root = tk.Tk()
    root.title("Application Comptable")
    root.geometry("1200x700")

    app = ComptabiliteApp(root)
    app.pack(side="top", fill="both", expand=True)

    root.mainloop()


# ========================================
# EXEMPLE 6: Script d'analyse de journal
# ========================================


def exemple_analyse_journal():
    """Analyser le journal sans interface"""
    import pandas as pd

    from config import CONFIG
    from models.data import DataManager

    df = DataManager.charger_feuille(
        CONFIG["fichier_defaut"], CONFIG["feuille_journal"]
    )

    if df is not None:
        # Statistiques
        total_debit = df["MontantDébit"].sum()
        total_credit = df["MontantCrédit"].sum()

        print(f"Total Débit: {total_debit}")
        print(f"Total Crédit: {total_credit}")
        print(f"Équilibré: {total_debit == total_credit}")

        # Balance par compte
        balance = df.groupby("CompteDébit")["MontantDébit"].sum()
        print("\nBalance Débit par compte:")
        print(balance)


# ========================================
# EXEMPLE 7: Créer un rapport personnalisé
# ========================================


def exemple_rapport():
    """Générer un rapport personnalisé"""
    import pandas as pd

    from config import CONFIG
    from models.data import DataManager
    from utils.formatters import format_montant

    df = DataManager.charger_feuille(
        CONFIG["fichier_defaut"], CONFIG["feuille_journal"]
    )

    if df is not None:
        # Rapport par année
        rapport = df.groupby("Année").agg(
            {"MontantDébit": "sum", "MontantCrédit": "sum"}
        )

        print("RAPPORT PAR ANNÉE")
        print("=" * 60)
        for annee, row in rapport.iterrows():
            print(f"Année {annee}:")
            print(f"  Débit: {format_montant(row['MontantDébit'])}")
            print(f"  Crédit: {format_montant(row['MontantCrédit'])}")
            print(
                f"  Solde: {format_montant(row['MontantDébit'] - row['MontantCrédit'])}"
            )
            print()


# ========================================
# EXEMPLE 8: Customiser l'application
# ========================================


def exemple_customisation():
    """Exemple de customisation simple"""
    import tkinter as tk

    from .ui.main import ComptabiliteApp

    class MonAppComptable(ComptabiliteApp):
        """Variante personnalisée de l'application"""

        def creer_menu_etats(self):
            """Override pour ajouter des états personnalisés"""
            menu = super().creer_menu_etats()
            menu.add_separator()
            menu.add_command(label="Rapport personnalisé", command=self.mon_rapport)
            return menu

        def mon_rapport(self):
            """Mon rapport personnalisé"""
            from tkinter import messagebox

            messagebox.showinfo("Mon Rapport", "Rapport en développement")

    root = tk.Tk()
    app = MonAppComptable(root)
    app.pack(side="top", fill="both", expand=True)
    root.mainloop()


if __name__ == "__main__":
    import sys

    print("=" * 70)
    print("EXEMPLES D'UTILISATION - apifactorise_refactored")
    print("=" * 70)

    exemples = {
        "1": ("Accéder à la configuration", exemple_config),
        "2": ("Charger des données", exemple_charger_donnees),
        "3": ("Sauvegarder des modifications", exemple_sauvegarder),
        "4": ("Formater des montants", exemple_formatage),
        "5": ("Lancer l'application", exemple_application),
        "6": ("Analyser le journal", exemple_analyse_journal),
        "7": ("Créer un rapport", exemple_rapport),
        "8": ("Customiser l'app", exemple_customisation),
    }

    print("\nExemples disponibles:")
    for key, (desc, _) in exemples.items():
        print(f"  {key}. {desc}")

    choix = input("\nQual exemple voulez-vous exécuter? (ou 'tous'): ").strip().lower()

    if choix == "tous":
        executed = []
        for key, (desc, func) in exemples.items():
            if key in ["5", "8"]:  # Passer les exemples qui lancent Tkinter
                continue
            try:
                print(f"\n{'='*70}")
                print(f"Exécution: {desc}")
                print(f"{'='*70}")
                func()
            except Exception as e:
                print(f"Erreur: {e}")
    elif choix in exemples:
        desc, func = exemples[choix]
        try:
            func()
        except Exception as e:
            print(f"Erreur lors de l'exécution: {e}")
