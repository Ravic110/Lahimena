"""
Dialogues pour l'ajout et modification d'écritures comptables
"""

import os
import sys
import tkinter as tk
from datetime import datetime
from tkinter import ttk

# Ajout du chemin parent pour les imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd

from finances.tsarakonta.config import CONFIG
from finances.tsarakonta.utils.formatters import extraire_numero_compte, parse_montant


class DialogueLigne(tk.Toplevel):
    """Dialogue pour créer ou modifier une écriture comptable"""

    def __init__(
        self, parent, df, pcg_comptes, pcg_numeros, ajout, index=None, donnees=None
    ):
        """
        Args:
            parent: fenêtre parent
            df: DataFrame contenant les données
            pcg_comptes: liste des comptes au format "numero - libelle"
            pcg_numeros: liste des numéros de comptes
            ajout: True pour ajouter, False pour modifier
            index: index de la ligne à modifier (si applicable)
            donnees: données de la ligne à éditer (si applicable)
        """
        super().__init__(parent)
        self.df_modifie = df.copy()
        self.ajout = ajout
        self.index = index
        self.resultat = False
        self.pcg_comptes = pcg_comptes
        self.pcg_numeros = pcg_numeros

        self.title("Ajouter écriture" if self.ajout else "Modifier écriture")
        self.geometry("550x500")
        self.transient(parent)
        self.grab_set()

        self.creer_widgets(donnees)

    def creer_widgets(self, donnees):
        """Crée les widgets du dialogue"""
        frame = ttk.Frame(self, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        self.vars = {}

        # Champs texte
        champs_texte = [
            ("Date:", 0),
            ("Libellé:", 1),
            ("Date Valeur:", 2),
            ("Année:", 7),
        ]
        row = 0

        for label, idx in champs_texte:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky=tk.W, pady=5)
            var = tk.StringVar(
                value=donnees[idx] if donnees and idx < len(donnees) else ""
            )
            self.vars[label.strip(":")] = var
            ttk.Entry(frame, textvariable=var, width=25).grid(
                row=row, column=1, sticky=(tk.W, tk.E), pady=5
            )
            row += 1

        # Montants
        self.creer_champ_montant(
            frame, row, "Montant Débit:", donnees, 3, "montant_debit_var"
        )
        row += 1
        self.creer_champ_montant(
            frame, row, "Montant Crédit:", donnees, 4, "montant_credit_var"
        )
        row += 1

        # Comptes
        self.creer_combobox_compte(frame, row, "Compte Débit:", donnees, 5, "debit")
        row += 1
        self.creer_combobox_compte(frame, row, "Compte Crédit:", donnees, 6, "credit")
        row += 1

        if not donnees:
            self.init_dates()

        self.creer_boutons(frame, row)
        frame.columnconfigure(1, weight=1)

    def creer_champ_montant(self, frame, row, label, donnees, idx, var_name):
        """Crée un champ montant"""
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky=tk.W, pady=5)
        setattr(self, var_name, tk.StringVar(value=donnees[idx] if donnees else "0,00"))
        ttk.Entry(frame, textvariable=getattr(self, var_name), width=20).grid(
            row=row, column=1, sticky=(tk.W, tk.E), pady=5
        )

    def creer_combobox_compte(self, frame, row, label, donnees, idx, type_compte):
        """Crée une combobox pour sélectionner un compte"""
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky=tk.W, pady=5)
        var_name = f"compte_{type_compte}_var"
        combo_name = f"combo_{type_compte}"

        setattr(self, var_name, tk.StringVar(value=donnees[idx] if donnees else ""))
        values = self.pcg_comptes if self.pcg_comptes else ["Saisir manuellement"]

        combo = ttk.Combobox(
            frame, textvariable=getattr(self, var_name), values=values, width=25
        )
        setattr(self, combo_name, combo)
        combo.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)

        if self.pcg_comptes:
            combo.bind("<KeyRelease>", self._creer_handler_key(type_compte))
            combo.bind("<<ComboboxSelected>>", self._creer_handler_select(type_compte))

    def _creer_handler_key(self, type_compte):
        """Factory pour créer le handler KeyRelease"""

        def handler(event):
            combobox = getattr(self, f"combo_{type_compte}")
            recherche = getattr(self, f"compte_{type_compte}_var").get()
            self.filtrer_combobox(combobox, recherche)

        return handler

    def _creer_handler_select(self, type_compte):
        """Factory pour créer le handler ComboboxSelected"""

        def handler(event):
            var = getattr(self, f"compte_{type_compte}_var")
            val = var.get()
            if " - " in val:
                var.set(extraire_numero_compte(val))

        return handler

    def filtrer_combobox(self, combobox, recherche):
        """Filtre les options de la combobox"""
        if self.pcg_comptes and recherche.strip():
            filtre = [
                c for c in self.pcg_comptes if recherche.strip() in c.split(" - ")[0]
            ]
            combobox["values"] = filtre if filtre else self.pcg_comptes
        else:
            combobox["values"] = self.pcg_comptes

    def creer_boutons(self, frame, row):
        """Crée les boutons Valider et Annuler"""
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=row + 1, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Valider", command=self.valider).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(btn_frame, text="Annuler", command=self.destroy).pack(
            side=tk.LEFT, padx=5
        )

    def init_dates(self):
        """Initialise les champs date avec la date actuelle"""
        today = datetime.now().strftime("%d/%m/%Y")
        self.vars["Date"].set(today)
        self.vars["Date Valeur"].set(today)
        self.vars["Année"].set(datetime.now().strftime("%Y"))

    def valider(self):
        """Valide et enregistre les données"""
        try:
            ligne = self.preparer_ligne()
            if self.ajout:
                self.df_modifie = pd.concat(
                    [self.df_modifie, pd.DataFrame([ligne])], ignore_index=True
                )
            else:
                for key, value in ligne.items():
                    self.df_modifie.iloc[
                        self.index, self.df_modifie.columns.get_loc(key)
                    ] = value
            self.resultat = True
            self.destroy()
        except ValueError:
            from tkinter import messagebox

            messagebox.showerror("Erreur", "Vérifiez les montants (format: 1234,56)")

    def preparer_ligne(self):
        """Prépare le dictionnaire de la ligne pour pandas"""
        return {
            "Date": self.vars["Date"].get(),
            "Libellé": self.vars["Libellé"].get(),
            "DateValeur": self.vars["Date Valeur"].get(),
            "MontantDébit": parse_montant(self.montant_debit_var.get()),
            "MontantCrédit": parse_montant(self.montant_credit_var.get()),
            "CompteDébit": self.compte_debit_var.get(),
            "CompteCrédit": self.compte_credit_var.get(),
            "Année": self.vars["Année"].get(),
        }
