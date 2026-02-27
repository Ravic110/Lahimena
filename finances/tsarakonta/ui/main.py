"""
Interface principale de l'application comptable
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os

from finances.tsarakonta.config import CONFIG
from finances.tsarakonta.models.data import DataManager, PCGManager
from finances.tsarakonta.utils.formatters import format_montant
from .dialogs import DialogueLigne
from .etat_resultat import CompteResultatNatureWindow
from .etat_resultat_fonction import CompteResultatFonctionWindow
from .bilan_actif import BilanActifWindow
from .bilan_passif import BilanPassifWindow
from .flux_tresorerie_direct import FluxTresorerieDirectWindow
from .flux_tresorerie_indirect import FluxTresorerieIndirectWindow
from .ratios import RatioResultatNatureWindow, RatioResultatFonctionWindow, RatioBilanWindow
from .settings import SettingsWindow


class ComptabiliteApp(tk.Frame):
    """Application principale de gestion comptable"""

    def __init__(self, parent, fichier_excel=None, etat_initial=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.master = parent
        self.fichier_excel = fichier_excel or CONFIG.get('fichier_defaut') or "LivreCompta.xlsx"
        self.df = None
        self.pcg_comptes, self.pcg_numeros, self.pcg_dict = [], [], {}

        self.creer_menu()
        self.creer_interface()
        self.charger_pcg()
        self.charger_donnees()

        if etat_initial:
            self.afficher_etat(etat_initial)
    
    # ==================== MENUS ====================
    def creer_menu(self):
        """Crée la barre de menu"""
        menu_target = self.master
        if not isinstance(menu_target, (tk.Tk, tk.Toplevel)):
            menu_target = self.winfo_toplevel()
        self._menu_parent = menu_target

        menubar = tk.Menu(self._menu_parent)
        try:
            self._menu_parent.configure(menu=menubar)
        except Exception:
            return
        
        menubar.add_cascade(label="Fichier", menu=self.creer_menu_fichier())
        # états financiers handled via accueil page dropdown
        menubar.add_cascade(label="Ratio", menu=self.creer_menu_ratios())
    
    def creer_menu_fichier(self):
        """Crée le menu Fichier"""
        menu = tk.Menu(getattr(self, "_menu_parent", self.master), tearoff=0)
        # Paramétrage (entête société)
        menu.add_command(label="Paramétrage", command=lambda: SettingsWindow(self.master))
        # Ouvrir le dossier d'exports
        def _ouvrir_exports_folder():
            folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'EtatFiFolder', 'exports')
            os.makedirs(folder, exist_ok=True)
            try:
                os.startfile(folder)
            except Exception:
                messagebox.showinfo('Info', f'Emplacement: {folder}')
        menu.add_command(label="Ouvrir dossier exports", command=_ouvrir_exports_folder)
        menu.add_command(label="Charger Excel", command=self.charger_fichier)
        menu.add_command(label="Sauvegarder", command=self.sauvegarder)
        menu.add_separator()
        menu.add_command(label="Quitter", command=self.master.quit)
        return menu
    
    def creer_menu_etats(self):
        """Crée le menu États Financiers"""
        menu = tk.Menu(getattr(self, "_menu_parent", self.master), tearoff=0)
        for etat in CONFIG['etats_financiers']:
            menu.add_command(label=etat, command=lambda e=etat: self.afficher_etat(e))
        menu.add_separator()
        menu.add_cascade(label="Grand Livre", menu=self.creer_menu_grand_livre())
        return menu

    def creer_menu_grand_livre(self):
        """Crée le sous-menu Grand Livre"""
        menu = tk.Menu(getattr(self, "_menu_parent", self.master), tearoff=0)
        menu.add_command(label="Grand Livre comptable", command=self.afficher_grand_livre)
        return menu

    def creer_menu_ratios(self):
        """Crée le menu Ratios"""
        menu = tk.Menu(getattr(self, "_menu_parent", self.master), tearoff=0)
        menu.add_command(
            label="Ratio Compte de résultat par nature",
            command=lambda: self.afficher_ratio("Ratio Compte de résultat par nature")
        )
        menu.add_command(
            label="Ration compte de résultat par fonction",
            command=lambda: self.afficher_ratio("Ration compte de résultat par fonction")
        )
        menu.add_command(
            label="Ratio Bilan",
            command=lambda: self.afficher_ratio("Ratio Bilan")
        )
        return menu
    
    # ==================== INTERFACE ====================
    def creer_interface(self):
        """Crée l'interface principale"""
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.creer_accueil_etats_financiers(main_frame)
        self.creer_barre_recherche(main_frame)
        self.creer_treeview(main_frame)
        self.creer_boutons(main_frame)
        self.creer_stats(main_frame)

    def creer_accueil_etats_financiers(self, parent):
        """Crée une page d'accueil des états financiers avec boutons directs."""
        accueil_frame = ttk.LabelFrame(
            parent, text='Accueil États Financiers', padding='10'
        )
        accueil_frame.pack(fill=tk.X, pady=(0, 10))

        # Dropdown menu replicating bar menu for easy access
        mb_frame = ttk.Frame(accueil_frame)
        mb_frame.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(mb_frame, text='Menu déroulant').pack(side=tk.LEFT)
        mb = ttk.Menubutton(mb_frame, text='États financiers', width=28)
        dropdown_menu = tk.Menu(mb, tearoff=0)
        for etat in CONFIG.get('etats_financiers', []):
            dropdown_menu.add_command(
                label=etat,
                command=lambda e=etat: self.afficher_etat(e),
            )
        dropdown_menu.add_separator()
        # reuse existing submenu builder for grand livre
        dropdown_menu.add_cascade(label='Grand Livre', menu=self.creer_menu_grand_livre())
        mb['menu'] = dropdown_menu
        mb.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Label(
            accueil_frame,
            text='Accès rapide aux états financiers (boutons).',
        ).pack(anchor='w', pady=(0, 8))

        etats_frame = ttk.Frame(accueil_frame)
        etats_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(etats_frame, text='États financiers').pack(anchor='w')

        grid_etats = ttk.Frame(etats_frame)
        grid_etats.pack(fill=tk.X, pady=(4, 0))
        for i, etat in enumerate(CONFIG.get('etats_financiers', [])):
            row = i // 3
            col = i % 3
            ttk.Button(
                grid_etats,
                text=etat,
                command=lambda e=etat: self.afficher_etat(e),
                width=40,
            ).grid(row=row, column=col, padx=4, pady=4, sticky='ew')
            grid_etats.columnconfigure(col, weight=1)

        ratio_frame = ttk.Frame(accueil_frame)
        ratio_frame.pack(fill=tk.X, pady=(6, 8))
        ttk.Label(ratio_frame, text='Ratios').pack(anchor='w')

        grid_ratios = ttk.Frame(ratio_frame)
        grid_ratios.pack(fill=tk.X, pady=(4, 0))
        ratios = [
            'Ratio Compte de résultat par nature',
            'Ration compte de résultat par fonction',
            'Ratio Bilan',
        ]
        for i, ratio in enumerate(ratios):
            ttk.Button(
                grid_ratios,
                text=ratio,
                command=lambda r=ratio: self.afficher_ratio(r),
                width=40,
            ).grid(row=0, column=i, padx=4, pady=4, sticky='ew')
            grid_ratios.columnconfigure(i, weight=1)

        grand_livre_frame = ttk.Frame(accueil_frame)
        grand_livre_frame.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(grand_livre_frame, text='Grand Livre').pack(anchor='w')
        ttk.Button(
            grand_livre_frame,
            text='Grand Livre comptable',
            command=self.afficher_grand_livre,
            width=40,
        ).pack(anchor='w', pady=(4, 0))
    
    def creer_barre_recherche(self, parent):
        """Crée la barre de recherche"""
        search_frame = ttk.Frame(parent)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(search_frame, text="Rechercher:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.search_entry.bind('<KeyRelease>', self.filtrer_donnees)
        
        ttk.Button(search_frame, text="Actualiser", command=self.charger_donnees).pack(side=tk.LEFT)
    
    def creer_treeview(self, parent):
        """Crée le tableau de visualisation des données"""
        columns = CONFIG['colonnes_journal']
        self.tree = ttk.Treeview(parent, columns=columns, show='headings', height=20)
        
        for col in columns:
            self.tree.heading(col, text=col)
            width = CONFIG['largeurs_colonnes'].get(col, 80)
            anchor = tk.E if col in ['MontantDébit', 'MontantCrédit'] else tk.W
            self.tree.column(col, width=width, anchor=anchor)
        
        v_scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def creer_boutons(self, parent):
        """Crée les boutons d'action"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(side=tk.RIGHT, padx=(10, 0))
        
        actions = [
            ("Ajouter", self.ajouter_ligne),
            ("Modifier", self.modifier_ligne),
            ("Supprimer", self.supprimer_ligne),
            ("Balance", self.afficher_balance)
        ]
        
        for text, cmd in actions:
            ttk.Button(button_frame, text=text, command=cmd).pack(pady=5, fill=tk.X)
    
    def creer_stats(self, parent):
        """Crée le panneau de statistiques"""
        stats_frame = ttk.LabelFrame(parent, text="Résumé du journal", padding="10")
        stats_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(14, 0))
        
        self.total_debit = tk.StringVar(value="Total des débits: 0,00")
        self.total_credit = tk.StringVar(value="Total des crédits: 0,00")
        self.solde_final = tk.StringVar(value="Solde: 0,00")
        
        resume_row = ttk.Frame(stats_frame)
        resume_row.pack(fill=tk.X, pady=(4, 2))
        resume_row.columnconfigure(0, weight=1)
        resume_row.columnconfigure(1, weight=1)
        resume_row.columnconfigure(2, weight=1)

        ttk.Label(resume_row, textvariable=self.total_debit, font=('Segoe UI', 10, 'bold'), anchor='w').grid(row=0, column=0, sticky='ew', padx=(8, 8))
        ttk.Label(resume_row, textvariable=self.total_credit, font=('Segoe UI', 10, 'bold'), anchor='center').grid(row=0, column=1, sticky='ew', padx=(8, 8))
        ttk.Label(resume_row, textvariable=self.solde_final, font=('Segoe UI', 10, 'bold'), anchor='e').grid(row=0, column=2, sticky='ew', padx=(8, 8))
    
    # ==================== GESTION DONNÉES ====================
    def charger_pcg(self):
        """Charge le Plan Comptable Général"""
        self.pcg_comptes, self.pcg_numeros, self.pcg_dict = PCGManager.charger_pcg(self.fichier_excel)
    
    def charger_donnees(self):
        """Charge les données du journal"""
        self.df = DataManager.charger_feuille(self.fichier_excel, CONFIG['feuille_journal'])
        
        if self.df is None:
            self.df = pd.DataFrame(columns=CONFIG['colonnes_journal'])
            self.sauvegarder()
        
        # Ajouter colonnes manquantes
        for col in CONFIG['colonnes_journal']:
            if col not in self.df.columns:
                self.df[col] = 0.0 if col in ['MontantDébit', 'MontantCrédit'] else ''
        
        self.afficher_donnees(self.df)
        self.mettre_a_jour_stats()
    
    def sauvegarder(self):
        """Sauvegarde les données dans Excel"""
        DataManager.sauvegarder_df(self.df, self.fichier_excel, CONFIG['feuille_journal'])
    
    def afficher_donnees(self, df):
        """Affiche les données dans le tableau"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for _, row in df.iterrows():
            values = (
                row.get('Date', ''),
                row.get('Libellé', ''),
                row.get('DateValeur', ''),
                format_montant(row.get('MontantDébit', 0)),
                format_montant(row.get('MontantCrédit', 0)),
                row.get('CompteDébit', ''),
                row.get('CompteCrédit', ''),
                row.get('Année', '')
            )
            self.tree.insert('', tk.END, values=values)
    
    def filtrer_donnees(self, event=None):
        """Filtre les données selon la recherche"""
        if self.df is not None:
            recherche = self.search_var.get().lower()
            df_filtre = self.df[self.df.astype(str).apply(
                lambda x: x.str.contains(recherche, na=False).any(), axis=1)]
            self.afficher_donnees(df_filtre)
    
    def mettre_a_jour_stats(self):
        """Met à jour les statistiques"""
        if self.df is not None:
            total_debit = self.df['MontantDébit'].sum()
            total_credit = self.df['MontantCrédit'].sum()
            solde = total_debit - total_credit
            
            self.total_debit.set(f"Total des débits: {format_montant(total_debit)}")
            self.total_credit.set(f"Total des crédits: {format_montant(total_credit)}")
            self.solde_final.set(f"Solde: {format_montant(solde)}")
    
    # ==================== ACTIONS ====================
    def ajouter_ligne(self):
        """Ouvre le dialogue pour ajouter une ligne"""
        self.ouvrir_dialogue(True)
    
    def modifier_ligne(self):
        """Ouvre le dialogue pour modifier une ligne"""
        selection = self.tree.selection()
        if selection:
            index = self.tree.index(selection[0])
            donnees = self.tree.item(selection[0])['values']
            self.ouvrir_dialogue(False, index, donnees)
        else:
            messagebox.showwarning("Attention", "Sélectionnez une écriture")
    
    def supprimer_ligne(self):
        """Supprime la ligne sélectionnée"""
        selection = self.tree.selection()
        if selection and messagebox.askyesno("Confirmer", "Supprimer l'écriture?"):
            index = self.tree.index(selection[0])
            self.df = self.df.drop(index).reset_index(drop=True)
            self.sauvegarder()
            self.charger_donnees()
    
    def ouvrir_dialogue(self, ajout, index=None, donnees=None):
        """Ouvre le dialogue d'édition"""
        dialogue = DialogueLigne(self.master, self.df, self.pcg_comptes, self.pcg_numeros, ajout, index, donnees)
        self.master.wait_window(dialogue)
        
        if dialogue.resultat:
            self.df = dialogue.df_modifie
            self.afficher_donnees(self.df)
            self.mettre_a_jour_stats()
            self.sauvegarder()
    
    def charger_fichier(self):
        """Ouvre un dialogue pour charger un fichier Excel"""
        fichier = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if fichier:
            self.fichier_excel = fichier
            self.charger_pcg()
            self.charger_donnees()
    
    def afficher_balance(self):
        """Affiche la balance du journal"""
        if self.df is not None:
            balance_debit = self.df.groupby('CompteDébit')['MontantDébit'].sum()
            balance_credit = self.df.groupby('CompteCrédit')['MontantCrédit'].sum()
            balance = pd.DataFrame({'Débit': balance_debit, 'Crédit': balance_credit}).fillna(0)
            messagebox.showinfo("Balance Journal", balance.to_string())
    
    def afficher_etat(self, type_etat):
        """Affiche un état financier"""
        if type_etat == "Bilan actif":
            BilanActifWindow(self.master, self.df)
            return

        if type_etat == "Bilan passif":
            BilanPassifWindow(self.master, self.df)
            return

        # Ouvrir la fenêtre dédiée pour le Compte de résultat par nature
        if type_etat == "Compte de résultat par nature":
            CompteResultatNatureWindow(self.master, self.df, self.pcg_dict)
            return

        if type_etat == "Compte de résultat par fonction":
            CompteResultatFonctionWindow(self.master, self.df)
            return

        if type_etat == "Tableau de flux de trésorerie (méthode directe)":
            FluxTresorerieDirectWindow(self.master)
            return

        if type_etat == "Tableau de flux de trésorerie (méthode indirecte)":
            FluxTresorerieIndirectWindow(self.master)
            return

        messagebox.showinfo("États", f"État '{type_etat}' en développement")

    def afficher_grand_livre(self):
        """Affiche le grand livre comptable sommaire par compte"""
        if self.df is None or self.df.empty:
            messagebox.showinfo("Grand Livre", "Aucune donnée disponible")
            return

        debit = self.df.groupby('CompteDébit')['MontantDébit'].sum()
        credit = self.df.groupby('CompteCrédit')['MontantCrédit'].sum()
        comptes = sorted(set(list(debit.index) + list(credit.index)))

        ledger = pd.DataFrame(index=comptes, columns=['Débit', 'Crédit']).fillna(0)
        for c in comptes:
            ledger.at[c, 'Débit'] = float(debit.get(c, 0))
            ledger.at[c, 'Crédit'] = float(credit.get(c, 0))
        ledger['Solde'] = ledger['Débit'] - ledger['Crédit']

        # Calcul des totaux pour la balance générale
        totaux = ledger[['Débit', 'Crédit', 'Solde']].sum()
        total_debit = float(totaux['Débit'])
        total_credit = float(totaux['Crédit'])
        total_solde = float(totaux['Solde'])

        win = tk.Toplevel(self.master)
        win.title("Grand Livre Comptable")
        win.geometry("800x500")

        # Affichage de la balance générale en en-tête
        header = ttk.Frame(win, padding=6)
        header.pack(fill=tk.X)
        ttk.Label(header, text=f"Total Débit: {format_montant(total_debit)}").pack(side=tk.LEFT, padx=(10,5))
        ttk.Label(header, text=f"Total Crédit: {format_montant(total_credit)}").pack(side=tk.LEFT, padx=(10,5))
        ttk.Label(header, text=f"Total Solde: {format_montant(total_solde)}").pack(side=tk.LEFT, padx=(10,5))

        # Bouton d'export
        btn_frame = ttk.Frame(win, padding=6)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(btn_frame, text="Exporter en Excel", command=lambda: self.exporter_grand_livre(ledger)).pack(side=tk.LEFT, padx=5)

        cols = ('Compte', 'Débit', 'Crédit', 'Solde')
        tree = ttk.Treeview(win, columns=cols, show='headings')
        for col in cols:
            tree.heading(col, text=col)
            anchor = tk.E if col in ['Débit', 'Crédit', 'Solde'] else tk.W
            tree.column(col, width=180, anchor=anchor)

        for compte, row in ledger.iterrows():
            tree.insert('', tk.END, values=(compte,
                                            format_montant(row['Débit']),
                                            format_montant(row['Crédit']),
                                            format_montant(row['Solde'])))

        vsb = ttk.Scrollbar(win, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def afficher_ratio(self, type_ratio):
        """Affiche la fenêtre de ratios demandée."""
        if type_ratio == "Ratio Compte de résultat par nature":
            RatioResultatNatureWindow(self.master, self.df)
            return

        if type_ratio in ["Ration compte de résultat par fonction", "Ratio compte de résultat par fonction"]:
            RatioResultatFonctionWindow(self.master, self.df)
            return

        if type_ratio == "Ratio Bilan":
            RatioBilanWindow(self.master, self.df)
            return

        messagebox.showinfo("Ratios", f"{type_ratio} indisponible")

    def exporter_grand_livre(self, ledger_df):
        """Exporte le grand livre en Excel dans LivreCompta.xlsx feuille 'GrandLivre'"""
        try:
            base_dir = os.path.dirname(os.path.dirname(__file__))
            excel_file = os.path.join(base_dir, 'EtatFiFolder', 'LivreCompta.xlsx')
            os.makedirs(os.path.dirname(excel_file), exist_ok=True)

            # Réinitialiser l'index pour que 'Compte' devienne une colonne
            export_df = ledger_df.reset_index()
            export_df.rename(columns={'index': 'Compte'}, inplace=True)

            # Vérifier si le fichier existe
            if os.path.exists(excel_file):
                # Charger le fichier existant
                with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    export_df.to_excel(writer, sheet_name='GrandLivre', index=False)
            else:
                # Créer un nouveau fichier
                export_df.to_excel(excel_file, sheet_name='GrandLivre', index=False)

            messagebox.showinfo("Export réussi", f"Grand Livre exporté dans:\n{excel_file}\nFeuille: GrandLivre")
        except PermissionError:
            messagebox.showerror("Erreur", "FERMEZ EXCEL avant d'exporter !")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'export: {str(e)}")

