"""
Module fenêtre : Compte de résultat par nature (PCG 2005 Madagascar)
Affiche un tableau comparatif 2025 / 2024 (affichage simple, sans export).
"""

import os
import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd

from finances.tsarakonta.config import CONFIG
from finances.tsarakonta.models.data import DataManager, PCGManager
from finances.tsarakonta.utils.exports import export_treeview_to_excel, export_treeview_to_pdf
from finances.tsarakonta.utils.formatters import format_montant

from .settings import format_header_text, load_header_settings, save_header_settings

STATE_FOLDER = os.path.join(os.path.dirname(os.path.dirname(__file__)), "EtatFiFolder")
MAPPING_CSV = os.path.join(STATE_FOLDER, "mapping_resultat_nature.csv")


class CompteResultatNatureWindow(tk.Toplevel):
    """Fenêtre du compte de résultat par nature"""

    def __init__(self, parent, df=None, pcg_dict=None):
        super().__init__(parent)
        self.parent = parent

        # Charger le journal depuis LivreCompta.xlsx
        livre_file = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), "LivreCompta.xlsx"
        )
        if os.path.exists(livre_file):
            self.df = DataManager.charger_feuille(livre_file, CONFIG["feuille_journal"])
            if self.df is None:
                self.df = pd.DataFrame(columns=CONFIG["colonnes_journal"])
        else:
            # Fallback sur le fichier par défaut si LivreCompta.xlsx n'existe pas
            self.df = (
                df
                if df is not None
                else pd.DataFrame(columns=CONFIG["colonnes_journal"])
            )

        self.pcg_dict = pcg_dict or {}

        os.makedirs(STATE_FOLDER, exist_ok=True)
        self._ensure_mapping()

        self.title("Compte de résultat par nature")
        self.geometry("900x600")

        annees_data = (
            self.df["Année"].dropna().astype(str).unique().tolist()
            if "Année" in self.df.columns
            else []
        )
        annees_fixes = [str(y) for y in range(2020, 2031)]
        toutes_annees = set(annees_fixes) | set(annees_data)

        def _sort_key(v):
            try:
                return int(v)
            except Exception:
                return -9999

        self.annees = sorted(toutes_annees, key=_sort_key, reverse=True)
        self.header_settings = load_header_settings()
        self.header_text = format_header_text(self.header_settings)

        self._build_ui()
        self._charger_tableau()
        # charger preferences (entête)
        self._load_settings()

    def _ensure_mapping(self):
        """Crée un fichier mapping CSV par défaut si absent."""
        if not os.path.exists(MAPPING_CSV):
            default = [
                # ordre, rubrique, prefixes (séparateur ;), type
                (
                    1,
                    "Chiffres d'affaires",
                    "70;701;702;703;704;705;706;707;708;7082;7083;7085;7086;7088",
                    "produit",
                ),
                (3, "Production Stockée", "71;713;714", "produit"),
                (4, "Production immobilisée", "72;721;722", "produit"),
                (5, "I- PRODUCTION DE L'EXERCICE", "", "subtotal"),
                (
                    6,
                    "Achat consommés",
                    "60;601;602;6021;6022;60221;60222;60223;60225;6023;60231;60235;60237;604;605;606;6061;6062;6063;6064;6068;607;608",
                    "charge",
                ),
                (
                    7,
                    "Services extérieurs et autres consommations",
                    "61;611;612;613;6132;6135;6136.614;615;616;617;618;62;621;622;623;624;6241;6242.625;626;627;628",
                    "charge",
                ),
                (8, "II- CONSOMMATION DE L'EXERCICE", "", "subtotal"),
                (9, "III- VALEUR AJOUTEE DE L'EXPLOITATION", "", "subtotal"),
                (10, "Charges de personnel", "64;641;644;645;646;647;648", "charge"),
                (11, "Impôts, taxes et versements assimilés", "63;631;635", "charge"),
                (12, "IV- EXCEDENT BRUTE D'EXPLOITATION", "", "subtotal"),
                (
                    13,
                    "Autres produits opérationnels",
                    "74;741;748;75;751;752;753;754;755;756;757;758",
                    "produit",
                ),
                (
                    14,
                    "Autres charges opérationnelles",
                    "65;651;652;653;654;655;656;657;658",
                    "charge",
                ),
                (15, "V- RESULTAT OPERATIONNEL", "", "subtotal"),
                (16, "Dotations aux amortissements", "68;681;685", "charge"),
                (
                    17,
                    "Produits financiers",
                    "76;761;762;763;764;766;767;768",
                    "produit",
                ),
                (18, "Charges financières", "66;661;664;665;666;667;668", "charge"),
                (19, "VI- RESULTAT FINANCIER", "", "subtotal"),
                (20, "VII- RESULTAT AVANT IMPOT", "", "subtotal"),
                (21, "Impôts exigibles sur résultats", "69;695;698", "charge"),
                (22, "Impôts différés (variations)", "692;693", "charge"),
                (23, "TOTAL PRODUITS DES ACTIVITES ORDINAIRES", "", "subtotal"),
                (24, "TOTAL CHARGES DES ACTIVITES ORDINAIRES", "", "subtotal"),
                (25, "Eléments extraordinaires(produits)", "77", "produit"),
                (26, "Eléments extraordinaires(charges)", "67", "charge"),
                (27, "IX- RESULTAT EXTRAORDINAIRE", "", "subtotal"),
                (28, "X- RESULTAT NET DE L'EXERCICE", "", "subtotal"),
            ]
            df_map = pd.DataFrame(
                default, columns=["ordre", "rubrique", "prefixes", "type"]
            )
            df_map.to_csv(MAPPING_CSV, index=False, encoding="utf-8")

    def _build_ui(self):
        # Header personnalisation
        hdr_frame = ttk.LabelFrame(self, text="En-tête / Options", padding=8)
        hdr_frame.pack(fill=tk.X, padx=8, pady=6)

        ttk.Label(hdr_frame, text="Titre:").grid(row=0, column=0, sticky=tk.W)
        self.titre_var = tk.StringVar(value="Compte de résultat par nature")
        ttk.Entry(hdr_frame, textvariable=self.titre_var, width=40).grid(
            row=0, column=1, sticky=tk.W
        )

        if self.header_text:
            self.header_info_var = tk.StringVar(value=self.header_text)
            ttk.Label(
                hdr_frame, textvariable=self.header_info_var, justify=tk.LEFT
            ).grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=(6, 0))

        ttk.Label(hdr_frame, text="Arrêté au:").grid(
            row=0, column=2, sticky=tk.W, padx=(10, 0)
        )
        self.annee_var = tk.StringVar(value="2025")
        self.annee_combo = ttk.Combobox(
            hdr_frame,
            values=["Both"] + self.annees,
            textvariable=self.annee_var,
            state="readonly",
            width=12,
        )
        self.annee_combo.grid(row=0, column=3, sticky=tk.W)
        self.annee_combo.bind("<<ComboboxSelected>>", lambda e: self._charger_tableau())

        ttk.Button(hdr_frame, text="Actualiser", command=self._charger_tableau).grid(
            row=0, column=4, padx=8
        )

        # Table + actions
        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        cols = ("Rubrique", "2025", "2024")
        self.tree = ttk.Treeview(body, columns=cols, show="headings")
        for col in cols:
            anchor = tk.E if col != "Rubrique" else tk.W
            self.tree.heading(col, text=col)
            self.tree.column(
                col, anchor=anchor, width=200 if col != "Rubrique" else 380
            )

        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        action_frame = ttk.Frame(self)
        action_frame.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Button(
            action_frame, text="Ouvrir mapping (CSV)", command=self.ouvrir_mapping
        ).pack(side=tk.LEFT)
        ttk.Button(
            action_frame, text="Ouvrir EtatFiFolder", command=self.ouvrir_folder
        ).pack(side=tk.LEFT, padx=6)
        ttk.Button(
            action_frame, text="Enregistrer entête", command=self.save_header_prefs
        ).pack(side=tk.LEFT, padx=6)
        ttk.Button(
            action_frame, text="Exporter Excel", command=self._export_excel
        ).pack(side=tk.RIGHT, padx=6)
        ttk.Button(action_frame, text="Exporter PDF", command=self._export_pdf).pack(
            side=tk.RIGHT, padx=6
        )

    def _fmt_or_dash(self, value):
        try:
            if abs(float(value)) < 0.005:
                return "-"
        except Exception:
            pass
        return format_montant(value)

    def _charger_tableau(self):
        # Lire mapping
        map_df = pd.read_csv(MAPPING_CSV, dtype=str)
        map_df = map_df.fillna("")

        # Convertir 'ordre' en numérique pour un tri correct (1, 2, 3... au lieu de 1, 10, 2...)
        map_df["ordre"] = pd.to_numeric(map_df["ordre"], errors="coerce")
        map_df = map_df.sort_values("ordre", na_position="last")

        # Préparer calculs par compte depuis le journal
        df = self.df.copy()

        # Convertir colonnes pour cohérence
        df["Année"] = df["Année"].astype(str)
        df["CompteDébit"] = (
            df["CompteDébit"]
            .fillna("")
            .apply(lambda x: str(int(float(x))) if x else "")
        )
        df["CompteCrédit"] = (
            df["CompteCrédit"]
            .fillna("")
            .apply(lambda x: str(int(float(x))) if x else "")
        )

        selection = self.annee_var.get() if hasattr(self, "annee_var") else "2025"
        if selection == "Both":
            current_year = "2025"
            previous_year = "2024"
        else:
            current_year = selection
            previous_year = str(int(selection) - 1)

        # Update tree columns dynamically
        self.tree["columns"] = ("Rubrique", current_year, previous_year)
        self.tree.heading("Rubrique", text="Rubrique")
        self.tree.heading(current_year, text=current_year)
        self.tree.heading(previous_year, text=previous_year)
        self.tree.column("Rubrique", anchor=tk.W, width=380)
        self.tree.column(current_year, anchor=tk.E, width=200)
        self.tree.column(previous_year, anchor=tk.E, width=200)
        debit_group = (
            df.groupby(["Année", "CompteDébit"])["MontantDébit"].sum().fillna(0)
        )
        credit_group = (
            df.groupby(["Année", "CompteCrédit"])["MontantCrédit"].sum().fillna(0)
        )

        def sum_for_prefixes(prefixes, annee, type_r):
            """
            Calcule la somme pour des comptes correspondant aux préfixes.
            Peut étre appelée dans les formules manuelles ci-dessous.
            """
            if not prefixes.strip():
                return 0.0
            total = 0.0
            for pref in [p.strip() for p in prefixes.split(";") if p.strip()]:
                # Cherche tous les comptes du journal qui commencent par ce préfixe
                for compte in debit_group.index.get_level_values(
                    "CompteDébit"
                ).unique():
                    if str(compte).startswith(
                        pref
                    ) and annee in debit_group.index.get_level_values("Année"):
                        try:
                            deb = float(debit_group.get((annee, compte), 0))
                            cred = float(credit_group.get((annee, compte), 0))
                            # Appliquer règle: produits = Crédit - Débit ; charges = Débit - Crédit
                            if type_r == "produit":
                                total += cred - deb
                            elif type_r == "charge":
                                total += deb - cred
                            else:
                                total += cred - deb
                        except:
                            pass
            return total

        def get_compte_solde(compte, annee):
            """Récupère le solde (Débit - Crédit) d'un compte pour une année"""
            try:
                deb = float(debit_group.get((annee, compte), 0))
                cred = float(credit_group.get((annee, compte), 0))
                return deb - cred
            except:
                return 0.0

        # Construire ledger selon mapping ordre
        rows = []
        values = {}  # Pour stocker les valeurs des rubriques pour les subtotaux
        for _, row in map_df.iterrows():
            ordre = row.get("ordre", "")
            rubrique = row.get("rubrique", "")
            prefixes = row.get("prefixes", "")
            type_r = row.get("type", "")

            # ===== FORMULES MANUELLES =====
            # C'est ici que vous devez entrer les formules de calcul pour chaque rubrique
            # Format:
            # if rubrique == "Nom de la rubrique":
            #     val2025 = get_compte_solde('701', '2025') + get_compte_solde('702', '2025') + ...
            #     val2024 = get_compte_solde('701', '2024') + get_compte_solde('702', '2024') + ...
            #
            # EXEMPLE - Décommentez et adaptez selon vos besoins:

            if rubrique == "Chiffres d'affaires":
                val_current = (
                    -get_compte_solde("70", current_year)
                    - get_compte_solde("701", current_year)
                    - get_compte_solde("702", current_year)
                    - get_compte_solde("703", current_year)
                    - get_compte_solde("705", current_year)
                    - get_compte_solde("706", current_year)
                    - get_compte_solde("707", current_year)
                    - get_compte_solde("708", current_year)
                    - get_compte_solde("7082", current_year)
                    - get_compte_solde("7083", current_year)
                    - get_compte_solde("7085", current_year)
                    - get_compte_solde("7086", current_year)
                    - get_compte_solde("7088", current_year)
                )
                val_previous = (
                    -get_compte_solde("70", previous_year)
                    - get_compte_solde("701", previous_year)
                    - get_compte_solde("702", previous_year)
                    - get_compte_solde("703", previous_year)
                    - get_compte_solde("705", previous_year)
                    - get_compte_solde("706", previous_year)
                    - get_compte_solde("707", previous_year)
                    - get_compte_solde("708", previous_year)
                    - get_compte_solde("7082", previous_year)
                    - get_compte_solde("7083", previous_year)
                    - get_compte_solde("7085", previous_year)
                    - get_compte_solde("7086", previous_year)
                    - get_compte_solde("7088", previous_year)
                )
            elif rubrique == "Production Stockée":
                val_current = (
                    -get_compte_solde("71", current_year)
                    - get_compte_solde("711", current_year)
                    - get_compte_solde("713", current_year)
                    - get_compte_solde("714", current_year)
                )
                val_previous = (
                    -get_compte_solde("71", previous_year)
                    - get_compte_solde("711", previous_year)
                    - get_compte_solde("713", previous_year)
                    - get_compte_solde("714", previous_year)
                )
            elif rubrique == "Production immobilisée":
                val_current = (
                    -get_compte_solde("72", current_year)
                    - get_compte_solde("721", current_year)
                    - get_compte_solde("722", current_year)
                )
                val_previous = (
                    -get_compte_solde("72", previous_year)
                    - get_compte_solde("721", previous_year)
                    - get_compte_solde("722", previous_year)
                )
            elif rubrique == "Achat consommés":
                val_current = (
                    get_compte_solde("60", current_year)
                    + get_compte_solde("601", current_year)
                    + get_compte_solde("602", current_year)
                    + get_compte_solde("6022", current_year)
                    + get_compte_solde("6023", current_year)
                    + get_compte_solde("60231", current_year)
                    + get_compte_solde("60232", current_year)
                    + get_compte_solde("60237", current_year)
                    + get_compte_solde("60221", current_year)
                    + get_compte_solde("60222", current_year)
                    + get_compte_solde("60223", current_year)
                    + get_compte_solde("60225", current_year)
                    + get_compte_solde("603", current_year)
                    + get_compte_solde("6031", current_year)
                    + get_compte_solde("6032", current_year)
                    + get_compte_solde("6037", current_year)
                    + get_compte_solde("604", current_year)
                    + get_compte_solde("605", current_year)
                    + get_compte_solde("606", current_year)
                    + get_compte_solde("6061", current_year)
                    + get_compte_solde("6062", current_year)
                    + get_compte_solde("6063", current_year)
                    + get_compte_solde("6064", current_year)
                    + get_compte_solde("6068", current_year)
                    + get_compte_solde("607", current_year)
                    + get_compte_solde("608", current_year)
                )
                val_previous = (
                    get_compte_solde("60", previous_year)
                    + get_compte_solde("601", previous_year)
                    + get_compte_solde("602", previous_year)
                    + get_compte_solde("6022", previous_year)
                    + get_compte_solde("6023", previous_year)
                    + get_compte_solde("60231", previous_year)
                    + get_compte_solde("60232", previous_year)
                    + get_compte_solde("60237", previous_year)
                    + get_compte_solde("60221", previous_year)
                    + get_compte_solde("60222", previous_year)
                    + get_compte_solde("60223", previous_year)
                    + get_compte_solde("60225", previous_year)
                    + get_compte_solde("603", previous_year)
                    + get_compte_solde("6031", previous_year)
                    + get_compte_solde("6032", previous_year)
                    + get_compte_solde("6037", previous_year)
                    + get_compte_solde("604", previous_year)
                    + get_compte_solde("605", previous_year)
                    + get_compte_solde("606", previous_year)
                    + get_compte_solde("6061", previous_year)
                    + get_compte_solde("6062", previous_year)
                    + get_compte_solde("6063", previous_year)
                    + get_compte_solde("6064", previous_year)
                    + get_compte_solde("6068", previous_year)
                    + get_compte_solde("607", previous_year)
                    + get_compte_solde("608", previous_year)
                )
            elif rubrique == "Services extérieurs et autres consommations":
                val_current = (
                    get_compte_solde("61", current_year)
                    + get_compte_solde("611", current_year)
                    + get_compte_solde("612", current_year)
                    + get_compte_solde("613", current_year)
                    + get_compte_solde("6132", current_year)
                    + get_compte_solde("6135", current_year)
                    + get_compte_solde("6136", current_year)
                    + get_compte_solde("614", current_year)
                    + get_compte_solde("615", current_year)
                    + get_compte_solde("616", current_year)
                    + get_compte_solde("617", current_year)
                    + get_compte_solde("618", current_year)
                    + get_compte_solde("62", current_year)
                    + get_compte_solde("621", current_year)
                    + get_compte_solde("622", current_year)
                    + get_compte_solde("623", current_year)
                    + get_compte_solde("624", current_year)
                    + get_compte_solde("6241", current_year)
                    + get_compte_solde("6242", current_year)
                    + get_compte_solde("625", current_year)
                    + get_compte_solde("626", current_year)
                    + get_compte_solde("627", current_year)
                    + get_compte_solde("628", current_year)
                )
                val_previous = (
                    get_compte_solde("61", previous_year)
                    + get_compte_solde("611", previous_year)
                    + get_compte_solde("612", previous_year)
                    + get_compte_solde("613", previous_year)
                    + get_compte_solde("6132", previous_year)
                    + get_compte_solde("6135", previous_year)
                    + get_compte_solde("6136", previous_year)
                    + get_compte_solde("614", previous_year)
                    + get_compte_solde("615", previous_year)
                    + get_compte_solde("616", previous_year)
                    + get_compte_solde("617", previous_year)
                    + get_compte_solde("618", previous_year)
                    + get_compte_solde("62", previous_year)
                    + get_compte_solde("621", previous_year)
                    + get_compte_solde("622", previous_year)
                    + get_compte_solde("623", previous_year)
                    + get_compte_solde("624", previous_year)
                    + get_compte_solde("6241", previous_year)
                    + get_compte_solde("6242", previous_year)
                    + get_compte_solde("625", previous_year)
                    + get_compte_solde("626", previous_year)
                    + get_compte_solde("627", previous_year)
                    + get_compte_solde("628", previous_year)
                )
            elif rubrique == "Charges de personnel":
                val_current = (
                    get_compte_solde("64", current_year)
                    + get_compte_solde("641", current_year)
                    + +get_compte_solde("644", current_year)
                    + get_compte_solde("644", current_year)
                    + get_compte_solde("645", current_year)
                    + get_compte_solde("646", current_year)
                    + get_compte_solde("647", current_year)
                    + get_compte_solde("648", current_year)
                )
                val_previous = (
                    get_compte_solde("64", previous_year)
                    + get_compte_solde("641", previous_year)
                    + +get_compte_solde("644", previous_year)
                    + get_compte_solde("644", previous_year)
                    + get_compte_solde("645", previous_year)
                    + get_compte_solde("646", previous_year)
                    + get_compte_solde("647", previous_year)
                    + get_compte_solde("648", previous_year)
                )
            elif rubrique == "Impôts, taxes et versements assimilés":
                val_current = (
                    get_compte_solde("63", current_year)
                    + get_compte_solde("631", current_year)
                    + get_compte_solde("635", current_year)
                )
                val_previous = (
                    get_compte_solde("63", previous_year)
                    + get_compte_solde("631", previous_year)
                    + get_compte_solde("635", previous_year)
                )
            elif rubrique == "Autres produits opérationnels":
                val_current = (
                    -get_compte_solde("74", current_year)
                    - get_compte_solde("741", current_year)
                    - get_compte_solde("748", current_year)
                    - get_compte_solde("75", current_year)
                    - get_compte_solde("751", current_year)
                    - get_compte_solde("752", current_year)
                    - get_compte_solde("753", current_year)
                    - get_compte_solde("754", current_year)
                    - get_compte_solde("755", current_year)
                    - get_compte_solde("756", current_year)
                    - get_compte_solde("757", current_year)
                    - get_compte_solde("758", current_year)
                )
                val_previous = (
                    -get_compte_solde("74", previous_year)
                    - get_compte_solde("741", previous_year)
                    - get_compte_solde("748", previous_year)
                    - get_compte_solde("75", previous_year)
                    - get_compte_solde("751", previous_year)
                    - get_compte_solde("752", previous_year)
                    - get_compte_solde("753", previous_year)
                    - get_compte_solde("754", previous_year)
                    - get_compte_solde("755", previous_year)
                    - get_compte_solde("756", previous_year)
                    - get_compte_solde("757", previous_year)
                    - get_compte_solde("758", previous_year)
                )
            elif rubrique == "Autres charges opérationnelles":
                val_current = (
                    get_compte_solde("65", current_year)
                    + get_compte_solde("651", current_year)
                    + get_compte_solde("652", current_year)
                    + get_compte_solde("653", current_year)
                    + get_compte_solde("654", current_year)
                    + get_compte_solde("655", current_year)
                    + get_compte_solde("656", current_year)
                    + get_compte_solde("657", current_year)
                    + get_compte_solde("658", current_year)
                )
                val_previous = (
                    get_compte_solde("65", previous_year)
                    + get_compte_solde("651", previous_year)
                    + get_compte_solde("652", previous_year)
                    + get_compte_solde("653", previous_year)
                    + get_compte_solde("654", previous_year)
                    + get_compte_solde("655", previous_year)
                    + get_compte_solde("656", previous_year)
                    + get_compte_solde("657", previous_year)
                    + get_compte_solde("658", previous_year)
                )
            elif rubrique == "Dotations aux amortissements":
                val_current = (
                    get_compte_solde("68", current_year)
                    + get_compte_solde("681", current_year)
                    + get_compte_solde("685", current_year)
                )
                val_previous = (
                    get_compte_solde("68", previous_year)
                    + get_compte_solde("681", previous_year)
                    + get_compte_solde("685", previous_year)
                )
            elif rubrique == "Produits financiers":
                val_current = (
                    -get_compte_solde("76", current_year)
                    - get_compte_solde("761", current_year)
                    - get_compte_solde("762", current_year)
                    - get_compte_solde("763", current_year)
                    - get_compte_solde("764", current_year)
                    - get_compte_solde("766", current_year)
                    - get_compte_solde("767", current_year)
                    - get_compte_solde("768", current_year)
                )
                val_previous = (
                    -get_compte_solde("76", previous_year)
                    - get_compte_solde("761", previous_year)
                    - get_compte_solde("762", previous_year)
                    - get_compte_solde("763", previous_year)
                    - get_compte_solde("764", previous_year)
                    - get_compte_solde("766", previous_year)
                    - get_compte_solde("767", previous_year)
                    - get_compte_solde("768", previous_year)
                )
            elif rubrique == "Charges financières":
                val_current = (
                    get_compte_solde("66", current_year)
                    + get_compte_solde("661", current_year)
                    + get_compte_solde("664", current_year)
                    + get_compte_solde("665", current_year)
                    + get_compte_solde("666", current_year)
                    + get_compte_solde("667", current_year)
                    + get_compte_solde("668", current_year)
                )
                val_previous = (
                    get_compte_solde("66", previous_year)
                    + get_compte_solde("661", previous_year)
                    + get_compte_solde("664", previous_year)
                    + get_compte_solde("665", previous_year)
                    + get_compte_solde("666", previous_year)
                    + get_compte_solde("667", previous_year)
                    + get_compte_solde("668", previous_year)
                )
            elif rubrique == "Eléments extraordinaires(produits)":
                val_current = -get_compte_solde("77", current_year)
                val_previous = -get_compte_solde("77", previous_year)
            elif rubrique == "Eléments extraordinaires(charges)":
                val_current = get_compte_solde("67", current_year)
                val_previous = get_compte_solde("67", previous_year)
            elif rubrique == "Impôts exigibles sur résultats":
                val_current = (
                    get_compte_solde("69", current_year)
                    + get_compte_solde("695", current_year)
                    + get_compte_solde("698", current_year)
                )
                val_previous = (
                    get_compte_solde("69", previous_year)
                    + get_compte_solde("695", previous_year)
                    + get_compte_solde("698", previous_year)
                )
            elif rubrique == "Impôts différés (variations)":
                val_current = get_compte_solde("692", current_year) + get_compte_solde(
                    "693", current_year
                )
                val_previous = get_compte_solde(
                    "692", previous_year
                ) + get_compte_solde("693", previous_year)
            # Subtotaux - calculés à partir d'autres rubriques
            elif rubrique == "I- PRODUCTION DE L'EXERCICE":
                val_current = (
                    values["Chiffres d'affaires"][0]
                    + values["Production Stockée"][0]
                    + values["Production immobilisée"][0]
                )
                val_previous = (
                    values["Chiffres d'affaires"][1]
                    + values["Production Stockée"][1]
                    + values["Production immobilisée"][1]
                )
            elif rubrique == "II- CONSOMMATION DE L'EXERCICE":
                val_current = (
                    values["Achat consommés"][0]
                    + values["Services extérieurs et autres consommations"][0]
                )
                val_previous = (
                    values["Achat consommés"][1]
                    + values["Services extérieurs et autres consommations"][1]
                )
            elif rubrique == "III- VALEUR AJOUTEE DE L'EXPLOITATION":
                val_current = (
                    values["I- PRODUCTION DE L'EXERCICE"][0]
                    - values["II- CONSOMMATION DE L'EXERCICE"][0]
                )
                val_previous = (
                    values["I- PRODUCTION DE L'EXERCICE"][1]
                    - values["II- CONSOMMATION DE L'EXERCICE"][1]
                )
            elif rubrique == "IV- EXCEDENT BRUTE D'EXPLOITATION":
                val_current = (
                    values["III- VALEUR AJOUTEE DE L'EXPLOITATION"][0]
                    - values["Charges de personnel"][0]
                    - values["Impôts, taxes et versements assimilés"][0]
                )
                val_previous = (
                    values["III- VALEUR AJOUTEE DE L'EXPLOITATION"][1]
                    - values["Charges de personnel"][1]
                    - values["Impôts, taxes et versements assimilés"][1]
                )
            elif rubrique == "V- RESULTAT OPERATIONNEL":
                val_current = (
                    values["Autres produits opérationnels"][0]
                    - values["Autres charges opérationnelles"][0]
                )
                val_previous = (
                    values["Autres produits opérationnels"][1]
                    - values["Autres charges opérationnelles"][1]
                )
            elif rubrique == "VI- RESULTAT FINANCIER":
                val_current = (
                    values["Produits financiers"][0] - values["Charges financières"][0]
                )
                val_previous = (
                    values["Produits financiers"][1] - values["Charges financières"][1]
                )
            elif rubrique == "VII- RESULTAT AVANT IMPOT":
                val_current = (
                    values["IV- EXCEDENT BRUTE D'EXPLOITATION"][0]
                    - values["Dotations aux amortissements"][0]
                    + values["V- RESULTAT OPERATIONNEL"][0]
                    + values["VI- RESULTAT FINANCIER"][0]
                )
                val_previous = (
                    values["IV- EXCEDENT BRUTE D'EXPLOITATION"][1]
                    - values["Dotations aux amortissements"][1]
                    + values["V- RESULTAT OPERATIONNEL"][1]
                    + values["VI- RESULTAT FINANCIER"][1]
                )
            elif rubrique == "IX- RESULTAT EXTRAORDINAIRE":
                val_current = (
                    values["Eléments extraordinaires(produits)"][0]
                    - values["Eléments extraordinaires(charges)"][0]
                )
                val_previous = (
                    values["Eléments extraordinaires(produits)"][1]
                    - values["Eléments extraordinaires(charges)"][1]
                )
            elif rubrique == "X- RESULTAT NET DE L'EXERCICE":
                val_current = (
                    values["VII- RESULTAT AVANT IMPOT"][0]
                    - values["Impôts exigibles sur résultats"][0]
                    + values["Impôts différés (variations)"][0]
                    + values["IX- RESULTAT EXTRAORDINAIRE"][0]
                )
                val_previous = (
                    values["VII- RESULTAT AVANT IMPOT"][1]
                    - values["Impôts exigibles sur résultats"][1]
                    + values["Impôts différés (variations)"][1]
                    + values["IX- RESULTAT EXTRAORDINAIRE"][1]
                )
            else:
                # Par défaut, utilise les préfixes du mapping
                val_current = sum_for_prefixes(prefixes, current_year, type_r)
                val_previous = sum_for_prefixes(prefixes, previous_year, type_r)
            # ===== FIN FORMULES MANUELLES =====

            # Stocker les valeurs pour les subtotaux
            values[rubrique] = (val_current, val_previous)

            rows.append((rubrique, val_current, val_previous, type_r))

        # Nettoyer tree
        for it in self.tree.get_children():
            self.tree.delete(it)

        for r in rows:
            self.tree.insert(
                "",
                tk.END,
                values=(r[0], self._fmt_or_dash(r[1]), self._fmt_or_dash(r[2])),
            )

        # sauvegarder tableau courant (utilisé uniquement pour affichage)
        # l'attribut `current_df` n'est plus nécessaire pour les exports supprimés

    def ouvrir_mapping(self):
        # ouvre l'explorateur pour éditer le CSV
        try:
            os.startfile(MAPPING_CSV)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir mapping: {e}")

    def ouvrir_folder(self):
        os.makedirs(STATE_FOLDER, exist_ok=True)
        try:
            os.startfile(STATE_FOLDER)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir dossier: {e}")

    def _load_settings(self):
        self.header_settings = load_header_settings()
        self.header_text = format_header_text(self.header_settings)
        if self.header_settings.get("nom_societe"):
            self.titre_var.set(self.header_settings.get("nom_societe"))
        if hasattr(self, "header_info_var"):
            self.header_info_var.set(self.header_text)

    def save_header_prefs(self):
        try:
            prefs = load_header_settings()
            prefs["nom_societe"] = self.titre_var.get()
            save_header_settings(prefs)
            self._load_settings()
            messagebox.showinfo("Enregistré", "Préférences d'en-tête enregistrées")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'enregistrer: {e}")

    def _export_excel(self):
        export_treeview_to_excel(self.tree, "compte_resultat_nature.xlsx", self)

    def _export_pdf(self):
        export_treeview_to_pdf(
            self.tree,
            "Compte de résultat par nature",
            "compte_resultat_nature.pdf",
            self,
        )
