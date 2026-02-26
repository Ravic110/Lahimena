"""
Module fenêtre : Compte de résultat par fonction
Affiche le tableau selon la matrice Excel "compte de résultat par fonction.xlsx".
"""

import os
import tkinter as tk
from tkinter import ttk

import pandas as pd

from finances.tsarakonta.config import CONFIG
from finances.tsarakonta.models.data import DataManager
from finances.tsarakonta.utils.exports import export_treeview_to_excel, export_treeview_to_pdf

from .settings import format_header_text, load_header_settings

STATE_FOLDER = os.path.join(os.path.dirname(os.path.dirname(__file__)), "EtatFiFolder")
MAPPING_CSV = os.path.join(STATE_FOLDER, "mapping_resultat_fonction.csv")
SOURCE_XLSX = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "compte de résultat par fonction.xlsx"
)


class CompteResultatFonctionWindow(tk.Toplevel):
    """Fenêtre du compte de résultat par fonction"""

    def __init__(self, parent, df=None):
        super().__init__(parent)
        self.parent = parent

        livre_file = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), "LivreCompta.xlsx"
        )
        if os.path.exists(livre_file):
            self.df = DataManager.charger_feuille(livre_file, CONFIG["feuille_journal"])
            if self.df is None:
                self.df = pd.DataFrame(columns=CONFIG["colonnes_journal"])
        else:
            self.df = (
                df
                if df is not None
                else pd.DataFrame(columns=CONFIG["colonnes_journal"])
            )

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

        os.makedirs(STATE_FOLDER, exist_ok=True)
        self._ensure_mapping(force=os.path.exists(SOURCE_XLSX))

        self.title("Compte de résultat par fonction")
        self.geometry("980x640")
        self.formules = self._initialiser_formules()
        self._preparer_journal()

        self._build_ui()
        self._charger_tableau()

    def _initialiser_formules(self):
        return {
            "Produits des activités ordinaires": self._fx_produits_activites_ordinaires,
            "Produits des ativités ordinaires": self._fx_produits_activites_ordinaires,
            "Coût des ventes": self._fx_cout_des_ventes,
            "MARGE BRUTE": self._fx_marge_brute,
            "Autres produits opérationnels": self._fx_autres_produits_operationnels,
            "Coût commerciaux": self._fx_cout_commerciaux,
            "Charges administratives": self._fx_charges_administratives,
            "Autres charges opérationnelles": self._fx_autres_charges_operationnelles,
            "RESULTAT OPERATIONNEL": self._fx_resultat_operationnel,
            "Résultat opérationnel": self._fx_resultat_operationnel,
            "Produit financiers": self._fx_produits_financiers,
            "Produits financiers": self._fx_produits_financiers,
            "Charges financiers": self._fx_charges_financieres,
            "Charges financières": self._fx_charges_financieres,
            "RESULTAT AVANT IMPOT": self._fx_resultat_avant_impot,
            "Résultat avant impôt": self._fx_resultat_avant_impot,
            "Impôt sur le résultat": self._fx_impot_sur_resultat,
            "Impôts sur le résultat": self._fx_impot_sur_resultat,
            "Impôt différés": self._fx_impot_differes,
            "Impôts différés": self._fx_impot_differes,
            "RESULTAT NET DES ACTIVITES ORDINAIRES": self._fx_resultat_net_activites_ordinaires,
            "Résultat net des activités ordinaires": self._fx_resultat_net_activites_ordinaires,
            "Charges extraordinaires": self._fx_charges_extraordinaires,
            "Produits extraordinaires": self._fx_produits_extraordinaires,
            "RESULTAT NET DE L'EXERCICE": self._fx_resultat_net_exercice,
            "Résultat net de l'exercice": self._fx_resultat_net_exercice,
        }

    def _to_float(self, value):
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        txt = str(value).strip().replace(" ", "").replace(",", ".")
        if not txt:
            return 0.0
        try:
            return float(txt)
        except ValueError:
            return 0.0

    def _normaliser_compte(self, value):
        if pd.isna(value):
            return ""
        txt = str(value).strip()
        if not txt:
            return ""
        try:
            if txt.endswith(".0"):
                return str(int(float(txt)))
        except ValueError:
            pass
        return txt

    def _preparer_journal(self):
        df = self.df.copy()
        for col in [
            "Année",
            "CompteDébit",
            "CompteCrédit",
            "MontantDébit",
            "MontantCrédit",
        ]:
            if col not in df.columns:
                df[col] = "" if col in ["Année", "CompteDébit", "CompteCrédit"] else 0.0

        df["Année"] = df["Année"].fillna("").astype(str).str.strip()
        df["CompteDébit"] = df["CompteDébit"].apply(self._normaliser_compte)
        df["CompteCrédit"] = df["CompteCrédit"].apply(self._normaliser_compte)
        df["MontantDébit"] = df["MontantDébit"].apply(self._to_float)
        df["MontantCrédit"] = df["MontantCrédit"].apply(self._to_float)

        self.debit_group = df.groupby(["Année", "CompteDébit"])["MontantDébit"].sum()
        self.credit_group = df.groupby(["Année", "CompteCrédit"])["MontantCrédit"].sum()

    def _compte_solde(self, compte, annee):
        debit = float(self.debit_group.get((annee, compte), 0.0))
        credit = float(self.credit_group.get((annee, compte), 0.0))
        return debit - credit

    def _sum_soldes(self, comptes, annee, negate=False):
        total = sum(self._compte_solde(c, annee) for c in comptes)
        return -total if negate else total

    def _fmt_fr(self, value):
        return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")

    def _fmt_or_dash(self, value):
        return "-" if abs(value) < 0.005 else self._fmt_fr(value)

    def _fmt_display_or_dash(self, value):
        txt = str(value).strip()
        if txt in ("", "-", "—"):
            return "-"
        cleaned = txt.replace("\u00a0", "").replace(" ", "").replace(",", ".")
        try:
            if abs(float(cleaned)) < 0.005:
                return "-"
        except Exception:
            pass
        return txt

    def _annees_selectionnees(self):
        annee_courante = self.annee_var.get() if hasattr(self, "annee_var") else "2025"
        try:
            annee_precedente = str(int(annee_courante) - 1)
        except ValueError:
            annee_courante = "2025"
            annee_precedente = "2024"
        return annee_courante, annee_precedente

    def _fx_produits_activites_ordinaires(self):
        annee_courante, annee_precedente = self._annees_selectionnees()

        comptes_chiffres_affaires = [
            "70",
            "701",
            "702",
            "703",
            "705",
            "706",
            "707",
            "708",
            "7082",
            "7083",
            "7085",
            "7086",
            "7088",
        ]
        comptes_production_stockee = ["71", "711", "713", "714"]
        comptes_production_immobilisee = ["72", "721", "722"]

        val_n = (
            self._sum_soldes(comptes_chiffres_affaires, annee_courante, negate=True)
            + self._sum_soldes(comptes_production_stockee, annee_courante, negate=True)
            + self._sum_soldes(
                comptes_production_immobilisee, annee_courante, negate=True
            )
        )
        val_n1 = (
            self._sum_soldes(comptes_chiffres_affaires, annee_precedente, negate=True)
            + self._sum_soldes(
                comptes_production_stockee, annee_precedente, negate=True
            )
            + self._sum_soldes(
                comptes_production_immobilisee, annee_precedente, negate=True
            )
        )

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_cout_des_ventes(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes_services_exterieurs = [
            "61",
            "611",
            "612",
            "613",
            "6132",
            "6135",
            "6136",
            "614",
            "615",
            "616",
            "617",
            "618",
            "62",
            "621",
            "622",
            "623",
            "624",
            "6241",
            "6242",
            "625",
            "626",
            "627",
            "628",
        ]

        val_n = self._sum_soldes(comptes_services_exterieurs, annee_courante)
        val_n1 = self._sum_soldes(comptes_services_exterieurs, annee_precedente)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_marge_brute(self):
        produits_n, produits_n1 = self._fx_produits_activites_ordinaires()
        couts_n, couts_n1 = self._fx_cout_des_ventes()

        val_n = self._to_float(produits_n) - self._to_float(couts_n)
        val_n1 = self._to_float(produits_n1) - self._to_float(couts_n1)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_autres_produits_operationnels(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes_autres_produits = [
            "74",
            "741",
            "748",
            "75",
            "751",
            "752",
            "753",
            "754",
            "755",
            "756",
            "757",
            "758",
        ]

        val_n = self._sum_soldes(comptes_autres_produits, annee_courante, negate=True)
        val_n1 = self._sum_soldes(
            comptes_autres_produits, annee_precedente, negate=True
        )

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_cout_commerciaux(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes_achat_consommes = [
            "60",
            "601",
            "602",
            "6022",
            "6023",
            "60231",
            "60232",
            "60237",
            "60221",
            "60222",
            "60223",
            "60225",
            "603",
            "6031",
            "6032",
            "6037",
            "604",
            "605",
            "606",
            "6061",
            "6062",
            "6063",
            "6064",
            "6068",
            "607",
            "608",
        ]

        val_n = self._sum_soldes(comptes_achat_consommes, annee_courante)
        val_n1 = self._sum_soldes(comptes_achat_consommes, annee_precedente)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_reprises_provisions_pertes_valeurs(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes_reprises = [
            "78",
            "781",
            "7811",
            "7815",
            "7816",
            "7817",
            "7818",
            "785",
            "786",
            "787",
        ]

        val_n = self._sum_soldes(comptes_reprises, annee_courante, negate=True)
        val_n1 = self._sum_soldes(comptes_reprises, annee_precedente, negate=True)

        return val_n, val_n1

    def _fx_charges_administratives(self):
        annee_courante, annee_precedente = self._annees_selectionnees()

        comptes_charges_personnel = ["64", "641", "644", "645", "646", "647", "648"]
        comptes_impots_taxes = ["63", "631", "635"]
        comptes_dotations_amort = ["68", "681", "685"]

        reprises_n, reprises_n1 = self._fx_reprises_provisions_pertes_valeurs()

        val_n = (
            self._sum_soldes(comptes_charges_personnel, annee_courante)
            + self._sum_soldes(comptes_impots_taxes, annee_courante)
            + self._sum_soldes(comptes_dotations_amort, annee_courante)
            + reprises_n
        )
        val_n1 = (
            self._sum_soldes(comptes_charges_personnel, annee_precedente)
            + self._sum_soldes(comptes_impots_taxes, annee_precedente)
            + self._sum_soldes(comptes_dotations_amort, annee_precedente)
            + reprises_n1
        )

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_autres_charges_operationnelles(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["65", "651", "652", "653", "654", "655", "656", "657", "658"]

        val_n = self._sum_soldes(comptes, annee_courante)
        val_n1 = self._sum_soldes(comptes, annee_precedente)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_resultat_operationnel(self):
        marge_n, marge_n1 = self._fx_marge_brute()
        produits_n, produits_n1 = self._fx_autres_produits_operationnels()
        couts_com_n, couts_com_n1 = self._fx_cout_commerciaux()
        charges_admin_n, charges_admin_n1 = self._fx_charges_administratives()
        autres_charges_n, autres_charges_n1 = self._fx_autres_charges_operationnelles()

        val_n = (
            self._to_float(marge_n)
            + self._to_float(produits_n)
            - self._to_float(couts_com_n)
            - self._to_float(charges_admin_n)
            - self._to_float(autres_charges_n)
        )
        val_n1 = (
            self._to_float(marge_n1)
            + self._to_float(produits_n1)
            - self._to_float(couts_com_n1)
            - self._to_float(charges_admin_n1)
            - self._to_float(autres_charges_n1)
        )

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_produits_financiers(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["76", "761", "762", "763", "764", "766", "767", "768"]

        val_n = self._sum_soldes(comptes, annee_courante, negate=True)
        val_n1 = self._sum_soldes(comptes, annee_precedente, negate=True)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_charges_financieres(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["66", "661", "664", "665", "666", "667", "668"]

        val_n = self._sum_soldes(comptes, annee_courante)
        val_n1 = self._sum_soldes(comptes, annee_precedente)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_resultat_avant_impot(self):
        resultat_op_n, resultat_op_n1 = self._fx_resultat_operationnel()
        produits_fin_n, produits_fin_n1 = self._fx_produits_financiers()
        charges_fin_n, charges_fin_n1 = self._fx_charges_financieres()

        val_n = (
            self._to_float(resultat_op_n)
            + self._to_float(produits_fin_n)
            - self._to_float(charges_fin_n)
        )
        val_n1 = (
            self._to_float(resultat_op_n1)
            + self._to_float(produits_fin_n1)
            - self._to_float(charges_fin_n1)
        )

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_impot_sur_resultat(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["69", "695", "698"]

        val_n = self._sum_soldes(comptes, annee_courante)
        val_n1 = self._sum_soldes(comptes, annee_precedente)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_impot_differes(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["692", "693"]

        val_n = self._sum_soldes(comptes, annee_courante)
        val_n1 = self._sum_soldes(comptes, annee_precedente)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_resultat_net_activites_ordinaires(self):
        resultat_avant_impot_n, resultat_avant_impot_n1 = (
            self._fx_resultat_avant_impot()
        )
        impot_resultat_n, impot_resultat_n1 = self._fx_impot_sur_resultat()
        impot_differes_n, impot_differes_n1 = self._fx_impot_differes()

        val_n = (
            self._to_float(resultat_avant_impot_n)
            - self._to_float(impot_resultat_n)
            + self._to_float(impot_differes_n)
        )
        val_n1 = (
            self._to_float(resultat_avant_impot_n1)
            - self._to_float(impot_resultat_n1)
            + self._to_float(impot_differes_n1)
        )

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_charges_extraordinaires(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["67"]

        val_n = self._sum_soldes(comptes, annee_courante)
        val_n1 = self._sum_soldes(comptes, annee_precedente)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_produits_extraordinaires(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["77"]

        val_n = self._sum_soldes(comptes, annee_courante, negate=True)
        val_n1 = self._sum_soldes(comptes, annee_precedente, negate=True)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_resultat_net_exercice(self):
        resultat_ord_n, resultat_ord_n1 = self._fx_resultat_net_activites_ordinaires()
        charges_extra_n, charges_extra_n1 = self._fx_charges_extraordinaires()
        produits_extra_n, produits_extra_n1 = self._fx_produits_extraordinaires()

        val_n = (
            self._to_float(resultat_ord_n)
            - self._to_float(charges_extra_n)
            + self._to_float(produits_extra_n)
        )
        val_n1 = (
            self._to_float(resultat_ord_n1)
            - self._to_float(charges_extra_n1)
            + self._to_float(produits_extra_n1)
        )

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _clean_excel_value(self, value):
        if pd.isna(value):
            return ""
        if isinstance(value, (int, float)):
            if float(value).is_integer():
                return str(int(value))
            return str(value)
        return str(value).strip()

    def _ensure_mapping(self, force=False):
        if os.path.exists(MAPPING_CSV) and not force:
            return

        rows = []
        if os.path.exists(SOURCE_XLSX):
            try:
                df_xlsx = pd.read_excel(SOURCE_XLSX)
                if df_xlsx.shape[1] >= 3:
                    subset = df_xlsx.iloc[:, :3].copy()
                    subset.columns = ["rubrique", "valeur_n", "valeur_n1"]
                    ordre = 1
                    for _, row in subset.iterrows():
                        rows.append(
                            (
                                ordre,
                                self._clean_excel_value(row["rubrique"]),
                                self._clean_excel_value(row["valeur_n"]),
                                self._clean_excel_value(row["valeur_n1"]),
                            )
                        )
                        ordre += 1
            except Exception:
                rows = []

        if not rows:
            rows = [
                (1, "Produits des activités ordinaires", "", ""),
                (2, "Coût des ventes", "", ""),
                (3, "", "", ""),
                (4, "MARGE BRUTE", "", ""),
                (5, "", "", ""),
                (6, "Autres produits opérationnels", "", ""),
                (7, "Coût commerciaux", "", ""),
                (8, "Charges administratives", "", ""),
                (9, "Autres charges opérationnelles", "", ""),
                (10, "", "", ""),
                (11, "RESULTAT OPERATIONNEL", "", ""),
                (12, "", "", ""),
                (13, "Produit financiers", "", ""),
                (14, "Charges financiers", "", ""),
                (15, "", "", ""),
                (16, "RESULTAT AVANT IMPOT", "", ""),
                (17, "", "", ""),
                (18, "Impôt sur le résultat", "", ""),
                (19, "Impôt différés", "", ""),
                (20, "", "", ""),
                (21, "RESULTAT NET DES ACTIVITES ORDINAIRES", "", ""),
                (22, "Charges extraordinaires", "", ""),
                (23, "Produits extraordinaires", "", ""),
                (24, "RESULTAT NET DE L'EXERCICE", "", ""),
            ]

        df_map = pd.DataFrame(
            rows, columns=["ordre", "rubrique", "valeur_n", "valeur_n1"]
        )
        df_map.to_csv(MAPPING_CSV, index=False, encoding="utf-8")

    def _build_ui(self):
        if self.header_text:
            header = ttk.LabelFrame(self, text="En-tête société", padding=8)
            header.pack(fill=tk.X, padx=8, pady=(8, 4))
            ttk.Label(header, text=self.header_text, justify=tk.LEFT).pack(anchor=tk.W)

        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Arrêté au:").pack(side=tk.LEFT, padx=(0, 6))
        self.annee_var = tk.StringVar(value="2025")
        self.annee_combo = ttk.Combobox(
            top,
            values=self.annees,
            textvariable=self.annee_var,
            state="readonly",
            width=10,
        )
        self.annee_combo.pack(side=tk.LEFT)
        self.annee_combo.bind("<<ComboboxSelected>>", lambda e: self._charger_tableau())

        ttk.Label(top, text="(Colonnes: N et N-1)").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(top, text="Actualiser", command=self._charger_tableau).pack(
            side=tk.RIGHT
        )

        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        cols = ("Période d'exercice", "Valeur N", "Valeur N-1")
        self.tree = ttk.Treeview(body, columns=cols, show="headings")

        for col in cols:
            anchor = tk.W if col == "Période d'exercice" else tk.E
            self.tree.heading(col, text=col)
            self.tree.column(
                col, anchor=anchor, width=520 if col == "Période d'exercice" else 190
            )

        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(body, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        action_frame = ttk.Frame(self)
        action_frame.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Button(
            action_frame, text="Exporter Excel", command=self._export_excel
        ).pack(side=tk.RIGHT, padx=6)
        ttk.Button(action_frame, text="Exporter PDF", command=self._export_pdf).pack(
            side=tk.RIGHT, padx=6
        )

    def _appliquer_formule(self, rubrique, valeurs_mapping):
        formule = self.formules.get(rubrique)
        if callable(formule):
            resultat = formule()
            if isinstance(resultat, tuple) and len(resultat) == 2:
                return resultat
        return valeurs_mapping

    def _charger_tableau(self):
        annee_courante = self.annee_var.get() if hasattr(self, "annee_var") else "2025"
        try:
            annee_precedente = str(int(annee_courante) - 1)
        except ValueError:
            annee_courante = "2025"
            annee_precedente = "2024"

        self.tree.heading("Valeur N", text=f"{annee_courante}")
        self.tree.heading("Valeur N-1", text=f"{annee_precedente}")

        for it in self.tree.get_children():
            self.tree.delete(it)

        if not os.path.exists(MAPPING_CSV):
            self._ensure_mapping()

        df_map = pd.read_csv(MAPPING_CSV, dtype=str).fillna("")
        if "ordre" in df_map.columns:
            ordre_num = pd.to_numeric(df_map["ordre"], errors="coerce")
            if ordre_num.notna().any():
                df_map = (
                    df_map.assign(_ordre_num=ordre_num)
                    .sort_values("_ordre_num", kind="stable")
                    .drop(columns=["_ordre_num"])
                )

        for _, row in df_map.iterrows():
            rubrique = row.get("rubrique", "")
            valeurs = self._appliquer_formule(
                rubrique, (row.get("valeur_n", ""), row.get("valeur_n1", ""))
            )
            self.tree.insert(
                "",
                tk.END,
                values=(
                    rubrique,
                    self._fmt_display_or_dash(valeurs[0]),
                    self._fmt_display_or_dash(valeurs[1]),
                ),
            )

    def _export_excel(self):
        export_treeview_to_excel(self.tree, "compte_resultat_fonction.xlsx", self)

    def _export_pdf(self):
        export_treeview_to_pdf(
            self.tree,
            "Compte de résultat par fonction",
            "compte_resultat_fonction.pdf",
            self,
        )
