"""
Module fenêtre : Bilan passif (PCG 2005 Madagascar)
Affiche le tableau de base selon la matrice de mapping, sans formules.
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
MAPPING_CSV = os.path.join(STATE_FOLDER, "mapping_bilan_passif.csv")
SOURCE_XLSX = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "Bilan passif.xlsx"
)


class BilanPassifWindow(tk.Toplevel):
    """Fenêtre du bilan passif"""

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

        self.title("Bilan passif")
        self.geometry("1100x650")
        self.formules = self._initialiser_formules()
        self._preparer_journal()

        self._build_ui()
        self._charger_tableau()

    def _initialiser_formules(self):
        return {
            "Fond propre": self._fx_fond_propre,
            "Primes et réserves": self._fx_primes_et_reserves,
            "Ecart d'évaluation": self._fx_ecart_evaluation,
            "Report à nouveau": self._fx_report_a_nouveau,
            "Résultat net": self._fx_resultat_net,
            "Produits différé - Subventions d'investissement": self._fx_produits_differes_subventions,
            "Produits différés - Subventions d'investissements": self._fx_produits_differes_subventions,
            "Emprunts et dettes financières +1ans": self._fx_emprunts_dettes_plus_un_an,
            "Emprunt et dettes financière plus de 1 ans": self._fx_emprunts_dettes_plus_un_an,
            "Provision et produits constatés d'avance": self._fx_provisions_produits_constates_avance,
            "Provisions et produits constatés d'avance": self._fx_provisions_produits_constates_avance,
            "Dettes court terme -1ans": self._fx_dettes_court_terme_moins_un_an,
            "Dettes court terme -1 ans": self._fx_dettes_court_terme_moins_un_an,
            "Fournisseurs et comptes rattachés": self._fx_fournisseurs_comptes_rattaches,
            "Provision et produits constatés d'avance (passifs courants)": self._fx_provisions_passifs_courants,
            "Provisions et produits constatés d'avance (passifs courants)": self._fx_provisions_passifs_courants,
            "povisions et produits constatés d'avance (passifs courants)": self._fx_provisions_passifs_courants,
            "Autres dettes": self._fx_autres_dettes,
            "Compte de trésorerie passifs (Découvert)": self._fx_tresorerie_passif_decouvert,
            "Compte de trésorerie passif (Découvert)": self._fx_tresorerie_passif_decouvert,
            "TOTAL CAPITAUX PROPORES": self._fx_total_capitaux_propres,
            "TOTAL CAPITAUX PROPRES": self._fx_total_capitaux_propres,
            "TOTAL PASSIFS NON COURANTS": self._fx_total_passifs_non_courants,
            "TOTAL PASSIFS COURANTS": self._fx_total_passifs_courants,
            "TOTAL DES PASSIFS": self._fx_total_des_passifs,
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
        comptes_debit = set(df["CompteDébit"].dropna().astype(str).tolist())
        comptes_credit = set(df["CompteCrédit"].dropna().astype(str).tolist())
        self.comptes = sorted((comptes_debit | comptes_credit) - {""})

    def _solde_passif_par_prefixes(self, annee, prefixes):
        total = 0.0
        for compte in self.comptes:
            if any(compte.startswith(p) for p in prefixes):
                debit = float(self.debit_group.get((annee, compte), 0.0))
                credit = float(self.credit_group.get((annee, compte), 0.0))
                total += credit - debit
        return total

    def _compte_solde(self, compte, annee):
        debit = float(self.debit_group.get((annee, compte), 0.0))
        credit = float(self.credit_group.get((annee, compte), 0.0))
        return debit - credit

    def _sum_soldes(self, comptes, annee, negate=False):
        total = sum(self._compte_solde(c, annee) for c in comptes)
        return -total if negate else total

    def _fmt_fr(self, value):
        txt = f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
        return txt

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

    def _fx_fond_propre(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes_fond_propre = ["101", "108"]

        solde_courant = self._solde_passif_par_prefixes(
            annee_courante, comptes_fond_propre
        )
        solde_precedent = self._solde_passif_par_prefixes(
            annee_precedente, comptes_fond_propre
        )

        return self._fmt_or_dash(solde_courant), self._fmt_or_dash(solde_precedent)

    def _fx_primes_et_reserves(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes_primes_reserves = ["104", "106"]

        solde_courant = self._solde_passif_par_prefixes(
            annee_courante, comptes_primes_reserves
        )
        solde_precedent = self._solde_passif_par_prefixes(
            annee_precedente, comptes_primes_reserves
        )

        return self._fmt_or_dash(solde_courant), self._fmt_or_dash(solde_precedent)

    def _fx_ecart_evaluation(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes_ecart = ["105"]

        solde_courant = self._solde_passif_par_prefixes(annee_courante, comptes_ecart)
        solde_precedent = self._solde_passif_par_prefixes(
            annee_precedente, comptes_ecart
        )

        return self._fmt_or_dash(solde_courant), self._fmt_or_dash(solde_precedent)

    def _fx_report_a_nouveau(self):
        annee_courante, annee_precedente = self._annees_selectionnees()

        compte_110_courant = self._solde_passif_par_prefixes(annee_courante, ["110"])
        compte_119_courant = self._solde_passif_par_prefixes(annee_courante, ["119"])
        valeur_courante = compte_110_courant - compte_119_courant

        compte_110_precedent = self._solde_passif_par_prefixes(
            annee_precedente, ["110"]
        )
        compte_119_precedent = self._solde_passif_par_prefixes(
            annee_precedente, ["119"]
        )
        valeur_precedente = compte_110_precedent - compte_119_precedent

        return self._fmt_or_dash(valeur_courante), self._fmt_or_dash(valeur_precedente)

    def _resultat_net_exercice_etat_resultat(self, annee):
        chiffres_affaires = self._sum_soldes(
            [
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
            ],
            annee,
            negate=True,
        )
        production_stockee = self._sum_soldes(
            ["71", "711", "713", "714"], annee, negate=True
        )
        production_immobilisee = self._sum_soldes(
            ["72", "721", "722"], annee, negate=True
        )

        achat_consommes = self._sum_soldes(
            [
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
            ],
            annee,
        )
        services_exterieurs = self._sum_soldes(
            [
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
            ],
            annee,
        )

        charges_personnel = self._sum_soldes(
            ["64", "641", "644", "644", "645", "646", "647", "648"], annee
        )
        impots_taxes = self._sum_soldes(["63", "631", "635"], annee)

        autres_produits_operationnels = self._sum_soldes(
            [
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
            ],
            annee,
            negate=True,
        )
        autres_charges_operationnelles = self._sum_soldes(
            ["65", "651", "652", "653", "654", "655", "656", "657", "658"], annee
        )

        dotations_amortissements = self._sum_soldes(["68", "681", "685"], annee)
        produits_financiers = self._sum_soldes(
            ["76", "761", "762", "763", "764", "766", "767", "768"], annee, negate=True
        )
        charges_financieres = self._sum_soldes(
            ["66", "661", "664", "665", "666", "667", "668"], annee
        )

        elements_extra_produits = self._sum_soldes(["77"], annee, negate=True)
        elements_extra_charges = self._sum_soldes(["67"], annee)

        impots_exigibles_resultats = self._sum_soldes(["69", "695", "698"], annee)
        impots_differes_variations = self._sum_soldes(["692", "693"], annee)

        production_exercice = (
            chiffres_affaires + production_stockee + production_immobilisee
        )
        consommation_exercice = achat_consommes + services_exterieurs
        valeur_ajoutee = production_exercice - consommation_exercice
        excedent_brut_exploitation = valeur_ajoutee - charges_personnel - impots_taxes
        resultat_operationnel = (
            autres_produits_operationnels - autres_charges_operationnelles
        )
        resultat_financier = produits_financiers - charges_financieres
        resultat_avant_impot = (
            excedent_brut_exploitation
            - dotations_amortissements
            + resultat_operationnel
            + resultat_financier
        )
        resultat_extraordinaire = elements_extra_produits - elements_extra_charges

        return (
            resultat_avant_impot
            - impots_exigibles_resultats
            + impots_differes_variations
            + resultat_extraordinaire
        )

    def _fx_resultat_net(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        valeur_courante = self._resultat_net_exercice_etat_resultat(annee_courante)
        valeur_precedente = self._resultat_net_exercice_etat_resultat(annee_precedente)
        return self._fmt_or_dash(valeur_courante), self._fmt_or_dash(valeur_precedente)

    def _valeurs_rubrique_numeriques(self, rubrique):
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
            if row.get("rubrique", "") == rubrique:
                valeurs = self._appliquer_formule(
                    rubrique, (row.get("valeur_n", ""), row.get("valeur_n1", ""))
                )
                return self._to_float(valeurs[0]), self._to_float(valeurs[1])

        return 0.0, 0.0

    def _fx_total_capitaux_propres(self):
        rubriques = [
            "Fond propre",
            "Primes et réserves",
            "Ecart d'évaluation",
            "Ecart d'équivalence",
            "Report à nouveau",
            "Résultat net",
            "Part de la société consolidant",
            "Part des minoritaires",
        ]

        total_n = 0.0
        total_n1 = 0.0

        for rubrique in rubriques:
            val_n, val_n1 = self._valeurs_rubrique_numeriques(rubrique)
            total_n += val_n
            total_n1 += val_n1

        return self._fmt_or_dash(total_n), self._fmt_or_dash(total_n1)

    def _fx_produits_differes_subventions(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["131", "132"]

        val_n = self._solde_passif_par_prefixes(annee_courante, comptes)
        val_n1 = self._solde_passif_par_prefixes(annee_precedente, comptes)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_emprunts_dettes_plus_un_an(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["1611", "1631", "1641", "1651", "1671", "1681"]

        val_n = self._solde_passif_par_prefixes(annee_courante, comptes)
        val_n1 = self._solde_passif_par_prefixes(annee_precedente, comptes)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_provisions_produits_constates_avance(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["158"]

        val_n = self._solde_passif_par_prefixes(annee_courante, comptes)
        val_n1 = self._solde_passif_par_prefixes(annee_precedente, comptes)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_dettes_court_terme_moins_un_an(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["1612", "1632", "1642", "1652", "1672", "1682"]

        val_n = self._solde_passif_par_prefixes(annee_courante, comptes)
        val_n1 = self._solde_passif_par_prefixes(annee_precedente, comptes)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_fournisseurs_comptes_rattaches(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["401", "403", "404", "405", "408", "409", "4091", "4096", "4098"]

        val_n = self._solde_passif_par_prefixes(annee_courante, comptes)
        val_n1 = self._solde_passif_par_prefixes(annee_precedente, comptes)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_provisions_passifs_courants(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["481", "487"]

        val_n = self._solde_passif_par_prefixes(annee_courante, comptes)
        val_n1 = self._solde_passif_par_prefixes(annee_precedente, comptes)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_autres_dettes(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        comptes = ["451", "455", "467", "468", "456"]

        val_n = self._solde_passif_par_prefixes(annee_courante, comptes)
        val_n1 = self._solde_passif_par_prefixes(annee_precedente, comptes)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_tresorerie_passif_decouvert(self):
        annee_courante, annee_precedente = self._annees_selectionnees()

        solde_512_courant = self._sum_soldes(["512"], annee_courante)
        solde_512_precedent = self._sum_soldes(["512"], annee_precedente)

        val_n = max(0.0, -solde_512_courant)
        val_n1 = max(0.0, -solde_512_precedent)

        return self._fmt_or_dash(val_n), self._fmt_or_dash(val_n1)

    def _fx_total_passifs_non_courants(self):
        rubriques = [
            "Produits différé - Subventions d'investissement",
            "Impôt différés",
            "Emprunts et dettes financières +1ans",
            "Provision et produits constatés d'avance",
        ]

        total_n = 0.0
        total_n1 = 0.0
        for rubrique in rubriques:
            val_n, val_n1 = self._valeurs_rubrique_numeriques(rubrique)
            total_n += val_n
            total_n1 += val_n1

        return self._fmt_or_dash(total_n), self._fmt_or_dash(total_n1)

    def _fx_total_passifs_courants(self):
        rubriques = [
            "Dettes court terme -1ans",
            "Fournisseurs et comptes rattachés",
            "Provision et produits constatés d'avance (passifs courants)",
            "Autres dettes",
            "Compte de trésorerie passifs (Découvert)",
        ]

        total_n = 0.0
        total_n1 = 0.0
        for rubrique in rubriques:
            val_n, val_n1 = self._valeurs_rubrique_numeriques(rubrique)
            total_n += val_n
            total_n1 += val_n1

        return self._fmt_or_dash(total_n), self._fmt_or_dash(total_n1)

    def _fx_total_des_passifs(self):
        rubriques = [
            "TOTAL CAPITAUX PROPORES",
            "TOTAL PASSIFS NON COURANTS",
            "TOTAL PASSIFS COURANTS",
        ]

        total_n = 0.0
        total_n1 = 0.0
        for rubrique in rubriques:
            val_n, val_n1 = self._valeurs_rubrique_numeriques(rubrique)
            total_n += val_n
            total_n1 += val_n1

        return self._fmt_or_dash(total_n), self._fmt_or_dash(total_n1)

    def _clean_excel_value(self, value):
        if pd.isna(value):
            return ""
        if isinstance(value, (int, float)):
            if float(value).is_integer():
                return str(int(value))
            return str(value)
        txt = str(value).strip()
        return txt

    def _ensure_mapping(self, force=False):
        if os.path.exists(MAPPING_CSV) and not force:
            return

        default_rows = [
            (1, "CAPITAUX PROPRES", "", ""),
            (2, "Fond propre", "", ""),
            (3, "Primes et réserves", "Solde du compte", "Solde du compte"),
            (4, "Ecart d'évaluation", "0", "0"),
            (5, "Ecart d'équivalence", "50000000", "0"),
            (6, "Report à nouveau", "0", "0"),
            (7, "Résultat net", "0", "0"),
            (8, "Part de la société consolidant", "", ""),
            (9, "Part des minoritaires", "", ""),
            (10, "TOTAL CAPITAUX PROPORES", "", ""),
            (11, "PASSIFS NON COURANTS", "", ""),
            (12, "Produits différé - Subventions d'investissement", "0", "0"),
            (13, "Impôt différés", "0", "0"),
            (14, "Emprunts et dettes financières +1ans", "0", "0"),
            (15, "Provision et produits constatés d'avance", "0", "0"),
            (16, "TOTAL PASSIFS NON COURANTS", "0", ""),
            (17, "Dettes court terme -1ans", "0", "0"),
            (18, "Fournisseurs et comptes rattachés", "350000000", "0"),
            (
                19,
                "Provision et produits constatés d'avance (passifs courants)",
                "0",
                "0",
            ),
            (20, "Autres dettes", "0", "0"),
            (21, "Compte de trésorerie passifs (Découvert)", "0", "0"),
            (22, "TOTAL PASSIFS COURANTS", "350000000", "0"),
            (23, "TOTAL DES PASSIFS", "", ""),
        ]

        if os.path.exists(SOURCE_XLSX):
            try:
                df_xlsx = pd.read_excel(SOURCE_XLSX)
                if df_xlsx.shape[1] >= 3:
                    subset = df_xlsx.iloc[:, :3].copy()
                    subset.columns = ["rubrique", "valeur_n", "valeur_n1"]
                    rows = []
                    ordre = 1
                    for _, row in subset.iterrows():
                        rubrique = self._clean_excel_value(row["rubrique"])
                        if not rubrique:
                            continue
                        rows.append(
                            (
                                ordre,
                                rubrique,
                                self._clean_excel_value(row["valeur_n"]),
                                self._clean_excel_value(row["valeur_n1"]),
                            )
                        )
                        ordre += 1
                    if rows:
                        default_rows = rows
            except Exception:
                pass

        df_map = pd.DataFrame(
            default_rows, columns=["ordre", "rubrique", "valeur_n", "valeur_n1"]
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
        self.coherence_var = tk.StringVar(value="Contrôle cohérence: en attente")
        self.coherence_label = tk.Label(
            top, textvariable=self.coherence_var, fg="#444444"
        )
        self.coherence_label.pack(side=tk.LEFT, padx=(14, 0))
        ttk.Button(top, text="Exporter PDF", command=self._export_pdf).pack(
            side=tk.RIGHT, padx=(6, 0)
        )
        ttk.Button(top, text="Exporter Excel", command=self._export_excel).pack(
            side=tk.RIGHT, padx=(6, 0)
        )
        ttk.Button(top, text="Actualiser", command=self._charger_tableau).pack(
            side=tk.RIGHT
        )

        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        cols = ("Période", "Valeur N", "Valeur N-1")
        self.tree = ttk.Treeview(body, columns=cols, show="headings")

        for col in cols:
            anchor = tk.W if col == "Période" else tk.E
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor=anchor, width=480 if col == "Période" else 220)

        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(body, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

    def _appliquer_formule(self, rubrique, valeurs_mapping):
        formule = self.formules.get(rubrique)
        if callable(formule):
            resultat = formule()
            if isinstance(resultat, tuple) and len(resultat) == 2:
                return resultat
        return valeurs_mapping

    def _solde_actif_par_prefixes(self, annee, prefixes):
        total = 0.0
        for compte in self.comptes:
            if any(compte.startswith(p) for p in prefixes):
                debit = float(self.debit_group.get((annee, compte), 0.0))
                credit = float(self.credit_group.get((annee, compte), 0.0))
                total += debit - credit
        return total

    def _calcul_total_actifs(self, annee):
        net_ecart_acquisition = self._solde_actif_par_prefixes(
            annee, ["29", "290", "291", "292", "293", "296", "297"]
        )

        incorp_brut = self._solde_actif_par_prefixes(
            annee, ["20", "203", "204", "205", "207", "208"]
        )
        incorp_amort = max(
            0.0,
            -self._solde_actif_par_prefixes(
                annee, ["280", "2803", "2804", "2807", "2808"]
            ),
        )
        net_incorp = incorp_brut - incorp_amort

        corporelles_brut = self._solde_actif_par_prefixes(
            annee,
            [
                "21",
                "211",
                "212",
                "213",
                "215",
                "2151",
                "2154",
                "2155",
                "2157",
                "218",
                "2181",
                "2182",
                "2183",
                "2184",
                "2185",
                "2186",
                "221",
                "22",
                "222",
                "223",
                "225",
                "228",
                "229",
                "232",
            ],
        )
        corporelles_amort = max(
            0.0,
            -self._solde_actif_par_prefixes(
                annee, ["281", "2811", "2812", "2813", "2815", "2818", "282", "291"]
            ),
        )
        net_corporelles = corporelles_brut - corporelles_amort

        net_titres = self._solde_actif_par_prefixes(
            annee, ["26", "261", "262", "265", "266", "267", "268", "269"]
        )
        net_participations = self._solde_actif_par_prefixes(
            annee, ["27", "271", "272", "273", "274", "275", "276", "277", "279"]
        )

        net_stock = self._solde_actif_par_prefixes(
            annee,
            [
                "31",
                "32",
                "321",
                "322",
                "3221",
                "3222",
                "3223",
                "3224",
                "3225",
                "326",
                "3261",
                "3267",
                "33",
                "331",
                "335",
                "35",
                "351",
                "355",
                "358",
                "37",
                "38",
                "39",
                "391",
                "392",
                "3921",
                "3922",
                "393",
                "394",
                "395",
                "397",
                "398",
            ],
        )
        net_clients = self._solde_actif_par_prefixes(
            annee, ["411", "409", "4091", "4096", "4098", "416", "417", "418"]
        )
        net_impot = self._solde_actif_par_prefixes(annee, ["4456", "4487"])
        net_caisse = self._solde_actif_par_prefixes(annee, ["53"])
        net_banque = max(0.0, self._solde_actif_par_prefixes(annee, ["512"]))

        return (
            net_ecart_acquisition
            + net_incorp
            + net_corporelles
            + net_titres
            + net_participations
            + net_stock
            + net_clients
            + net_impot
            + net_caisse
            + net_banque
        )

    def _mettre_a_jour_controle_coherence(self):
        annee_courante, _ = self._annees_selectionnees()
        total_passifs_n, _ = self._valeurs_rubrique_numeriques("TOTAL DES PASSIFS")
        total_actifs_n = self._calcul_total_actifs(annee_courante)

        ecart = total_passifs_n - total_actifs_n
        if abs(ecart) < 0.01:
            self.coherence_var.set(
                f"Contrôle cohérence ({annee_courante}): OK | Passifs = {self._fmt_fr(total_passifs_n)} = Actifs"
            )
            self.coherence_label.configure(fg="green")
        else:
            self.coherence_var.set(
                f"Contrôle cohérence ({annee_courante}): ÉCART | Passifs = {self._fmt_fr(total_passifs_n)} | Actifs = {self._fmt_fr(total_actifs_n)} | Δ = {self._fmt_fr(ecart)}"
            )
            self.coherence_label.configure(fg="red")

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
            val_n = row.get("valeur_n", "")
            val_n1 = row.get("valeur_n1", "")
            valeurs = self._appliquer_formule(rubrique, (val_n, val_n1))
            self.tree.insert(
                "",
                tk.END,
                values=(
                    rubrique,
                    self._fmt_display_or_dash(valeurs[0]),
                    self._fmt_display_or_dash(valeurs[1]),
                ),
            )

        self._mettre_a_jour_controle_coherence()

    def _export_excel(self):
        export_treeview_to_excel(self.tree, "bilan_passif.xlsx", self)

    def _export_pdf(self):
        export_treeview_to_pdf(self.tree, "Bilan passif", "bilan_passif.pdf", self)
