"""
Module fenêtre : Bilan actif (PCG 2005 Madagascar)
Affiche le tableau de base selon la matrice fournie.
"""

import os
import tkinter as tk
from tkinter import ttk

import pandas as pd

from finances.tsarakonta.config import CONFIG
from finances.tsarakonta.models.data import DataManager, PCGManager
from finances.tsarakonta.utils.exports import export_treeview_to_excel, export_treeview_to_pdf

from .settings import format_header_text, load_header_settings


class BilanActifWindow(tk.Toplevel):
    """Fenêtre du bilan actif"""

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

        self.pcg_comptes, self.pcg_numeros, self.pcg_dict = PCGManager.charger_pcg()
        self._prefix_cache = {}

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

        self.title("Bilan actif")
        self.geometry("1100x650")
        self.formules = self._initialiser_formules()
        self._preparer_journal()

        self._build_ui()
        self._charger_tableau()

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

    def _solde_par_prefixes(self, annee, prefixes):
        prefixes = self._prefixes_valides(prefixes)
        total = 0.0
        for compte in self.comptes:
            if any(compte.startswith(p) for p in prefixes):
                debit = float(self.debit_group.get((annee, compte), 0.0))
                credit = float(self.credit_group.get((annee, compte), 0.0))
                total += debit - credit
        return total

    def _prefixes_valides(self, prefixes):
        key = tuple(prefixes)
        if key in self._prefix_cache:
            return self._prefix_cache[key]

        if not self.pcg_numeros:
            self._prefix_cache[key] = list(prefixes)
            return self._prefix_cache[key]

        valides = [
            p
            for p in prefixes
            if any(str(num).startswith(p) for num in self.pcg_numeros)
        ]
        self._prefix_cache[key] = valides if valides else list(prefixes)
        return self._prefix_cache[key]

    def _valeurs_actif(
        self,
        annee_courante,
        annee_precedente,
        prefixes_brut,
        prefixes_amort=None,
        positive_only=False,
    ):
        prefixes_amort = prefixes_amort or []

        brut_courant = self._solde_par_prefixes(annee_courante, prefixes_brut)
        brut_precedent = self._solde_par_prefixes(annee_precedente, prefixes_brut)

        if positive_only:
            brut_courant = max(0.0, brut_courant)
            brut_precedent = max(0.0, brut_precedent)

        amort_courant = (
            max(0.0, -self._solde_par_prefixes(annee_courante, prefixes_amort))
            if prefixes_amort
            else 0.0
        )
        amort_precedent = (
            max(0.0, -self._solde_par_prefixes(annee_precedente, prefixes_amort))
            if prefixes_amort
            else 0.0
        )

        net_courant = brut_courant - amort_courant
        net_precedent = brut_precedent - amort_precedent

        return brut_courant, amort_courant, net_courant, net_precedent

    def _format_valeurs(self, brut, amort, net_n, net_n1):
        return (
            self._fmt_or_dash(brut),
            self._fmt_or_dash(amort),
            self._fmt_or_dash(net_n),
            self._fmt_or_dash(net_n1),
        )

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

    def _initialiser_formules(self):
        return {
            "Ecart d'acquisition": self._fx_ecart_acquisition,
            "Immobilisations incorporelles": self._fx_immobilisations_incorporelles,
            "Immobilisations corporelles": self._fx_immobilisations_corporelles,
            "Immobilisation financières": self._fx_immobilisations_financieres,
            "Titres mis en équivalence": self._fx_titres_mis_equivalence,
            "Autres participations et créances rattachées": self._fx_autres_participations,
            "Autres titres immobilisés": self._fx_autres_titres_immobilises,
            "Prêt et autres immobilisations financières": self._fx_prets_autres_immobilisations,
            "TOTAL ACTIFS NON COURANTS": self._fx_total_actifs_non_courants,
            "Stock et en cours": self._fx_stock_en_cours,
            "Créances et emplois assimilés": self._fx_creances_emplois_assimiles,
            "Client et autres débiteurs": self._fx_clients_autres_debiteurs,
            "Impôt": self._fx_impot,
            "Autres créances et actifs assimilés": self._fx_autres_creances_actifs,
            "Trésorerie et équivalent de trésoreries": self._fx_tresorerie_equivalents,
            "Caisse": self._fx_caisse,
            "Banque": self._fx_banque,
            "TOTAL DES ACTIFS COURANTS": self._fx_total_actifs_courants,
            "TOTAL DES ACTIFS": self._fx_total_des_actifs,
        }

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

        ttk.Label(
            top, text="(Colonnes: N en brut/amortissement/net et N-1 en net)"
        ).pack(side=tk.LEFT, padx=(10, 0))
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

        cols = (
            "Rubrique",
            "Valeur Brute",
            "Amortissement/Provision",
            "Valeur Nette N",
            "Valeur Nette N-1",
        )
        self.tree = ttk.Treeview(body, columns=cols, show="headings")

        for col in cols:
            anchor = tk.W if col == "Rubrique" else tk.E
            self.tree.heading(col, text=col)
            self.tree.column(
                col, anchor=anchor, width=360 if col == "Rubrique" else 180
            )

        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(body, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

    def _get_lignes_mapping(self):
        return [
            ("ACTIFS NON COURANTS", "", "", "", ""),
            ("Ecart d'acquisition", "", "", "", ""),
            ("Immobilisations incorporelles", "-", "", "-", ""),
            (
                "Immobilisations corporelles",
                "125 000 000,00",
                "23 680 555,56",
                "101 319 444,44",
                "126 666 666,67",
            ),
            ("Immobilisation financières", "-", "", "-", "-"),
            ("Titres mis en équivalence", "", "", "", ""),
            ("Autres participations et créances rattachées", "", "", "", ""),
            ("Autres titres immobilisés", "", "", "", ""),
            ("Prêt et autres immobilisations financières", "", "", "", ""),
            (
                "TOTAL ACTIFS NON COURANTS",
                "125000000",
                "",
                "101 319 444,44",
                "126 666 666,67",
            ),
            ("", "", "", "", ""),
            ("ACTIFS COURANTS", "", "", "", ""),
            ("Stock et en cours", "150 000 000,00", "-", "150 000 000,00", "-"),
            ("Créances et emplois assimilés", "", "", "-", ""),
            ("Client et autres débiteurs", "110 000 000,00", "", "110 000 000,00", "-"),
            ("Impôt", "25 000 000,00", "", "25 000 000,00", ""),
            ("Autres créances et actifs assimilés", "-", "", "-", "-"),
            ("Trésorerie et équivalent de trésoreries", "", "", "", ""),
            ("Caisse", "-", "", "-", "-"),
            ("Banque", "175 000 000,00", "", "175 000 000,00", "-"),
            ("TOTAL DES ACTIFS COURANTS", "", "", "460 000 000,00", "-"),
            ("TOTAL DES ACTIFS", "", "", "561 319 444,44", "126 666 666,67"),
        ]

    def _appliquer_formule(self, rubrique, valeurs_mapping):
        formule = self.formules.get(rubrique)
        if callable(formule):
            resultat = formule()
            if isinstance(resultat, tuple) and len(resultat) == 4:
                return resultat
        return valeurs_mapping

    def _valeurs_rubrique_numeriques(self, rubrique):
        for row in self._get_lignes_mapping():
            if row[0] == rubrique:
                valeurs = self._appliquer_formule(rubrique, row[1:])
                return tuple(self._to_float(v) for v in valeurs)
        return 0.0, 0.0, 0.0, 0.0

    def _annees_selectionnees(self):
        annee_courante = self.annee_var.get() if hasattr(self, "annee_var") else "2025"
        try:
            annee_precedente = str(int(annee_courante) - 1)
        except ValueError:
            annee_courante = "2025"
            annee_precedente = "2024"
        return annee_courante, annee_precedente

    def _fx_ecart_acquisition(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante, annee_precedente, ["207"], ["2807", "2907"]
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_immobilisations_incorporelles(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante, annee_precedente, ["20"], ["280", "290"]
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_immobilisations_corporelles(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante,
            annee_precedente,
            ["21", "22", "23"],
            ["281", "282", "283", "291", "292", "293"],
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_immobilisations_financieres(self):
        rubriques = [
            "Titres mis en équivalence",
            "Autres participations et créances rattachées",
            "Autres titres immobilisés",
            "Prêt et autres immobilisations financières",
        ]
        total_brut = total_amort = total_net_n = total_net_n1 = 0.0
        for rubrique in rubriques:
            brut, amort, net_n, net_n1 = self._valeurs_rubrique_numeriques(rubrique)
            total_brut += brut
            total_amort += amort
            total_net_n += net_n
            total_net_n1 += net_n1
        return self._format_valeurs(total_brut, total_amort, total_net_n, total_net_n1)

    def _fx_titres_mis_equivalence(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante, annee_precedente, ["261", "262"], ["2961", "2962"]
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_autres_participations(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante,
            annee_precedente,
            ["265", "266", "267", "268"],
            ["2965", "2966", "2967", "2968"],
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_autres_titres_immobilises(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante,
            annee_precedente,
            ["271", "272", "273"],
            ["2971", "2972", "2973"],
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_prets_autres_immobilisations(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante,
            annee_precedente,
            ["274", "275", "276", "277", "279"],
            ["2974", "2975", "2976", "2977", "2979"],
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_total_actifs_non_courants(self):
        rubriques_non_courantes = [
            "Ecart d'acquisition",
            "Immobilisations incorporelles",
            "Immobilisations corporelles",
            "Immobilisation financières",
        ]

        total_brut = 0.0
        total_amort = 0.0
        total_net_n = 0.0
        total_net_n1 = 0.0

        for rubrique in rubriques_non_courantes:
            brut, amort, net_n, net_n1 = self._valeurs_rubrique_numeriques(rubrique)
            total_brut += brut
            total_amort += amort
            total_net_n += net_n
            total_net_n1 += net_n1

        return (
            self._fmt_or_dash(total_brut),
            self._fmt_or_dash(total_amort),
            self._fmt_or_dash(total_net_n),
            self._fmt_or_dash(total_net_n1),
        )

    def _fx_stock_en_cours(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante,
            annee_precedente,
            ["31", "32", "33", "35", "37", "38"],
            ["39"],
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_creances_emplois_assimiles(self):
        rubriques = [
            "Client et autres débiteurs",
            "Impôt",
            "Autres créances et actifs assimilés",
        ]
        total_brut = total_amort = total_net_n = total_net_n1 = 0.0
        for rubrique in rubriques:
            brut, amort, net_n, net_n1 = self._valeurs_rubrique_numeriques(rubrique)
            total_brut += brut
            total_amort += amort
            total_net_n += net_n
            total_net_n1 += net_n1
        return self._format_valeurs(total_brut, total_amort, total_net_n, total_net_n1)

    def _fx_clients_autres_debiteurs(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante, annee_precedente, ["41"], ["491", "496", "498"]
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_impot(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante, annee_precedente, ["4456", "4487"], []
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_autres_creances_actifs(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante,
            annee_precedente,
            ["42", "43", "44", "46", "467", "468"],
            ["495", "496", "497", "498"],
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_tresorerie_equivalents(self):
        rubriques = ["Caisse", "Banque"]
        total_brut = total_amort = total_net_n = total_net_n1 = 0.0
        for rubrique in rubriques:
            brut, amort, net_n, net_n1 = self._valeurs_rubrique_numeriques(rubrique)
            total_brut += brut
            total_amort += amort
            total_net_n += net_n
            total_net_n1 += net_n1
        return self._format_valeurs(total_brut, total_amort, total_net_n, total_net_n1)

    def _fx_caisse(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante, annee_precedente, ["53"], []
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_banque(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        brut, amort, net_n, net_n1 = self._valeurs_actif(
            annee_courante,
            annee_precedente,
            ["512", "514", "515", "518"],
            [],
            positive_only=True,
        )
        return self._format_valeurs(brut, amort, net_n, net_n1)

    def _fx_total_actifs_courants(self):
        rubriques_courantes = [
            "Stock et en cours",
            "Créances et emplois assimilés",
            "Trésorerie et équivalent de trésoreries",
        ]

        total_brut = 0.0
        total_amort = 0.0
        total_net_n = 0.0
        total_net_n1 = 0.0

        for rubrique in rubriques_courantes:
            brut, amort, net_n, net_n1 = self._valeurs_rubrique_numeriques(rubrique)
            total_brut += brut
            total_amort += amort
            total_net_n += net_n
            total_net_n1 += net_n1

        return (
            self._fmt_or_dash(total_brut),
            self._fmt_or_dash(total_amort),
            self._fmt_or_dash(total_net_n),
            self._fmt_or_dash(total_net_n1),
        )

    def _fx_total_des_actifs(self):
        brut_nc, amort_nc, net_n_nc, net_n1_nc = self._valeurs_rubrique_numeriques(
            "TOTAL ACTIFS NON COURANTS"
        )
        brut_c, amort_c, net_n_c, net_n1_c = self._valeurs_rubrique_numeriques(
            "TOTAL DES ACTIFS COURANTS"
        )

        total_brut = brut_nc + brut_c
        total_amort = amort_nc + amort_c
        total_net_n = net_n_nc + net_n_c
        total_net_n1 = net_n1_nc + net_n1_c

        return (
            self._fmt_or_dash(total_brut),
            self._fmt_or_dash(total_amort),
            self._fmt_or_dash(total_net_n),
            self._fmt_or_dash(total_net_n1),
        )

    def _charger_tableau(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        self.tree.heading("Valeur Brute", text=f"{annee_courante} Valeur Brute")
        self.tree.heading(
            "Amortissement/Provision", text=f"{annee_courante} Amortissement/Provision"
        )
        self.tree.heading("Valeur Nette N", text=f"{annee_courante} Valeur Nette")
        self.tree.heading("Valeur Nette N-1", text=f"{annee_precedente} Valeur Nette")

        for it in self.tree.get_children():
            self.tree.delete(it)

        for row in self._get_lignes_mapping():
            rubrique = row[0]
            if rubrique in self.formules:
                valeurs = self._appliquer_formule(rubrique, row[1:])
                self.tree.insert("", tk.END, values=(rubrique, *valeurs))
            else:
                self.tree.insert(
                    "",
                    tk.END,
                    values=(
                        row[0],
                        self._fmt_display_or_dash(row[1]),
                        self._fmt_display_or_dash(row[2]),
                        self._fmt_display_or_dash(row[3]),
                        self._fmt_display_or_dash(row[4]),
                    ),
                )

    def _export_excel(self):
        export_treeview_to_excel(self.tree, "bilan_actif.xlsx", self)

    def _export_pdf(self):
        export_treeview_to_pdf(self.tree, "Bilan actif", "bilan_actif.pdf", self)
