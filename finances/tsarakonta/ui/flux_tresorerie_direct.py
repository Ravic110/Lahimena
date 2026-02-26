"""
Module fenêtre : Tableau de flux de trésorerie (méthode directe)
Affiche le mapping tel que défini dans le fichier Excel source, avec formules par rubrique.
"""

import os
import tkinter as tk
import unicodedata
from tkinter import messagebox, ttk

import pandas as pd

from finances.tsarakonta.config import CONFIG
from finances.tsarakonta.models.data import DataManager
from finances.tsarakonta.utils.exports import export_treeview_to_excel, export_treeview_to_pdf

from .settings import format_header_text, load_header_settings

STATE_FOLDER = os.path.join(os.path.dirname(os.path.dirname(__file__)), "EtatFiFolder")
MAPPING_CSV = os.path.join(STATE_FOLDER, "mapping_flux_tresorerie_direct.csv")


class FluxTresorerieDirectWindow(tk.Toplevel):
    """Fenêtre d'affichage du tableau de flux de trésorerie (méthode directe)."""

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        self.title("Tableau de flux de trésorerie (méthode directe)")
        self.geometry("1120x680")

        os.makedirs(STATE_FOLDER, exist_ok=True)
        self.source_xlsx = self._resolve_source_excel()
        self.annees = self._charger_annees_disponibles()
        self.header_settings = load_header_settings()
        self.header_text = format_header_text(self.header_settings)
        self._preparer_journal()
        self._ensure_mapping(force=bool(self.source_xlsx))

        self._build_ui()
        self._charger_tableau()

    def _resolve_source_excel(self):
        base_dir = os.path.dirname(os.path.dirname(__file__))
        candidates = [
            os.path.join(base_dir, "Flux tresorerie directe.xlsx"),
            os.path.join(base_dir, "Flux de tresorerie directe.xlsx"),
            os.path.join(base_dir, "Tableau de flux de trésorerie.xlsx"),
            os.path.join(base_dir, "Tableaux Flux de trésorerie.xlsx"),
        ]
        for path in candidates:
            if os.path.exists(path):
                return path
        return None

    def _charger_annees_disponibles(self):
        base_dir = os.path.dirname(os.path.dirname(__file__))
        candidates = [
            os.path.join(base_dir, "LivreCompta.xlsx"),
            os.path.join(base_dir, "EtatFiFolder", "LivreCompta.xlsx"),
        ]

        df = None
        for fichier in candidates:
            if os.path.exists(fichier):
                df = DataManager.charger_feuille(fichier, CONFIG["feuille_journal"])
                if df is not None:
                    break

        annees_fixes = [str(y) for y in range(2020, 2031)]

        if df is None or "Année" not in df.columns:
            return sorted(annees_fixes, reverse=True)

        annees_data = df["Année"].dropna().astype(str).unique().tolist()
        toutes_annees = set(annees_fixes) | set(annees_data)

        def _sort_key(v):
            try:
                return int(v)
            except Exception:
                return -9999

        return sorted(toutes_annees, key=_sort_key, reverse=True)

    def _to_float(self, value):
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        txt = (
            str(value).strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
        )
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

    def _prepare_text(self, txt):
        base = unicodedata.normalize("NFD", str(txt).lower())
        base = "".join(ch for ch in base if unicodedata.category(ch) != "Mn")
        base = (
            base.replace("'", " ").replace("-", " ").replace("(", " ").replace(")", " ")
        )
        base = " ".join(base.split())
        return base

    def _preparer_journal(self):
        base_dir = os.path.dirname(os.path.dirname(__file__))
        candidates = [
            os.path.join(base_dir, "LivreCompta.xlsx"),
            os.path.join(base_dir, "EtatFiFolder", "LivreCompta.xlsx"),
        ]

        df = None
        for fichier in candidates:
            if os.path.exists(fichier):
                df = DataManager.charger_feuille(fichier, CONFIG["feuille_journal"])
                if df is not None:
                    break

        if df is None:
            df = pd.DataFrame(columns=CONFIG["colonnes_journal"])

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
        self.df_journal = df

        comptes_debit = set(df["CompteDébit"].dropna().astype(str).tolist())
        comptes_credit = set(df["CompteCrédit"].dropna().astype(str).tolist())
        self.comptes = sorted((comptes_debit | comptes_credit) - {""})

    def _annees_selectionnees(self):
        annee_courante = self.annee_var.get() if hasattr(self, "annee_var") else "2025"
        try:
            annee_precedente = str(int(annee_courante) - 1)
        except ValueError:
            annee_courante = "2025"
            annee_precedente = "2024"
        return annee_courante, annee_precedente

    def _is_colonne_n(self, col):
        txt = str(col).lower().replace(" ", "").replace("_", "").replace("-", "")
        return txt in ["valeurn", "n", "montantn", "valeurnetn"]

    def _is_colonne_n1(self, col):
        txt = str(col).lower().replace(" ", "").replace("_", "").replace("-", "")
        return (
            txt in ["valeurn1", "n1", "montantn1", "valeurnetn1"]
            or "n-1" in str(col).lower()
        )

    def _solde_compte(self, annee, compte):
        debit = float(self.debit_group.get((annee, compte), 0.0))
        credit = float(self.credit_group.get((annee, compte), 0.0))
        return debit - credit

    def _sum_actif(self, annee, prefixes):
        total = 0.0
        for compte in self.comptes:
            if any(compte.startswith(p) for p in prefixes):
                total += self._solde_compte(annee, compte)
        return total

    def _sum_passif(self, annee, prefixes):
        total = 0.0
        for compte in self.comptes:
            if any(compte.startswith(p) for p in prefixes):
                total += -self._solde_compte(annee, compte)
        return total

    def _sum_charge(self, annee, prefixes):
        return self._sum_actif(annee, prefixes)

    def _sum_produit(self, annee, prefixes):
        return self._sum_passif(annee, prefixes)

    def _format_montant(self, value):
        return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")

    def _format_or_dash(self, value):
        return "-" if abs(float(value)) < 0.005 else self._format_montant(value)

    def _normaliser_affichage_montants(self, df_map):
        if df_map.shape[1] < 3:
            return df_map

        for col in df_map.columns[1:3]:
            for idx, raw in df_map[col].items():
                txt = str(raw).strip()
                if txt == "":
                    df_map.at[idx, col] = "-"
                    continue

                cleaned = txt.replace("\u00a0", "").replace(" ", "").replace(",", ".")
                try:
                    if abs(float(cleaned)) < 0.005:
                        df_map.at[idx, col] = "-"
                except Exception:
                    pass
        return df_map

    def _is_tresorerie(self, compte):
        return any(
            str(compte).startswith(p)
            for p in ["53", "57", "512", "514", "515", "518", "519"]
        )

    def _collect_mouvements_tresorerie(self, annee):
        inflows = {}
        outflows = {}

        rows = self.df_journal[self.df_journal["Année"] == str(annee)]
        for _, row in rows.iterrows():
            compte_debit = self._normaliser_compte(row.get("CompteDébit", ""))
            compte_credit = self._normaliser_compte(row.get("CompteCrédit", ""))
            montant_debit = self._to_float(row.get("MontantDébit", 0.0))
            montant_credit = self._to_float(row.get("MontantCrédit", 0.0))

            debit_tresorerie = self._is_tresorerie(compte_debit)
            credit_tresorerie = self._is_tresorerie(compte_credit)

            if debit_tresorerie and not credit_tresorerie and montant_debit > 0:
                inflows[compte_credit] = inflows.get(compte_credit, 0.0) + montant_debit

            if credit_tresorerie and not debit_tresorerie and montant_credit > 0:
                outflows[compte_debit] = (
                    outflows.get(compte_debit, 0.0) + montant_credit
                )

        return inflows, outflows

    def _sum_mouvements(self, mouvements, prefixes):
        total = 0.0
        for compte, montant in mouvements.items():
            if any(str(compte).startswith(p) for p in prefixes):
                total += float(montant)
        return total

    def _resultat_net_exercice(self, annee):
        chiffres_affaires = self._sum_produit(
            annee,
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
        )
        production_stockee = self._sum_produit(annee, ["71", "711", "713", "714"])
        production_immobilisee = self._sum_produit(annee, ["72", "721", "722"])

        achat_consommes = self._sum_charge(
            annee,
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
        )
        services_exterieurs = self._sum_charge(
            annee,
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
        )
        charges_personnel = self._sum_charge(
            annee, ["64", "641", "644", "645", "646", "647", "648"]
        )
        impots_taxes = self._sum_charge(annee, ["63", "631", "635"])

        autres_produits = self._sum_produit(
            annee,
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
        )
        autres_charges = self._sum_charge(
            annee, ["65", "651", "652", "653", "654", "655", "656", "657", "658"]
        )

        dotations = self._sum_charge(annee, ["68", "681", "685"])
        produits_financiers = self._sum_produit(
            annee, ["76", "761", "762", "763", "764", "766", "767", "768"]
        )
        charges_financieres = self._sum_charge(
            annee, ["66", "661", "664", "665", "666", "667", "668"]
        )

        extra_produits = self._sum_produit(annee, ["77"])
        extra_charges = self._sum_charge(annee, ["67"])

        impots_exigibles = self._sum_charge(annee, ["69", "695", "698"])
        impots_differes = self._sum_charge(annee, ["692", "693"])

        production_exercice = (
            chiffres_affaires + production_stockee + production_immobilisee
        )
        consommation_exercice = achat_consommes + services_exterieurs
        valeur_ajoutee = production_exercice - consommation_exercice
        excedent_brut_exploitation = valeur_ajoutee - charges_personnel - impots_taxes
        resultat_operationnel = autres_produits - autres_charges
        resultat_financier = produits_financiers - charges_financieres
        resultat_avant_impot = (
            excedent_brut_exploitation
            - dotations
            + resultat_operationnel
            + resultat_financier
        )
        resultat_extraordinaire = extra_produits - extra_charges
        return (
            resultat_avant_impot
            - impots_exigibles
            + impots_differes
            + resultat_extraordinaire
        )

    def _calc_values_for_year(self, annee):
        try:
            annee_precedente = str(int(annee) - 1)
        except ValueError:
            annee_precedente = "2024"

        inflows, outflows = self._collect_mouvements_tresorerie(annee)

        encaissements_clients = self._sum_mouvements(inflows, ["41", "70", "71", "72"])
        versements_fournisseurs_personnel = -self._sum_mouvements(
            outflows,
            [
                "40",
                "401",
                "403",
                "404",
                "405",
                "408",
                "409",
                "42",
                "43",
                "44",
                "60",
                "61",
                "62",
                "63",
                "64",
            ],
        )
        interets_frais_payes = -self._sum_mouvements(outflows, ["66"])

        flux_avant_extra = (
            encaissements_clients
            + versements_fournisseurs_personnel
            + interets_frais_payes
        )
        flux_evenements_extra = self._sum_mouvements(
            inflows, ["77"]
        ) - self._sum_mouvements(outflows, ["67"])
        flux_activite_a = flux_avant_extra + flux_evenements_extra

        immocorp_n = self._sum_actif(annee, ["20", "21", "22", "23"])
        immocorp_n1 = self._sum_actif(annee_precedente, ["20", "21", "22", "23"])
        delta_immocorp = immocorp_n - immocorp_n1

        immo_fin_n = self._sum_actif(annee, ["26", "27"])
        immo_fin_n1 = self._sum_actif(annee_precedente, ["26", "27"])
        delta_immo_fin = immo_fin_n - immo_fin_n1

        decaissement_immo_corp = -self._sum_mouvements(
            outflows, ["20", "21", "22", "23"]
        )
        encaissement_cession_immo_corp = self._sum_mouvements(
            inflows, ["20", "21", "22", "23"]
        )
        decaissement_immo_fin = -self._sum_mouvements(outflows, ["26", "27"])
        encaissement_cession_immo_fin = self._sum_mouvements(inflows, ["26", "27"])

        if abs(decaissement_immo_corp) < 0.005:
            decaissement_immo_corp = -max(0.0, delta_immocorp)
        if abs(encaissement_cession_immo_corp) < 0.005:
            encaissement_cession_immo_corp = max(0.0, -delta_immocorp)
        if abs(decaissement_immo_fin) < 0.005:
            decaissement_immo_fin = -max(0.0, delta_immo_fin)
        if abs(encaissement_cession_immo_fin) < 0.005:
            encaissement_cession_immo_fin = max(0.0, -delta_immo_fin)

        interets_encaisses = self._sum_mouvements(
            inflows, ["76", "763", "764", "766", "767", "768"]
        )
        dividendes_quote_part_recus = self._sum_mouvements(inflows, ["761", "762"])

        flux_investissement_b = (
            decaissement_immo_corp
            + encaissement_cession_immo_corp
            + decaissement_immo_fin
            + encaissement_cession_immo_fin
            + interets_encaisses
            + dividendes_quote_part_recus
        )

        emission_actions = self._sum_mouvements(inflows, ["101", "104", "106", "108"])
        dividendes_distribues = -self._sum_mouvements(outflows, ["457"])
        encaissement_emprunts = self._sum_mouvements(inflows, ["16"])
        remboursement_emprunts = -self._sum_mouvements(outflows, ["16"])

        flux_financement_c = (
            emission_actions
            + dividendes_distribues
            + encaissement_emprunts
            + remboursement_emprunts
        )

        incidence_change = 0.0
        variation_periode = flux_activite_a + flux_investissement_b + flux_financement_c

        treso_cloture = self._sum_actif(
            annee, ["53", "57", "512", "514", "515", "518"]
        ) - self._sum_passif(annee, ["519"])
        treso_ouverture = self._sum_actif(
            annee_precedente, ["53", "57", "512", "514", "515", "518"]
        ) - self._sum_passif(annee_precedente, ["519"])

        return {
            "encaissements_clients": encaissements_clients,
            "versements_fournisseurs_personnel": versements_fournisseurs_personnel,
            "interets_frais_payes": interets_frais_payes,
            "flux_avant_extra": flux_avant_extra,
            "flux_evenements_extra": flux_evenements_extra,
            "flux_activite_a": flux_activite_a,
            "decaissement_immo_corp": decaissement_immo_corp,
            "encaissement_cession_immo_corp": encaissement_cession_immo_corp,
            "decaissement_immo_fin": decaissement_immo_fin,
            "encaissement_cession_immo_fin": encaissement_cession_immo_fin,
            "interets_encaisses": interets_encaisses,
            "dividendes_quote_part_recus": dividendes_quote_part_recus,
            "flux_investissement_b": flux_investissement_b,
            "emission_actions": emission_actions,
            "dividendes_distribues": dividendes_distribues,
            "encaissement_emprunts": encaissement_emprunts,
            "remboursement_emprunts": remboursement_emprunts,
            "flux_financement_c": flux_financement_c,
            "incidence_change": incidence_change,
            "variation_periode": variation_periode,
            "treso_ouverture": treso_ouverture,
            "treso_cloture": treso_cloture,
            "variation_tresorerie_2_1": treso_cloture - treso_ouverture,
            "rapprochement_resultat": self._resultat_net_exercice(annee),
        }

    def _apply_formules(self, df_map, annee_courante, annee_precedente):
        if df_map.shape[1] < 3:
            return df_map

        col_rubrique = df_map.columns[0]
        col_n = df_map.columns[1]
        col_n1 = df_map.columns[2]

        n_values = self._calc_values_for_year(annee_courante)
        n1_values = self._calc_values_for_year(annee_precedente)

        def get_values_by_rubrique(rubrique):
            key = self._prepare_text(rubrique)

            mapping = {
                "encaissements recus des clients": (
                    "encaissements_clients",
                    "encaissements_clients",
                ),
                "sommes versees aux fournisseurs et au personnel": (
                    "versements_fournisseurs_personnel",
                    "versements_fournisseurs_personnel",
                ),
                "interets et autres frais financiers payes": (
                    "interets_frais_payes",
                    "interets_frais_payes",
                ),
                "flux de tresorerie avant elements extraordinaires": (
                    "flux_avant_extra",
                    "flux_avant_extra",
                ),
                "flux de tresorerie lie a des evenements extraordinaire a preciser": (
                    "flux_evenements_extra",
                    "flux_evenements_extra",
                ),
                "flux de tresorerie net provenant des activites operationnelle a": (
                    "flux_activite_a",
                    "flux_activite_a",
                ),
                "decaissemnet sur acquisition d immobilisations corporelles ou icorporelles": (
                    "decaissement_immo_corp",
                    "decaissement_immo_corp",
                ),
                "decaissement sur acquisition d immobilisations corporelles ou incorporelles": (
                    "decaissement_immo_corp",
                    "decaissement_immo_corp",
                ),
                "encaissements sur cessions d immobilisations corporelles ou incorporelles": (
                    "encaissement_cession_immo_corp",
                    "encaissement_cession_immo_corp",
                ),
                "decaissements sur acquisition d immobilisations financieres": (
                    "decaissement_immo_fin",
                    "decaissement_immo_fin",
                ),
                "encaissements sur cession d immobilisations financieres": (
                    "encaissement_cession_immo_fin",
                    "encaissement_cession_immo_fin",
                ),
                "interets encaisses sur placements financieres": (
                    "interets_encaisses",
                    "interets_encaisses",
                ),
                "dividendes et quote part de resultats recus": (
                    "dividendes_quote_part_recus",
                    "dividendes_quote_part_recus",
                ),
                "flux de tresorerie net provenant des activites d investissement b": (
                    "flux_investissement_b",
                    "flux_investissement_b",
                ),
                "encaissements suite a l emission d actions": (
                    "emission_actions",
                    "emission_actions",
                ),
                "dividendes et autres distributions effectues": (
                    "dividendes_distribues",
                    "dividendes_distribues",
                ),
                "encaissements provenant d emprunts": (
                    "encaissement_emprunts",
                    "encaissement_emprunts",
                ),
                "remboursement d emprunts ou d autres dettes assimiles": (
                    "remboursement_emprunts",
                    "remboursement_emprunts",
                ),
                "flux de tresorerie net provenant des activites de financement c": (
                    "flux_financement_c",
                    "flux_financement_c",
                ),
                "incidences des variations des taux de change sur liquidites et quasi lilquidites": (
                    "incidence_change",
                    "incidence_change",
                ),
                "variation de tresorerie de la periode a b c": (
                    "variation_periode",
                    "variation_periode",
                ),
                "tresorerie et equivalents de tresorerie a l ouverture de l exercice": (
                    "treso_ouverture",
                    "treso_ouverture",
                ),
                "tresorerie et equivalents de tresorerie a la cloture de l exercice": (
                    "treso_cloture",
                    "treso_cloture",
                ),
                "variation de tresorerie de la periode": (
                    "variation_tresorerie_2_1",
                    "variation_tresorerie_2_1",
                ),
                "rapprochement avec le resultat comptable": (
                    "rapprochement_resultat",
                    "rapprochement_resultat",
                ),
            }

            if key in mapping:
                kn, kn1 = mapping[key]
                return n_values.get(kn, 0.0), n1_values.get(kn1, 0.0)
            return None, None

        for idx, row in df_map.iterrows():
            rubrique = str(row.get(col_rubrique, "")).strip()
            if not rubrique:
                continue
            val_n, val_n1 = get_values_by_rubrique(rubrique)
            if val_n is None:
                continue
            df_map.at[idx, col_n] = self._format_or_dash(val_n)
            df_map.at[idx, col_n1] = self._format_or_dash(val_n1)

        return self._normaliser_affichage_montants(df_map)

    def _detect_direct_sheet(self, xlsx_path):
        try:
            xls = pd.ExcelFile(xlsx_path)
            for sheet in xls.sheet_names:
                norm = self._prepare_text(sheet)
                if "direct" in norm:
                    return sheet
            for sheet in xls.sheet_names:
                df_head = pd.read_excel(
                    xlsx_path, sheet_name=sheet, nrows=8, dtype=str
                ).fillna("")
                flat = " ".join(df_head.astype(str).stack().tolist())
                if "encaissements recus des clients" in self._prepare_text(flat):
                    return sheet
            return xls.sheet_names[0] if xls.sheet_names else 0
        except Exception:
            return 0

    def _ensure_mapping(self, force=False):
        if os.path.exists(MAPPING_CSV) and not force:
            return

        if self.source_xlsx and os.path.exists(self.source_xlsx):
            try:
                sheet_name = self._detect_direct_sheet(self.source_xlsx)
                df_map = pd.read_excel(
                    self.source_xlsx, sheet_name=sheet_name, dtype=str
                ).fillna("")
                if df_map.shape[1] == 0:
                    df_map = pd.DataFrame(columns=["Rubriques", "N", "N-1"])
                df_map.to_csv(MAPPING_CSV, index=False, encoding="utf-8")
                return
            except Exception as exc:
                messagebox.showwarning(
                    "Attention",
                    f"Impossible de lire la source Excel. Un mapping vide sera utilisé.\nDétail: {exc}",
                    parent=self,
                )

        df_default = pd.DataFrame(columns=["Rubriques", "N", "N-1"])
        df_default.to_csv(MAPPING_CSV, index=False, encoding="utf-8")

    def _build_ui(self):
        if self.header_text:
            header = ttk.LabelFrame(self, text="En-tête société", padding=8)
            header.pack(fill=tk.X, padx=8, pady=(8, 4))
            ttk.Label(header, text=self.header_text, justify=tk.LEFT).pack(anchor=tk.W)

        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Arrêté au:").pack(side=tk.LEFT, padx=(0, 6))
        self.annee_var = tk.StringVar(value=self.annees[0] if self.annees else "2025")
        self.annee_combo = ttk.Combobox(
            top,
            values=self.annees,
            textvariable=self.annee_var,
            state="readonly",
            width=10,
        )
        self.annee_combo.pack(side=tk.LEFT, padx=(0, 12))
        self.annee_combo.bind("<<ComboboxSelected>>", lambda e: self._charger_tableau())

        source_label = (
            self.source_xlsx if self.source_xlsx else "Source Excel non trouvée"
        )
        ttk.Label(top, text=f"Source mapping: {source_label}").pack(side=tk.LEFT)

        ttk.Button(
            top, text="Actualiser depuis Excel", command=self._actualiser_depuis_excel
        ).pack(side=tk.RIGHT)

        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.tree = ttk.Treeview(body, columns=(), show="headings")

        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(body, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        actions = ttk.Frame(self, padding=(8, 0, 8, 8))
        actions.pack(fill=tk.X)
        ttk.Button(
            actions, text="Ouvrir mapping CSV", command=self._ouvrir_mapping_csv
        ).pack(side=tk.LEFT)
        ttk.Button(actions, text="Exporter Excel", command=self._export_excel).pack(
            side=tk.RIGHT, padx=6
        )
        ttk.Button(actions, text="Exporter PDF", command=self._export_pdf).pack(
            side=tk.RIGHT
        )

    def _actualiser_depuis_excel(self):
        self.source_xlsx = self._resolve_source_excel()
        self._ensure_mapping(force=True)
        self._charger_tableau()

    def _charger_tableau(self):
        annee_courante, annee_precedente = self._annees_selectionnees()

        if not os.path.exists(MAPPING_CSV):
            self._ensure_mapping()

        df_map = pd.read_csv(MAPPING_CSV, dtype=str).fillna("")
        df_map = self._apply_formules(df_map, annee_courante, annee_precedente)
        columns = list(df_map.columns)
        if not columns:
            columns = ["Rubriques"]
            df_map = pd.DataFrame(columns=columns)

        self.tree["columns"] = columns
        for idx, col in enumerate(columns):
            if idx == 1 or self._is_colonne_n(col):
                header = annee_courante
            elif idx == 2 or self._is_colonne_n1(col):
                header = annee_precedente
            else:
                header = str(col)

            self.tree.heading(col, text=header)
            width = 500 if col == columns[0] else 220
            self.tree.column(col, width=width, anchor=tk.W)

        for it in self.tree.get_children():
            self.tree.delete(it)

        for _, row in df_map.iterrows():
            values = [row.get(c, "") for c in columns]
            self.tree.insert("", tk.END, values=values)

    def _ouvrir_mapping_csv(self):
        try:
            os.startfile(MAPPING_CSV)
        except Exception:
            messagebox.showinfo(
                "Info", f"Mapping disponible ici:\n{MAPPING_CSV}", parent=self
            )

    def _export_excel(self):
        export_treeview_to_excel(self.tree, "tableau_flux_tresorerie_direct.xlsx", self)

    def _export_pdf(self):
        export_treeview_to_pdf(
            self.tree,
            "Tableau de flux de trésorerie (méthode directe)",
            "tableau_flux_tresorerie_direct.pdf",
            self,
        )
