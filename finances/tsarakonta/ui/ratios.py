"""
Fenêtres de ratios financiers (N / N-1).
"""

import os
import tkinter as tk
from tkinter import ttk

import pandas as pd

from finances.tsarakonta.config import CONFIG
from finances.tsarakonta.models.data import DataManager

from .settings import format_header_text, load_header_settings


class RatioBaseWindow(tk.Toplevel):
    """Base commune pour les fenêtres de ratios."""

    def __init__(self, parent, title, df=None):
        super().__init__(parent)
        self.parent = parent
        self.title(title)
        self.geometry("900x560")

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

    def _annees_selectionnees(self):
        annee_courante = self.annee_var.get() if hasattr(self, "annee_var") else "2025"
        try:
            annee_precedente = str(int(annee_courante) - 1)
        except ValueError:
            annee_courante = "2025"
            annee_precedente = "2024"
        return annee_courante, annee_precedente

    def _sum_by_prefixes(self, annee, prefixes, mode="actif"):
        total = 0.0
        for compte in self.comptes:
            if any(compte.startswith(p) for p in prefixes):
                debit = float(self.debit_group.get((annee, compte), 0.0))
                credit = float(self.credit_group.get((annee, compte), 0.0))
                if mode in ("actif", "charge"):
                    total += debit - credit
                else:
                    total += credit - debit
        return total

    def _safe_ratio(self, numerateur, denominateur):
        if abs(denominateur) < 1e-9:
            return None
        return numerateur / denominateur

    def _fmt_amount(self, value):
        if value is None:
            return "-"
        return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")

    def _fmt_percent(self, value):
        if value is None:
            return "-"
        return f"{(value * 100):.2f}%".replace(".", ",")

    def _fmt_coef(self, value):
        if value is None:
            return "-"
        return f"{value:.2f}".replace(".", ",")

    def _parse_display_value(self, value):
        if value is None:
            return None
        txt = str(value).strip()
        if txt in ("", "-"):
            return None
        is_percent = txt.endswith("%")
        txt = txt.replace("%", "").replace(" ", "").replace(",", ".")
        try:
            num = float(txt)
            return num / 100.0 if is_percent else num
        except ValueError:
            return None

    def _normalize_ratio_row(self, row):
        if isinstance(row, dict):
            return {
                "name": row.get("name", ""),
                "n_display": row.get("n_display", "-"),
                "n1_display": row.get("n1_display", "-"),
                "n_value": row.get("n_value"),
                "n1_value": row.get("n1_value"),
            }

        if isinstance(row, (tuple, list)) and len(row) >= 3:
            n_display = row[1]
            n1_display = row[2]
            return {
                "name": row[0],
                "n_display": n_display,
                "n1_display": n1_display,
                "n_value": self._parse_display_value(n_display),
                "n1_value": self._parse_display_value(n1_display),
            }

        return {
            "name": "",
            "n_display": "-",
            "n1_display": "-",
            "n_value": None,
            "n1_value": None,
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

        ttk.Button(top, text="Actualiser", command=self._charger_tableau).pack(
            side=tk.RIGHT
        )

        main_pane = ttk.Panedwindow(self, orient=tk.VERTICAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        body = ttk.Frame(main_pane)
        main_pane.add(body, weight=3)

        cols = ("Ratio", "Valeur N", "Valeur N-1")
        self.tree = ttk.Treeview(body, columns=cols, show="headings")
        self.tree.heading("Ratio", text="Ratio")
        self.tree.heading("Valeur N", text="Valeur N")
        self.tree.heading("Valeur N-1", text="Valeur N-1")
        self.tree.column("Ratio", anchor=tk.W, width=520)
        self.tree.column("Valeur N", anchor=tk.E, width=170)
        self.tree.column("Valeur N-1", anchor=tk.E, width=170)

        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        graph_zone = ttk.LabelFrame(main_pane, text="Diagrammes circulaires", padding=6)
        main_pane.add(graph_zone, weight=2)

        graph_zone.columnconfigure(0, weight=1)
        graph_zone.columnconfigure(1, weight=1)
        graph_zone.rowconfigure(1, weight=1)

        ttk.Label(graph_zone, text="Répartition des ratios - N").grid(
            row=0, column=0, sticky="w", padx=4, pady=(0, 4)
        )
        ttk.Label(graph_zone, text="Répartition des ratios - N-1").grid(
            row=0, column=1, sticky="w", padx=4, pady=(0, 4)
        )

        self.canvas_pie_n = tk.Canvas(
            graph_zone, bg="white", highlightthickness=1, highlightbackground="#d9d9d9"
        )
        self.canvas_pie_n1 = tk.Canvas(
            graph_zone, bg="white", highlightthickness=1, highlightbackground="#d9d9d9"
        )
        self.canvas_pie_n.grid(row=1, column=0, sticky="nsew", padx=(4, 6), pady=(0, 4))
        self.canvas_pie_n1.grid(
            row=1, column=1, sticky="nsew", padx=(6, 4), pady=(0, 4)
        )

        legend = ttk.Frame(graph_zone)
        legend.grid(row=2, column=0, columnspan=2, sticky="w", padx=4)
        ttk.Label(
            legend,
            text="Les parts sont basées sur la valeur absolue des ratios (hors valeurs nulles).",
        ).pack(side=tk.LEFT)

        self.canvas_pie_n.bind("<Configure>", lambda e: self._redraw_graphs())
        self.canvas_pie_n1.bind("<Configure>", lambda e: self._redraw_graphs())

    def _redraw_graphs(self):
        if hasattr(self, "current_rows"):
            self._dessiner_graphes(self.current_rows)

    def _dessiner_graphes(self, rows):
        self._dessiner_diagramme_circulaire(self.canvas_pie_n, rows, "n_value", "N")
        self._dessiner_diagramme_circulaire(self.canvas_pie_n1, rows, "n1_value", "N-1")

    def _dessiner_diagramme_circulaire(self, canvas, rows, key, label_periode):
        canvas.delete("all")

        width = max(canvas.winfo_width(), 360)
        height = max(canvas.winfo_height(), 220)
        left_pad = 10
        top_pad = 10
        legend_x = int(width * 0.56)

        palette = [
            "#2E86DE",
            "#48C9B0",
            "#F4D03F",
            "#AF7AC5",
            "#EB984E",
            "#5DADE2",
            "#58D68D",
            "#F5B041",
            "#E74C3C",
            "#7DCEA0",
            "#85C1E9",
            "#F8C471",
        ]

        data = []
        for r in rows:
            v = r.get(key)
            if v is None:
                continue
            if abs(v) < 1e-12:
                continue
            data.append((r["name"], abs(v), r))

        total = sum(v for _, v, _ in data)
        if total <= 0:
            canvas.create_text(
                width / 2,
                height / 2,
                text=f"Aucune donnée {label_periode}",
                fill="#777777",
            )
            return

        diameter = min(int(width * 0.48), height - 20)
        diameter = max(120, diameter)
        x1 = left_pad
        y1 = top_pad
        x2 = x1 + diameter
        y2 = y1 + diameter

        start = 0.0
        for idx, (name, value, row) in enumerate(data):
            extent = (value / total) * 360.0
            color = palette[idx % len(palette)]
            canvas.create_arc(
                x1, y1, x2, y2, start=start, extent=extent, fill=color, outline="white"
            )
            row["_pie_color"] = color
            row[f"_pie_share_{key}"] = value / total
            start += extent

        canvas.create_text(
            (x1 + x2) / 2,
            (y1 + y2) / 2,
            text=f"{label_periode}\n100%",
            justify="center",
            fill="#2c3e50",
            font=("Segoe UI", 9, "bold"),
        )

        legend_y = top_pad
        max_legend_items = min(len(data), max(4, (height - 20) // 20))
        for idx, (name, _, row) in enumerate(data[:max_legend_items]):
            share = row.get(f"_pie_share_{key}", 0.0)
            pct = f"{share * 100:.1f}%".replace(".", ",")
            color = row.get("_pie_color", "#2E86DE")
            canvas.create_rectangle(
                legend_x,
                legend_y + 4,
                legend_x + 10,
                legend_y + 14,
                fill=color,
                outline="",
            )
            canvas.create_text(
                legend_x + 15,
                legend_y + 9,
                anchor="w",
                text=f"{name[:26]} - {pct}",
                fill="#333333",
                font=("Segoe UI", 8),
            )
            legend_y += 20

        if len(data) > max_legend_items:
            canvas.create_text(
                legend_x,
                legend_y + 6,
                anchor="nw",
                text=f"... +{len(data) - max_legend_items} ratios",
                fill="#666666",
                font=("Segoe UI", 8, "italic"),
            )

    def _charger_tableau(self):
        annee_courante, annee_precedente = self._annees_selectionnees()
        self.tree.heading("Valeur N", text=annee_courante)
        self.tree.heading("Valeur N-1", text=annee_precedente)

        for it in self.tree.get_children():
            self.tree.delete(it)

        rows = [
            self._normalize_ratio_row(r)
            for r in self._calculer_ratios(annee_courante, annee_precedente)
        ]
        for r in rows:
            self.tree.insert(
                "", tk.END, values=(r["name"], r["n_display"], r["n1_display"])
            )

        self.current_rows = rows
        self._dessiner_graphes(rows)

    def _calculer_ratios(self, annee_courante, annee_precedente):
        raise NotImplementedError()


class RatioResultatNatureWindow(RatioBaseWindow):
    """Ratios calculés à partir du compte de résultat par nature."""

    def __init__(self, parent, df=None):
        super().__init__(parent, "Ratio Compte de résultat par nature", df=df)

    def _metriques_nature(self, annee):
        chiffre_affaires = self._sum_by_prefixes(
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
            mode="produit",
        )
        production_stockee = self._sum_by_prefixes(
            annee, ["71", "711", "713", "714"], mode="produit"
        )
        production_immobilisee = self._sum_by_prefixes(
            annee, ["72", "721", "722"], mode="produit"
        )
        achats_consommes = self._sum_by_prefixes(
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
            mode="charge",
        )
        services_exterieurs = self._sum_by_prefixes(
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
            mode="charge",
        )
        charges_personnel = self._sum_by_prefixes(
            annee, ["64", "641", "644", "645", "646", "647", "648"], mode="charge"
        )
        impots_taxes = self._sum_by_prefixes(annee, ["63", "631", "635"], mode="charge")
        autres_produits = self._sum_by_prefixes(
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
            mode="produit",
        )
        autres_charges = self._sum_by_prefixes(
            annee,
            ["65", "651", "652", "653", "654", "655", "656", "657", "658"],
            mode="charge",
        )
        dotations = self._sum_by_prefixes(annee, ["68", "681", "685"], mode="charge")
        produits_financiers = self._sum_by_prefixes(
            annee,
            ["76", "761", "762", "763", "764", "766", "767", "768"],
            mode="produit",
        )
        charges_financieres = self._sum_by_prefixes(
            annee, ["66", "661", "664", "665", "666", "667", "668"], mode="charge"
        )
        impots_resultat = self._sum_by_prefixes(
            annee, ["69", "695", "698"], mode="charge"
        )
        impots_differes = self._sum_by_prefixes(annee, ["692", "693"], mode="charge")
        extra_produits = self._sum_by_prefixes(annee, ["77"], mode="produit")
        extra_charges = self._sum_by_prefixes(annee, ["67"], mode="charge")

        production_exercice = (
            chiffre_affaires + production_stockee + production_immobilisee
        )
        consommation_exercice = achats_consommes + services_exterieurs
        valeur_ajoutee = production_exercice - consommation_exercice
        ebe = valeur_ajoutee - charges_personnel - impots_taxes
        resultat_operationnel = autres_produits - autres_charges
        resultat_financier = produits_financiers - charges_financieres
        resultat_avant_impot = (
            ebe - dotations + resultat_operationnel + resultat_financier
        )
        resultat_extraordinaire = extra_produits - extra_charges
        resultat_net = (
            resultat_avant_impot
            - impots_resultat
            + impots_differes
            + resultat_extraordinaire
        )

        return {
            "production_exercice": production_exercice,
            "consommation_exercice": consommation_exercice,
            "valeur_ajoutee": valeur_ajoutee,
            "ebe": ebe,
            "resultat_operationnel": resultat_operationnel,
            "charges_financieres": charges_financieres,
            "resultat_net": resultat_net,
        }

    def _calculer_ratios(self, annee_courante, annee_precedente):
        n = self._metriques_nature(annee_courante)
        n1 = self._metriques_nature(annee_precedente)

        return [
            (
                "Taux de valeur ajoutée / production",
                self._fmt_percent(
                    self._safe_ratio(n["valeur_ajoutee"], n["production_exercice"])
                ),
                self._fmt_percent(
                    self._safe_ratio(n1["valeur_ajoutee"], n1["production_exercice"])
                ),
            ),
            (
                "Taux EBE / production",
                self._fmt_percent(self._safe_ratio(n["ebe"], n["production_exercice"])),
                self._fmt_percent(
                    self._safe_ratio(n1["ebe"], n1["production_exercice"])
                ),
            ),
            (
                "Marge opérationnelle",
                self._fmt_percent(
                    self._safe_ratio(
                        n["resultat_operationnel"], n["production_exercice"]
                    )
                ),
                self._fmt_percent(
                    self._safe_ratio(
                        n1["resultat_operationnel"], n1["production_exercice"]
                    )
                ),
            ),
            (
                "Poids des charges financières",
                self._fmt_percent(
                    self._safe_ratio(n["charges_financieres"], n["production_exercice"])
                ),
                self._fmt_percent(
                    self._safe_ratio(
                        n1["charges_financieres"], n1["production_exercice"]
                    )
                ),
            ),
            (
                "Marge nette",
                self._fmt_percent(
                    self._safe_ratio(n["resultat_net"], n["production_exercice"])
                ),
                self._fmt_percent(
                    self._safe_ratio(n1["resultat_net"], n1["production_exercice"])
                ),
            ),
            (
                "Consommation / production",
                self._fmt_percent(
                    self._safe_ratio(
                        n["consommation_exercice"], n["production_exercice"]
                    )
                ),
                self._fmt_percent(
                    self._safe_ratio(
                        n1["consommation_exercice"], n1["production_exercice"]
                    )
                ),
            ),
        ]


class RatioResultatFonctionWindow(RatioBaseWindow):
    """Ratios calculés à partir du compte de résultat par fonction."""

    def __init__(self, parent, df=None):
        super().__init__(parent, "Ratio compte de résultat par fonction", df=df)

    def _metriques_fonction(self, annee):
        produits = (
            self._sum_by_prefixes(
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
                mode="produit",
            )
            + self._sum_by_prefixes(annee, ["71", "711", "713", "714"], mode="produit")
            + self._sum_by_prefixes(annee, ["72", "721", "722"], mode="produit")
        )
        cout_ventes = self._sum_by_prefixes(
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
            mode="charge",
        )
        marge_brute = produits - cout_ventes
        autres_produits = self._sum_by_prefixes(
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
            mode="produit",
        )
        cout_commerciaux = self._sum_by_prefixes(
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
            mode="charge",
        )
        charges_administratives = (
            self._sum_by_prefixes(
                annee, ["64", "641", "644", "645", "646", "647", "648"], mode="charge"
            )
            + self._sum_by_prefixes(annee, ["63", "631", "635"], mode="charge")
            + self._sum_by_prefixes(annee, ["68", "681", "685"], mode="charge")
            - self._sum_by_prefixes(
                annee,
                [
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
                ],
                mode="produit",
            )
        )
        autres_charges = self._sum_by_prefixes(
            annee,
            ["65", "651", "652", "653", "654", "655", "656", "657", "658"],
            mode="charge",
        )

        resultat_operationnel = (
            marge_brute
            + autres_produits
            - cout_commerciaux
            - charges_administratives
            - autres_charges
        )
        produits_financiers = self._sum_by_prefixes(
            annee,
            ["76", "761", "762", "763", "764", "766", "767", "768"],
            mode="produit",
        )
        charges_financieres = self._sum_by_prefixes(
            annee, ["66", "661", "664", "665", "666", "667", "668"], mode="charge"
        )
        resultat_avant_impot = (
            resultat_operationnel + produits_financiers - charges_financieres
        )
        impot_resultat = self._sum_by_prefixes(
            annee, ["69", "695", "698"], mode="charge"
        )
        impots_differes = self._sum_by_prefixes(annee, ["692", "693"], mode="charge")
        charges_extra = self._sum_by_prefixes(annee, ["67"], mode="charge")
        produits_extra = self._sum_by_prefixes(annee, ["77"], mode="produit")
        resultat_net = (
            resultat_avant_impot
            - impot_resultat
            + impots_differes
            - charges_extra
            + produits_extra
        )

        return {
            "produits": produits,
            "cout_ventes": cout_ventes,
            "marge_brute": marge_brute,
            "charges_administratives": charges_administratives,
            "resultat_operationnel": resultat_operationnel,
            "resultat_net": resultat_net,
        }

    def _calculer_ratios(self, annee_courante, annee_precedente):
        n = self._metriques_fonction(annee_courante)
        n1 = self._metriques_fonction(annee_precedente)

        return [
            (
                "Marge brute / produits ordinaires",
                self._fmt_percent(self._safe_ratio(n["marge_brute"], n["produits"])),
                self._fmt_percent(self._safe_ratio(n1["marge_brute"], n1["produits"])),
            ),
            (
                "Coût des ventes / produits ordinaires",
                self._fmt_percent(self._safe_ratio(n["cout_ventes"], n["produits"])),
                self._fmt_percent(self._safe_ratio(n1["cout_ventes"], n1["produits"])),
            ),
            (
                "Charges administratives / produits ordinaires",
                self._fmt_percent(
                    self._safe_ratio(n["charges_administratives"], n["produits"])
                ),
                self._fmt_percent(
                    self._safe_ratio(n1["charges_administratives"], n1["produits"])
                ),
            ),
            (
                "Marge opérationnelle",
                self._fmt_percent(
                    self._safe_ratio(n["resultat_operationnel"], n["produits"])
                ),
                self._fmt_percent(
                    self._safe_ratio(n1["resultat_operationnel"], n1["produits"])
                ),
            ),
            (
                "Marge nette",
                self._fmt_percent(self._safe_ratio(n["resultat_net"], n["produits"])),
                self._fmt_percent(self._safe_ratio(n1["resultat_net"], n1["produits"])),
            ),
        ]


class RatioBilanWindow(RatioBaseWindow):
    """Ratios calculés à partir des masses du bilan."""

    def __init__(self, parent, df=None):
        super().__init__(parent, "Ratio Bilan", df=df)

    def _masses_bilan(self, annee):
        actifs_non_courants = (
            self._sum_by_prefixes(
                annee, ["29", "290", "291", "292", "293", "296", "297"], mode="actif"
            )
            + (
                self._sum_by_prefixes(
                    annee, ["20", "203", "204", "205", "207", "208"], mode="actif"
                )
                - max(
                    0.0,
                    self._sum_by_prefixes(
                        annee, ["280", "2803", "2804", "2807", "2808"], mode="passif"
                    ),
                )
            )
            + (
                self._sum_by_prefixes(
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
                    mode="actif",
                )
                - max(
                    0.0,
                    self._sum_by_prefixes(
                        annee,
                        ["281", "2811", "2812", "2813", "2815", "2818", "282", "291"],
                        mode="passif",
                    ),
                )
            )
            + self._sum_by_prefixes(
                annee,
                ["26", "261", "262", "265", "266", "267", "268", "269"],
                mode="actif",
            )
            + self._sum_by_prefixes(
                annee,
                ["27", "271", "272", "273", "274", "275", "276", "277", "279"],
                mode="actif",
            )
        )

        stocks = self._sum_by_prefixes(
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
            mode="actif",
        )
        clients = self._sum_by_prefixes(
            annee,
            ["411", "409", "4091", "4096", "4098", "416", "417", "418"],
            mode="actif",
        )
        impot = self._sum_by_prefixes(annee, ["4456", "4487"], mode="actif")
        caisse = self._sum_by_prefixes(annee, ["53"], mode="actif")
        banque = max(0.0, self._sum_by_prefixes(annee, ["512"], mode="actif"))
        actifs_courants = stocks + clients + impot + caisse + banque

        capitaux_propres = self._sum_by_prefixes(
            annee, ["10", "11", "12"], mode="passif"
        )
        passifs_non_courants = (
            self._sum_by_prefixes(annee, ["13"], mode="passif")
            + self._sum_by_prefixes(annee, ["158"], mode="passif")
            + self._sum_by_prefixes(
                annee, ["1611", "1631", "1641", "1651", "1671", "1681"], mode="passif"
            )
        )
        passifs_courants = (
            self._sum_by_prefixes(
                annee, ["1612", "1632", "1642", "1652", "1672", "1682"], mode="passif"
            )
            + self._sum_by_prefixes(
                annee,
                ["401", "403", "404", "405", "408", "409", "4091", "4096", "4098"],
                mode="passif",
            )
            + self._sum_by_prefixes(annee, ["481", "487"], mode="passif")
            + self._sum_by_prefixes(
                annee, ["451", "455", "467", "468", "456"], mode="passif"
            )
            + max(0.0, -self._sum_by_prefixes(annee, ["512"], mode="actif"))
        )

        total_actifs = actifs_non_courants + actifs_courants
        total_dettes = passifs_non_courants + passifs_courants
        total_passifs = capitaux_propres + total_dettes

        return {
            "actifs_non_courants": actifs_non_courants,
            "stocks": stocks,
            "actifs_courants": actifs_courants,
            "capitaux_propres": capitaux_propres,
            "passifs_courants": passifs_courants,
            "total_dettes": total_dettes,
            "total_actifs": total_actifs,
            "total_passifs": total_passifs,
        }

    def _calculer_ratios(self, annee_courante, annee_precedente):
        n = self._masses_bilan(annee_courante)
        n1 = self._masses_bilan(annee_precedente)

        return [
            (
                "Liquidité générale (Actifs courants / Passifs courants)",
                self._fmt_coef(
                    self._safe_ratio(n["actifs_courants"], n["passifs_courants"])
                ),
                self._fmt_coef(
                    self._safe_ratio(n1["actifs_courants"], n1["passifs_courants"])
                ),
            ),
            (
                "Liquidité réduite ((Actifs courants - Stocks) / Passifs courants)",
                self._fmt_coef(
                    self._safe_ratio(
                        n["actifs_courants"] - n["stocks"], n["passifs_courants"]
                    )
                ),
                self._fmt_coef(
                    self._safe_ratio(
                        n1["actifs_courants"] - n1["stocks"], n1["passifs_courants"]
                    )
                ),
            ),
            (
                "Autonomie financière (Capitaux propres / Total passifs)",
                self._fmt_percent(
                    self._safe_ratio(n["capitaux_propres"], n["total_passifs"])
                ),
                self._fmt_percent(
                    self._safe_ratio(n1["capitaux_propres"], n1["total_passifs"])
                ),
            ),
            (
                "Endettement global (Total dettes / Capitaux propres)",
                self._fmt_coef(
                    self._safe_ratio(n["total_dettes"], n["capitaux_propres"])
                ),
                self._fmt_coef(
                    self._safe_ratio(n1["total_dettes"], n1["capitaux_propres"])
                ),
            ),
            (
                "Couverture des actifs non courants (Capitaux propres / Actifs non courants)",
                self._fmt_coef(
                    self._safe_ratio(n["capitaux_propres"], n["actifs_non_courants"])
                ),
                self._fmt_coef(
                    self._safe_ratio(n1["capitaux_propres"], n1["actifs_non_courants"])
                ),
            ),
            (
                "Solvabilité générale (Total actifs / Total dettes)",
                self._fmt_coef(self._safe_ratio(n["total_actifs"], n["total_dettes"])),
                self._fmt_coef(
                    self._safe_ratio(n1["total_actifs"], n1["total_dettes"])
                ),
            ),
        ]
