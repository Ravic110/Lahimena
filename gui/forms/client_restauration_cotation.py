"""
Cotation restauration par client — tableau récapitulatif par hôtel.

Colonnes : Ville | Nuits | Hôtel | Nb pax | Forfait | Repas inclus
           | Prix unitaire | Total

Les prix de repas sont récupérés depuis la base hôtels (section REPAS).
Pas de marge : Total = Prix unitaire × Nuits.
"""

import re
import tkinter as tk
import unicodedata
from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    ACCENT_TEXT_COLOR,
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_RED,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    MUTED_TEXT_COLOR,
    PANEL_BG_COLOR,
    RESTAURATIONS,
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.excel_handler import (
    load_all_hotels,
    load_client_hotel_cotation,
    load_client_restauration_cotation,
    save_client_restauration_cotation_to_excel,
)
from utils.validators import convert_currency, get_exchange_rates

# ── Constantes ────────────────────────────────────────────────────────────────

# Clés repas : (clé interne, libellé court, libellé complet)
_MEAL_TYPES = [
    ("petit_dejeuner", "PDJ",             "Petit-déjeuner"),
    ("dejeuner",       "DJ",              "Déjeuner"),
    ("diner",          "DR",              "Dîner"),
    ("repas_guide",    "Repas guide",     "Repas guide"),
    ("repas_chauffeur","Repas chauffeur", "Repas chauffeur"),
]

# Repas pouvant être marqués "Gratuit" (inclus dans le prix chambre)
_GRATUIT_ALLOWED = {"petit_dejeuner", "dejeuner", "diner"}

# Repas inclus selon le forfait
_FORFAIT_MEALS: dict = {
    "Sans restauration":   [],
    "Petit déjeuner":      ["petit_dejeuner"],
    "Demi-pension":        ["petit_dejeuner", "diner"],
    "Pension complète":    ["petit_dejeuner", "dejeuner", "diner"],
    "All inclusive soft":  ["petit_dejeuner", "dejeuner", "diner"],
    "All inclusive":       ["petit_dejeuner", "dejeuner", "diner",
                            "repas_guide", "repas_chauffeur"],
    "Ultra all inclusive": ["petit_dejeuner", "dejeuner", "diner",
                            "repas_guide", "repas_chauffeur"],
}

_HOVER_GREEN = "#0A6870"
_HOVER_BLUE  = "#0B6080"
_HOVER_RED   = "#A82020"
_HOVER_GREY  = "#9EA7AA"


# ── Utilitaires ───────────────────────────────────────────────────────────────

def _normalize(name: str) -> str:
    if not name:
        return ""
    text = unicodedata.normalize("NFKD", str(name).strip().lower())
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _to_float(s, default: float = 0.0) -> float:
    try:
        return float(str(s).replace(",", ".").strip() or default)
    except (ValueError, TypeError):
        return default


def _to_int(s, default: int = 0) -> int:
    try:
        return max(0, int(_to_float(s, default)))
    except (ValueError, TypeError):
        return default


def _fmt(value: float) -> str:
    return f"{value:,.2f}"


def _compute_prix_unitaire(meal_prices: dict) -> float:
    """Prix unitaire = Σ (nb × prix) pour les repas non gratuits."""
    total = 0.0
    for mk, _, _ in _MEAL_TYPES:
        entry = meal_prices.get(mk, {})
        if not entry.get("gratuit", False):
            total += _to_int(entry.get("count", 0)) * _to_float(entry.get("price", 0))
    return total


def _build_meal_summary(meal_prices: dict) -> str:
    """Résumé compact : '2 PDJ (grat.), 2 DR'."""
    parts = []
    for mk, short, _ in _MEAL_TYPES:
        entry = meal_prices.get(mk, {})
        n = _to_int(entry.get("count", 0))
        if n > 0:
            label = short + (" (grat.)" if entry.get("gratuit") else "")
            parts.append(f"{n} {label}")
    return ", ".join(parts) if parts else "—"


def _make_row(ville="", nuits="", hotel="", nb_pax="",
              forfait="", meal_prices=None) -> dict:
    if meal_prices is None:
        meal_prices = {mk: {"count": 0, "price": 0.0, "gratuit": False}
                       for mk, _, _ in _MEAL_TYPES}
    prix  = _compute_prix_unitaire(meal_prices)
    n     = _to_float(nuits, 1.0)
    total = prix * n
    return {
        "ville":         ville,
        "nuits":         nuits,
        "hotel":         hotel,
        "nb_pax":        nb_pax,
        "forfait":       forfait,
        "meal_prices":   meal_prices,
        "prix_unitaire": prix,
        "total":         total,
    }


# ── Classe principale ─────────────────────────────────────────────────────────

class ClientRestaurationCotation:
    """Tableau de cotation restauration par hôtel/ville pour un client donné."""

    _COLS = [
        ("ville",         "Ville",          130),
        ("nuits",         "Nuits",           50),
        ("hotel",         "Hôtel",          180),
        ("nb_pax",        "Nb pax",          55),
        ("forfait",       "Forfait",        145),
        ("repas",         "Repas inclus",   180),
        ("prix_unitaire", "Prix unitaire",  115),
        ("total",         "Total",          115),
    ]

    def __init__(self, parent: tk.Widget, client: dict, on_back=None, embedded=False):
        self.parent   = parent
        self.client   = client
        self.on_back  = on_back
        self.embedded = embedded
        self._rows: list = []

        # Charger les hôtels avec leurs données repas
        raw_hotels = load_all_hotels() or []
        self._hotel_labels: list = []
        self._hotels_by_label: dict = {}
        self._hotels_by_city: dict = {}
        seen: set = set()
        for h in raw_hotels:
            nom  = (h.get("nom") or "").strip()
            lieu = (h.get("lieu") or "").strip()
            cat  = (h.get("categorie") or "").strip()
            lbl  = f"{nom} — {lieu}" + (f" ({cat})" if cat else "")
            if lbl not in seen:
                seen.add(lbl)
                self._hotel_labels.append(lbl)
                self._hotels_by_label[lbl] = h
            norm_lieu = _normalize(lieu)
            if norm_lieu:
                self._hotels_by_city.setdefault(norm_lieu, [])
                if lbl not in self._hotels_by_city[norm_lieu]:
                    self._hotels_by_city[norm_lieu].append(lbl)

        try:
            self._rates = get_exchange_rates()
        except Exception:
            self._rates = {"EUR": 5235.0, "USD": 4900.0}

        self._build_ui()
        self._populate_initial_rows()
        self._refresh_tree()
        self._refresh_totals()

    # ── Construction de l'interface ────────────────────────────────────────────

    def _build_ui(self):
        client  = self.client
        nom     = client.get("nom", "")
        prenom  = client.get("prenom", "")
        dossier = client.get("numero_dossier", "")
        pax     = str(client.get("nombre_participants") or client.get("nombre_adultes") or "")
        sejour  = str(client.get("duree_sejour") or "")

        # ── Mode page complète (non embarqué) ─────────────────────────────────
        if not self.embedded:
            tk.Label(
                self.parent,
                text="COTATION RESTAURATION",
                font=TITLE_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR,
            ).pack(pady=(20, 4))

            client_name = f"{prenom} {nom}".strip() or "—"
            tk.Label(
                self.parent,
                text="Client : " + client_name + (f"   |   Dossier : {dossier}" if dossier else ""),
                font=ENTRY_FONT, fg=MUTED_TEXT_COLOR, bg=MAIN_BG_COLOR,
            ).pack(pady=(0, 2))

            if self.on_back:
                back_bar = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
                back_bar.pack(fill="x", padx=20, pady=(4, 0))
                ctk.CTkButton(
                    back_bar, text="⬅  Retour à l'accueil",
                    command=self.on_back,
                    fg_color=BUTTON_BLUE, hover_color=_HOVER_BLUE, text_color="white",
                    font=BUTTON_FONT, corner_radius=8, cursor="hand2",
                ).pack(side="left")

            root = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
            root.pack(fill="both", expand=True, padx=20, pady=(8, 20))

            # Carte Infos client
            info_card = tk.Frame(root, bg=PANEL_BG_COLOR)
            info_card.pack(fill="x", pady=(0, 10))
            tk.Label(
                info_card, text="Informations client",
                font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
            ).pack(anchor="w", padx=12, pady=(8, 2))
            info_inner = tk.Frame(info_card, bg=PANEL_BG_COLOR)
            info_inner.pack(fill="x", padx=12, pady=(0, 8))
            for col_i, (lbl_text, val_text) in enumerate([
                ("Participants :", pax or "—"),
                ("Durée séjour :", f"{sejour} j" if sejour else "—"),
                ("Formule :", str(client.get("restauration") or "—")),
            ]):
                tk.Label(info_inner, text=lbl_text,
                         font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
                         ).grid(row=0, column=col_i * 2,
                                sticky="w", padx=(0 if col_i == 0 else 20, 4), pady=4)
                tk.Label(info_inner, text=val_text,
                         font=ENTRY_FONT, fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR,
                         ).grid(row=0, column=col_i * 2 + 1, sticky="w", pady=4)

            container = root
            btn_bg    = MAIN_BG_COLOR
        else:
            # ── Mode embarqué ──────────────────────────────────────────────────
            container = self.parent
            btn_bg    = PANEL_BG_COLOR

        # ── Barre d'actions ────────────────────────────────────────────────────
        action_bar = tk.Frame(container, bg=btn_bg)
        action_bar.pack(fill="x", pady=(0, 6))

        ctk.CTkButton(
            action_bar, text="＋  Ajouter une ligne",
            command=self._add_row_dialog,
            fg_color=BUTTON_GREEN, hover_color=_HOVER_GREEN, text_color="white",
            font=BUTTON_FONT, corner_radius=8, cursor="hand2",
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="✏️  Modifier la ligne",
            command=self._edit_selected,
            fg_color=BUTTON_BLUE, hover_color=_HOVER_BLUE, text_color="white",
            font=BUTTON_FONT, corner_radius=8, cursor="hand2",
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="🗑️  Supprimer la ligne",
            command=self._delete_selected,
            fg_color=BUTTON_RED, hover_color=_HOVER_RED, text_color="white",
            font=BUTTON_FONT, corner_radius=8, cursor="hand2",
        ).pack(side="left")

        ctk.CTkButton(
            action_bar, text="💾  Sauvegarder",
            command=self._save_to_excel,
            fg_color=BUTTON_GREEN, hover_color=_HOVER_GREEN, text_color="white",
            font=BUTTON_FONT, corner_radius=8, cursor="hand2",
        ).pack(side="right")

        # ── Tableau ────────────────────────────────────────────────────────────
        table_outer = tk.Frame(container, bg=PANEL_BG_COLOR)
        table_outer.pack(fill="both", expand=True, pady=(0, 6))
        table_inner = tk.Frame(table_outer, bg=PANEL_BG_COLOR)
        table_inner.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure(
            "Rest.Treeview",
            background=INPUT_BG_COLOR, foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR, rowheight=30,
        )
        style.configure("Rest.Treeview.Heading", font=LABEL_FONT)
        style.map("Rest.Treeview", background=[("selected", BUTTON_BLUE)])

        col_ids = [c[0] for c in self._COLS]
        self._tree = ttk.Treeview(
            table_inner, columns=col_ids, show="headings",
            height=10, style="Rest.Treeview",
        )
        for key, heading, width in self._COLS:
            self._tree.heading(key, text=heading)
            anchor = "e" if key in ("nuits", "nb_pax", "prix_unitaire", "total") else "w"
            self._tree.column(key, width=width, anchor=anchor, stretch=False)

        vsb = ttk.Scrollbar(table_inner, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(table_inner, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_inner.grid_rowconfigure(0, weight=1)
        table_inner.grid_columnconfigure(0, weight=1)
        self._tree.bind("<Double-1>", lambda e: self._edit_selected())
        self._tree.tag_configure("odd",  background="#E4F2F6")
        self._tree.tag_configure("even", background=INPUT_BG_COLOR)

        # ── Totaux ─────────────────────────────────────────────────────────────
        totals_card = tk.Frame(container, bg=PANEL_BG_COLOR)
        totals_card.pack(fill="x")
        totals_inner = tk.Frame(totals_card, bg=PANEL_BG_COLOR)
        totals_inner.pack(fill="x", padx=12, pady=(4, 8))

        tk.Label(
            totals_inner, text="Total restauration :",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)

        self._lbl_total = tk.Label(
            totals_inner, text="0.00",
            font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR,
        )
        self._lbl_total.grid(row=0, column=1, sticky="w", pady=4)

    # ── Données initiales ──────────────────────────────────────────────────────

    def _populate_initial_rows(self):
        saved = load_client_restauration_cotation(self.client)
        if saved:
            self._rows = saved
            return

        hotel_rows = load_client_hotel_cotation(self.client)
        pax = str(self.client.get("nombre_participants") or
                  self.client.get("nombre_adultes") or "")
        forfait_client = str(self.client.get("restauration") or "")

        if hotel_rows:
            for hr in hotel_rows:
                hotel_lbl  = hr.get("hotel", "")
                meal_prices = self._default_meal_prices(hotel_lbl, pax, forfait_client)
                self._rows.append(_make_row(
                    ville       = hr.get("ville", ""),
                    nuits       = hr.get("nuits", ""),
                    hotel       = hotel_lbl,
                    nb_pax      = hr.get("nb_pax", pax),
                    forfait     = forfait_client,
                    meal_prices = meal_prices,
                ))
        else:
            self._rows.append(_make_row(nb_pax=pax, forfait=forfait_client))

    def _default_meal_prices(self, hotel_label: str, pax: str, forfait: str) -> dict:
        """Construit meal_prices depuis la BD hôtels selon le forfait."""
        h = self._hotels_by_label.get(hotel_label)
        pax_count = _to_int(pax, 0)
        included  = _FORFAIT_MEALS.get(forfait, [])
        mp = {}
        for mk, _, _ in _MEAL_TYPES:
            price = 0.0
            if h:
                raw_price = h.get("meals", {}).get(mk, 0) or 0
                unite = (h.get("unite") or "MGA").strip().upper()
                if raw_price and unite not in ("MGA", "ARIARY", "AR", ""):
                    try:
                        raw_price = convert_currency(float(raw_price), unite, "MGA", self._rates)
                    except Exception:
                        pass
                price = float(raw_price)
            count = pax_count if mk in included else 0
            mp[mk] = {"count": count, "price": price, "gratuit": False}
        return mp

    def _hotels_for_city(self, ville: str) -> list:
        norm = _normalize(ville)
        if not norm:
            return self._hotel_labels
        filtered = self._hotels_by_city.get(norm, [])
        return filtered if filtered else self._hotel_labels

    def _hotel_currency(self, hotel_label: str) -> str:
        h = self._hotels_by_label.get(hotel_label)
        if not h:
            return "MGA"
        return (h.get("unite") or "MGA").strip().upper() or "MGA"

    def _get_hotel_meal_prices(self, hotel_label: str) -> dict:
        """Retourne {meal_key: price_in_MGA} depuis la BD hôtel."""
        h = self._hotels_by_label.get(hotel_label)
        if not h:
            return {}
        unite = self._hotel_currency(hotel_label)
        prices = {}
        for mk, _, _ in _MEAL_TYPES:
            raw = h.get("meals", {}).get(mk, 0) or 0
            if raw and unite not in ("MGA", "ARIARY", "AR", ""):
                try:
                    raw = convert_currency(float(raw), unite, "MGA", self._rates)
                except Exception:
                    pass
            prices[mk] = float(raw)
        return prices

    # ── Tableau ────────────────────────────────────────────────────────────────

    def _refresh_tree(self):
        self._tree.delete(*self._tree.get_children())
        for i, rd in enumerate(self._rows):
            repas_str = _build_meal_summary(rd["meal_prices"])
            prix  = rd["prix_unitaire"]
            total = rd["total"]
            tag   = "odd" if i % 2 else "even"
            self._tree.insert("", "end", iid=str(i), values=(
                rd["ville"],
                rd["nuits"],
                rd["hotel"],
                rd["nb_pax"],
                rd.get("forfait", ""),
                repas_str,
                _fmt(prix)  if prix  else "",
                _fmt(total) if total else "",
            ), tags=(tag,))

    def _refresh_totals(self):
        self._lbl_total.configure(text=_fmt(sum(r["total"] for r in self._rows)))

    # ── Sauvegarde Excel ───────────────────────────────────────────────────────

    def _save_to_excel(self):
        if not self._rows:
            messagebox.showwarning("Aucune donnée", "Le tableau est vide. Rien à sauvegarder.")
            return
        result = save_client_restauration_cotation_to_excel(self.client, self._rows)
        if result > 0:
            messagebox.showinfo(
                "Sauvegarde réussie",
                f"{result} ligne(s) enregistrée(s) dans la base de données.",
            )
        elif result == -2:
            messagebox.showerror(
                "Fichier verrouillé",
                "Le fichier Excel est ouvert ailleurs.\n"
                "Fermez data.xlsx puis réessayez.",
            )
        else:
            messagebox.showerror(
                "Erreur",
                "La sauvegarde a échoué. Consultez les logs pour plus de détails.",
            )

    # ── Actions ────────────────────────────────────────────────────────────────

    def _add_row_dialog(self):
        pax     = str(self.client.get("nombre_participants") or
                      self.client.get("nombre_adultes") or "")
        forfait = str(self.client.get("restauration") or "")
        self._open_row_dialog(_make_row(nb_pax=pax, forfait=forfait), row_index=None)

    def _edit_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Aucune sélection", "Sélectionnez une ligne à modifier.")
            return
        self._open_row_dialog(self._rows[int(sel[0])], row_index=int(sel[0]))

    def _delete_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Aucune sélection", "Sélectionnez une ligne à supprimer.")
            return
        if not messagebox.askyesno("Supprimer", "Supprimer la ligne sélectionnée ?"):
            return
        del self._rows[int(sel[0])]
        self._refresh_tree()
        self._refresh_totals()

    # ── Dialog d'édition ──────────────────────────────────────────────────────

    def _open_row_dialog(self, row: dict, row_index):
        win = tk.Toplevel(self.parent)
        win.title("Modifier la ligne" if row_index is not None else "Ajouter une ligne")
        win.configure(bg=MAIN_BG_COLOR)
        win.resizable(False, False)
        win.transient(self.parent)
        win.after(0, lambda: [win.lift(), win.focus_set()])

        # ── Variables ──────────────────────────────────────────────────────────
        v_ville   = tk.StringVar(value=row["ville"])
        v_nuits   = tk.StringVar(value=row["nuits"])
        v_hotel   = tk.StringVar(value=row["hotel"])
        v_pax     = tk.StringVar(value=row["nb_pax"])
        v_forfait = tk.StringVar(value=row.get("forfait", ""))

        local_mp: dict = {
            mk: {
                "count":  _to_int(row["meal_prices"].get(mk, {}).get("count", 0)),
                "price":  _to_float(row["meal_prices"].get(mk, {}).get("price", 0)),
                "gratuit": bool(row["meal_prices"].get(mk, {}).get("gratuit", False)),
            }
            for mk, _, _ in _MEAL_TYPES
        }

        mp_vars: dict = {
            mk: {
                "count":  tk.StringVar(value=str(local_mp[mk]["count"])),
                "price":  tk.StringVar(value=str(local_mp[mk]["price"]) if local_mp[mk]["price"] else ""),
                "gratuit": tk.BooleanVar(value=local_mp[mk]["gratuit"]),
            }
            for mk, _, _ in _MEAL_TYPES
        }

        # ── Layout ────────────────────────────────────────────────────────────
        outer = tk.Frame(win, bg=MAIN_BG_COLOR)
        outer.pack(padx=24, pady=(20, 0))

        SEC = PANEL_BG_COLOR

        def _lbl(parent, text, r, c=0, span=1, bg=None):
            tk.Label(
                parent, text=text,
                font=LABEL_FONT, fg=TEXT_COLOR, bg=bg or SEC, anchor="w",
            ).grid(row=r, column=c, columnspan=span, sticky="w", padx=(0, 10), pady=5)

        def _entry(parent, var, r, c=1, w=28, justify="left"):
            e = tk.Entry(
                parent, textvariable=var,
                font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                width=w, justify=justify,
                insertbackground=TEXT_COLOR, relief="flat",
            )
            e.grid(row=r, column=c, sticky="ew", pady=5)
            return e

        def _make_section(title):
            card = tk.Frame(outer, bg=SEC)
            card.pack(fill="x", pady=(0, 10))
            tk.Label(card, text=title, font=LABEL_FONT, fg=TEXT_COLOR, bg=SEC,
                     ).pack(anchor="w", padx=10, pady=(8, 2))
            inner = tk.Frame(card, bg=SEC)
            inner.pack(fill="x", padx=10, pady=(0, 8))
            return inner

        # ── Section 1 : Lieu, Hôtel & Forfait ─────────────────────────────────
        s1 = _make_section("Lieu, Hôtel & Forfait")

        _lbl(s1, "Ville :", 0)
        _entry(s1, v_ville, 0, w=28)

        _lbl(s1, "Nuits :", 1)
        _entry(s1, v_nuits, 1, w=6, justify="center")

        _lbl(s1, "Hôtel :", 2)
        combo_hotel = ttk.Combobox(
            s1, textvariable=v_hotel,
            values=self._hotels_for_city(row["ville"]),
            font=ENTRY_FONT, width=33, state="normal",
        )
        combo_hotel.grid(row=2, column=1, sticky="ew", pady=5)

        lbl_devise = tk.Label(s1, text="", font=ENTRY_FONT,
                              fg=ACCENT_TEXT_COLOR, bg=SEC, anchor="w")
        lbl_devise.grid(row=2, column=2, sticky="w", padx=(8, 0), pady=5)

        _lbl(s1, "Nb pax :", 3)
        _entry(s1, v_pax, 3, w=6, justify="center")

        _lbl(s1, "Forfait :", 4)
        combo_forfait = ttk.Combobox(
            s1, textvariable=v_forfait,
            values=[""] + list(RESTAURATIONS),
            font=ENTRY_FONT, width=28, state="readonly",
        )
        combo_forfait.grid(row=4, column=1, sticky="ew", pady=5)

        # ── Section 2 : Repas ─────────────────────────────────────────────────
        s2 = _make_section("Repas — détail par type")

        # En-têtes
        for c_i, (hdr, w) in enumerate([
            ("Type", 16), ("Nb personnes", 13), ("Prix / repas", 15),
            ("Sous-total", 12), ("Gratuit", 7),
        ]):
            tk.Label(
                s2, text=hdr,
                font=LABEL_FONT, fg=MUTED_TEXT_COLOR, bg=SEC,
                width=w, anchor="center",
            ).grid(row=0, column=c_i, padx=4, pady=(0, 2))

        tk.Frame(s2, bg=MUTED_TEXT_COLOR, height=1).grid(
            row=1, column=0, columnspan=5, sticky="ew", padx=2, pady=(0, 4),
        )

        subtotal_labels: dict = {}
        price_entries:   dict = {}

        def _update_subtotal(mk, *_):
            entry = mp_vars[mk]
            count   = _to_int(entry["count"].get())
            price   = _to_float(entry["price"].get())
            gratuit = entry["gratuit"].get()
            sub = 0.0 if gratuit else count * price
            if mk in subtotal_labels:
                subtotal_labels[mk].configure(
                    text=_fmt(sub) if (sub or gratuit) else "0.00",
                    fg=MUTED_TEXT_COLOR if gratuit else (
                        ACCENT_TEXT_COLOR if sub > 0 else MUTED_TEXT_COLOR),
                )
            # Griser le champ prix si gratuit
            if mk in price_entries:
                price_entries[mk].configure(
                    state="disabled" if gratuit else "normal",
                    bg="#D0D0D0" if gratuit else INPUT_BG_COLOR,
                )
            _update_preview()

        for row_i, (mk, short, full) in enumerate(_MEAL_TYPES):
            grid_r      = row_i + 2
            can_gratuit = mk in _GRATUIT_ALLOWED

            tk.Label(s2, text=full, font=LABEL_FONT, fg=TEXT_COLOR,
                     bg=SEC, width=16, anchor="w",
                     ).grid(row=grid_r, column=0, padx=4, pady=3)

            e_count = tk.Entry(
                s2, textvariable=mp_vars[mk]["count"],
                font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                width=13, justify="center", relief="flat",
                insertbackground=TEXT_COLOR,
            )
            e_count.grid(row=grid_r, column=1, padx=4, pady=3)

            e_price = tk.Entry(
                s2, textvariable=mp_vars[mk]["price"],
                font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                width=15, justify="right", relief="flat",
                insertbackground=TEXT_COLOR,
            )
            e_price.grid(row=grid_r, column=2, padx=4, pady=3)
            price_entries[mk] = e_price

            lbl_sub = tk.Label(s2, text="0.00", font=ENTRY_FONT,
                               fg=MUTED_TEXT_COLOR, bg=SEC, width=12, anchor="e")
            lbl_sub.grid(row=grid_r, column=3, padx=4, pady=3)
            subtotal_labels[mk] = lbl_sub

            if can_gratuit:
                chk = tk.Checkbutton(
                    s2, text="",
                    variable=mp_vars[mk]["gratuit"],
                    bg=SEC, activebackground=SEC,
                    selectcolor=BUTTON_GREEN,
                    command=lambda m=mk: _update_subtotal(m),
                )
                chk.grid(row=grid_r, column=4, padx=4, pady=3)
            else:
                tk.Label(s2, text="", bg=SEC, width=7).grid(
                    row=grid_r, column=4, padx=4, pady=3)

            mp_vars[mk]["count"].trace_add("write", lambda *_, m=mk: _update_subtotal(m))
            mp_vars[mk]["price"].trace_add("write", lambda *_, m=mk: _update_subtotal(m))

        # Séparateur + prix total
        sep_r = len(_MEAL_TYPES) + 2
        tk.Frame(s2, bg=MUTED_TEXT_COLOR, height=1).grid(
            row=sep_r, column=0, columnspan=5, sticky="ew", padx=2, pady=(4, 2),
        )
        tk.Label(s2, text="Prix unitaire total :",
                 font=LABEL_FONT, fg=TEXT_COLOR, bg=SEC,
                 ).grid(row=sep_r + 1, column=2, columnspan=2, sticky="e", padx=(0, 4), pady=4)

        lbl_prix_total = tk.Label(s2, text="0.00",
                                  font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=SEC,
                                  width=12, anchor="e")
        lbl_prix_total.grid(row=sep_r + 1, column=4, padx=4, pady=4)

        # ── Preview total ──────────────────────────────────────────────────────
        s3 = tk.Frame(outer, bg=MAIN_BG_COLOR)
        s3.pack(fill="x", pady=(0, 8))

        lbl_prev = tk.Label(
            s3, text="Total : —",
            font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=MAIN_BG_COLOR,
        )
        lbl_prev.pack(anchor="w")

        def _update_preview(*_):
            prix  = sum(
                (0.0 if mp_vars[mk]["gratuit"].get() else
                 _to_int(mp_vars[mk]["count"].get()) * _to_float(mp_vars[mk]["price"].get()))
                for mk, _, _ in _MEAL_TYPES
            )
            nuits = _to_float(v_nuits.get(), 1.0)
            total = prix * nuits
            lbl_prix_total.configure(text=_fmt(prix))
            lbl_prev.configure(
                text=f"Prix unitaire : {_fmt(prix)}   |   Total : {_fmt(total)}"
            )

        v_nuits.trace_add("write", _update_preview)

        # ── Auto-remplissage depuis le forfait ─────────────────────────────────
        def _apply_forfait(*_):
            forfait  = v_forfait.get()
            included = _FORFAIT_MEALS.get(forfait, None)
            if included is None:
                return
            pax_count = str(_to_int(v_pax.get(), 0))
            for mk, _, _ in _MEAL_TYPES:
                mp_vars[mk]["count"].set(pax_count if mk in included else "0")
            _update_preview()

        combo_forfait.bind("<<ComboboxSelected>>", _apply_forfait)
        v_pax.trace_add("write", _update_preview)

        # ── Filtre hôtel par ville ─────────────────────────────────────────────
        def _update_hotel_filter(*_):
            filtered = self._hotels_for_city(v_ville.get())
            combo_hotel["values"] = filtered
            if v_hotel.get() and v_hotel.get() not in filtered:
                v_hotel.set("")
                lbl_devise.configure(text="")

        v_ville.trace_add("write", _update_hotel_filter)

        # ── Auto-remplissage prix depuis la BD ────────────────────────────────
        def _autofill_prices(*_):
            hotel_lbl = v_hotel.get().strip()
            unite = self._hotel_currency(hotel_lbl) if hotel_lbl else "MGA"
            if unite not in ("MGA", "ARIARY", "AR", ""):
                rate = self._rates.get(unite)
                lbl_devise.configure(
                    text=(f"⚠ {unite} → MGA  (1 {unite} ≈ {rate:,.0f} Ar)" if rate
                          else f"⚠ {unite} → MGA")
                )
            else:
                lbl_devise.configure(text="")

            prices = self._get_hotel_meal_prices(hotel_lbl)
            for mk, _, _ in _MEAL_TYPES:
                db_price = prices.get(mk)
                if db_price:
                    mp_vars[mk]["price"].set(f"{float(db_price):.2f}")
            _update_preview()

        combo_hotel.bind("<<ComboboxSelected>>", _autofill_prices)

        # Initialisation
        for mk, _, _ in _MEAL_TYPES:
            _update_subtotal(mk)
        _update_preview()

        # ── Boutons Valider / Annuler ──────────────────────────────────────────
        btn_bar = tk.Frame(win, bg=MAIN_BG_COLOR)
        btn_bar.pack(pady=(8, 20), padx=24, fill="x")

        def _save():
            new_mp = {
                mk: {
                    "count":  _to_int(mp_vars[mk]["count"].get()),
                    "price":  _to_float(mp_vars[mk]["price"].get()),
                    "gratuit": mp_vars[mk]["gratuit"].get(),
                }
                for mk, _, _ in _MEAL_TYPES
            }
            new_row = _make_row(
                ville       = v_ville.get().strip(),
                nuits       = v_nuits.get().strip(),
                hotel       = v_hotel.get().strip(),
                nb_pax      = v_pax.get().strip(),
                forfait     = v_forfait.get().strip(),
                meal_prices = new_mp,
            )
            if row_index is None:
                self._rows.append(new_row)
            else:
                self._rows[row_index] = new_row
            self._refresh_tree()
            self._refresh_totals()
            win.destroy()

        ctk.CTkButton(
            btn_bar, text="✔  Valider",
            command=_save,
            fg_color=BUTTON_GREEN, hover_color=_HOVER_GREEN, text_color="white",
            font=BUTTON_FONT, corner_radius=8, cursor="hand2",
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_bar, text="Annuler",
            command=win.destroy,
            fg_color="#9EA7AA", hover_color=_HOVER_GREY, text_color="white",
            font=BUTTON_FONT, corner_radius=8, cursor="hand2",
        ).pack(side="left")

        # Centrer
        win.update_idletasks()
        pw = self.parent.winfo_rootx() + self.parent.winfo_width() // 2
        ph = self.parent.winfo_rooty() + self.parent.winfo_height() // 2
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        win.geometry(f"{w}x{h}+{pw - w // 2}+{ph - h // 2}")
