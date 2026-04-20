"""
Cotation hôtelière par client — tableau récapitulatif par ville.

Colonnes : Ville | Nuits | Hôtel | Rooming list | Nb pax
           | Prix unitaire | Dépense | Marge (%) | Total

Le prix unitaire est calculé depuis la base hôtels :
  prix_unitaire = Σ (nb_chambres_type × prix_type_depuis_BD)
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
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.excel_handler import (
    load_all_hotels,
    load_client_hotel_cotation,
    save_client_hotel_cotation_to_excel,
)
from utils.validators import convert_currency, get_exchange_rates

# ── Constantes ────────────────────────────────────────────────────────────────

# Groupes de chambre (clé BD → libellé affiché)
_GROUPS = [
    ("standard",  "Standard"),
    ("bungalows", "Bungalows"),
    ("deluxe",    "De luxe"),
    ("suite",     "Suite"),
]
_GROUP_LABELS  = [g[1] for g in _GROUPS]
_GROUP_BY_LBL  = {g[1]: g[0] for g in _GROUPS}
_GROUP_BY_KEY  = {g[0]: g[1] for g in _GROUPS}

# Types de chambre du rooming client → clé room_rates
_ROOM_TYPES = [
    ("sgl_count", "single",    "SGL"),
    ("dbl_count", "double",    "DBL"),
    ("twn_count", "twin",      "TWN"),
    ("tpl_count", "triple",    "TPL"),
    ("fml_count", "familiale", "FML"),
]

# Couleurs hover pour les boutons CTk
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


def _parse_cities(raw: str) -> list:
    if not raw:
        return []
    seen, result = set(), []
    for city in re.split(r"[,;>/|\n]+", str(raw)):
        city = city.strip()
        if " - " in city:
            city = city.split(" - ", 1)[1].strip()
        city = re.sub(r"\s+", " ", city).strip()
        norm = _normalize(city)
        if city and norm and norm not in seen:
            seen.add(norm)
            result.append(city)
    return result


def _parse_city_days(raw: str) -> dict:
    result = {}
    if not raw:
        return result
    for segment in re.split(r"[,;>/|\n]+", str(raw)):
        segment = segment.strip()
        if not segment:
            continue
        day_m = re.search(r"(\d+)\s*(?:j|jour|jours)\b", segment, re.IGNORECASE)
        city_part = segment.split(" - ", 1)[1] if " - " in segment else segment
        city_part = re.sub(r"\d+\s*(?:j|jour|jours)\b", "", city_part, flags=re.IGNORECASE)
        city_part = re.sub(r"\([^)]*\)", " ", city_part)
        city_part = re.sub(r"\s+", " ", city_part).strip()
        norm = _normalize(city_part)
        if norm and day_m:
            result[norm] = int(day_m.group(1))
    return result


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


def _client_group_key(client: dict) -> str:
    """Déduit le groupe de chambre depuis le champ 'chambre' du client."""
    raw = str(client.get("chambre") or "").strip().lower()
    if "bungalow" in raw:
        return "bungalows"
    if "deluxe" in raw or "luxe" in raw:
        return "deluxe"
    if "suite" in raw:
        return "suite"
    return "standard"


def _build_rooming_str(room_prices: dict) -> str:
    parts = []
    for _, rk, lbl in _ROOM_TYPES:
        n = _to_int(room_prices.get(rk, {}).get("count", 0))
        if n > 0:
            parts.append(f"{n} {lbl}")
    return ", ".join(parts) if parts else "—"


def _compute_prix_unitaire(room_prices: dict) -> float:
    total = 0.0
    for _, rk, _ in _ROOM_TYPES:
        entry = room_prices.get(rk, {})
        total += _to_int(entry.get("count", 0)) * _to_float(entry.get("price", 0))
    return total


# ── Classe principale ─────────────────────────────────────────────────────────

class ClientHotelCotation:
    """Tableau de cotation hôtelière par ville pour un client donné."""

    _COLS = [
        ("ville",         "Ville",          160),
        ("nuits",         "Nuits",           55),
        ("hotel",         "Hôtel",          200),
        ("rooming",       "Rooming list",   150),
        ("nb_pax",        "Nb pax",          60),
        ("prix_unitaire", "Prix unitaire",  120),
        ("depense",       "Dépense",        120),
        ("marge",         "Marge (%)",       80),
        ("total",         "Total",          120),
    ]

    def __init__(self, parent: tk.Widget, client: dict, on_back=None):
        self.parent  = parent
        self.client  = client
        self.on_back = on_back
        self._rows: list = []

        # Charger les hôtels avec leurs prix complets
        raw_hotels = load_all_hotels() or []
        self._hotel_labels: list = []
        self._hotels_by_label: dict = {}   # label → dict complet avec room_rates
        self._hotels_by_city: dict = {}    # normalized_city → [label, ...]
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

        # Taux de change (récupérés une seule fois, avec fallback)
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
        from gui.ui_style import card_frame, setup_treeview_style, create_card as _create_card

        client  = self.client
        nom     = client.get("nom", "")
        prenom  = client.get("prenom", "")
        dossier = client.get("numero_dossier", "")
        pax     = str(client.get("nombre_participants") or client.get("nombre_adultes") or "")
        sejour  = str(client.get("duree_sejour") or "")
        rooming = _build_rooming_str(
            {rk: {"count": _to_int(client.get(ck, 0)), "price": 0}
             for ck, rk, _ in _ROOM_TYPES}
        )
        client_name = f"{prenom} {nom}".strip() or "—"

        root = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        root.pack(fill="both", expand=True, padx=16, pady=12)

        # ── En-tête ────────────────────────────────────────────────────────
        _, hdr = card_frame(root, pady=(0, 8))
        hdr_top = tk.Frame(hdr, bg=PANEL_BG_COLOR)
        hdr_top.pack(fill="x")

        if self.on_back:
            ctk.CTkButton(
                hdr_top, text="← Retour",
                command=self.on_back,
                fg_color=BUTTON_BLUE, hover_color=_HOVER_BLUE,
                text_color="white", font=("Poppins", 10, "bold"),
                corner_radius=8, cursor="hand2", width=100, height=30,
            ).pack(side="left", padx=(0, 12))

        tk.Label(
            hdr_top, text="Cotation Hôtelière",
            font=TITLE_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(side="left")

        info_right = tk.Frame(hdr_top, bg=PANEL_BG_COLOR)
        info_right.pack(side="right")
        for lbl, val in [("Client", client_name), ("Dossier", dossier or "—")]:
            tk.Label(info_right, text=f"{lbl} : ", font=LABEL_FONT,
                     fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left")
            tk.Label(info_right, text=val, font=LABEL_FONT,
                     fg=TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left", padx=(0, 16))

        info_row = tk.Frame(hdr, bg=PANEL_BG_COLOR)
        info_row.pack(fill="x", pady=(6, 0))
        for lbl, val in [("Participants", pax or "—"),
                          ("Durée séjour", f"{sejour} j" if sejour else "—"),
                          ("Rooming", rooming)]:
            tk.Label(info_row, text=f"{lbl} : ", font=LABEL_FONT,
                     fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left")
            tk.Label(info_row, text=val, font=ENTRY_FONT,
                     fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left", padx=(0, 24))

        # ── Système d'onglets ──────────────────────────────────────────────
        _active_panel = [None]
        _panels       = [None, None]   # [hotel_panel, rest_panel]

        def _switch_tab(tab_name):
            if _active_panel[0]:
                _active_panel[0].pack_forget()
            idx = {"Hébergement par ville": 0, "Restauration": 1}.get(tab_name, -1)
            if idx >= 0 and _panels[idx] is not None:
                _panels[idx].pack(fill="both", expand=True)
                _active_panel[0] = _panels[idx]

        cotation_card = _create_card(
            root,
            title=None,
            tabs=[("Hébergement par ville", True), ("Restauration", False)],
            show_controls=False,
            expand=True,
            on_tab_click=_switch_tab,
        )

        # ── Panneau : Hébergement par ville ───────────────────────────────────
        hotel_panel = tk.Frame(cotation_card, bg=PANEL_BG_COLOR)
        _panels[0]       = hotel_panel
        _active_panel[0] = hotel_panel
        hotel_panel.pack(fill="both", expand=True)

        # Barre d'actions hôtel
        action_bar = tk.Frame(hotel_panel, bg=PANEL_BG_COLOR)
        action_bar.pack(fill="x", pady=(0, 6))

        for text, cmd, color, hover in [
            ("＋ Ajouter une ligne",   self._add_row_dialog,  BUTTON_GREEN, _HOVER_GREEN),
            ("✏️ Modifier la ligne",   self._edit_selected,   BUTTON_BLUE,  _HOVER_BLUE),
            ("🗑️ Supprimer la ligne", self._delete_selected,  BUTTON_RED,   _HOVER_RED),
        ]:
            ctk.CTkButton(
                action_bar, text=text, command=cmd,
                fg_color=color, hover_color=hover,
                text_color="white", font=BUTTON_FONT,
                corner_radius=8, cursor="hand2", height=32,
            ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="💾 Sauvegarder",
            command=self._save_to_excel,
            fg_color=BUTTON_GREEN, hover_color=_HOVER_GREEN,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2", height=32,
        ).pack(side="right")

        # Treeview hôtel
        setup_treeview_style("Hotel.Treeview")
        col_ids = [c[0] for c in self._COLS]
        tree_outer = tk.Frame(hotel_panel, bg=PANEL_BG_COLOR)
        tree_outer.pack(fill="both", expand=True, pady=(0, 6))

        self._tree = ttk.Treeview(
            tree_outer, columns=col_ids, show="headings",
            height=10, style="Hotel.Treeview",
        )
        for key, heading, width in self._COLS:
            self._tree.heading(key, text=heading)
            anchor = "e" if key in (
                "nuits", "nb_pax", "prix_unitaire", "depense", "marge", "total"
            ) else "w"
            self._tree.column(key, width=width, anchor=anchor, stretch=False)

        vsb = ctk.CTkScrollbar(tree_outer, orientation="vertical",   command=self._tree.yview)
        hsb = ctk.CTkScrollbar(tree_outer, orientation="horizontal",  command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self._tree.bind("<Double-1>", lambda e: self._edit_selected())
        self._tree.tag_configure("odd",  background="#E4F2F6")
        self._tree.tag_configure("even", background=INPUT_BG_COLOR)

        # Totaux hôtel
        totals_inner = tk.Frame(hotel_panel, bg=PANEL_BG_COLOR)
        totals_inner.pack(fill="x", pady=(4, 6))

        def _total_item(parent, label):
            tk.Label(parent, text=label, font=LABEL_FONT,
                     fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left")
            lbl = tk.Label(parent, text="0.00", font=("Poppins", 12, "bold"),
                           fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR)
            lbl.pack(side="left", padx=(2, 0))
            tk.Label(parent, text=" Ar  ", font=("Poppins", 10),
                     fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR).pack(side="left")
            return lbl

        self._lbl_dep  = _total_item(totals_inner, "Total dépenses : ")
        tk.Frame(totals_inner, width=1, bg="#C9DDE3").pack(side="left", fill="y", padx=8)
        self._lbl_glob = _total_item(totals_inner, "Total global (avec marges) : ")

        # ── Panneau : Restauration ─────────────────────────────────────────────
        rest_panel = tk.Frame(cotation_card, bg=PANEL_BG_COLOR)
        _panels[1] = rest_panel
        # Non empaqueté initialement (masqué)

        from gui.forms.client_restauration_cotation import ClientRestaurationCotation
        ClientRestaurationCotation(rest_panel, client=self.client, embedded=True)


    # ── Données initiales ──────────────────────────────────────────────────────

    def _populate_initial_rows(self):
        # Priorité : données déjà sauvegardées en base
        saved = load_client_hotel_cotation(self.client)
        if saved:
            self._rows = saved
            return

        # Sinon : générer depuis l'itinéraire du client
        client   = self.client
        raw_va   = client.get("ville_arrivee", "") or ""
        raw_itin = client.get("itineraire_circuit", "") or ""
        raw_dep  = (client.get("ville_depart") or "").strip()

        cities = _parse_cities(raw_va) or _parse_cities(raw_itin)
        days_m = _parse_city_days(raw_va) or _parse_city_days(raw_itin)

        all_cities: list = []
        seen: set = set()
        for c in ([raw_dep] if raw_dep else []) + cities:
            if c and c not in seen:
                seen.add(c)
                all_cities.append(c)

        pax         = str(client.get("nombre_participants") or
                          client.get("nombre_adultes") or "")
        hotel_group = _client_group_key(client)
        room_prices = self._default_room_prices(client)

        if all_cities:
            for city in all_cities:
                nuits = str(days_m.get(_normalize(city), ""))
                self._rows.append(self._make_row(
                    ville=city, nuits=nuits, nb_pax=pax,
                    hotel_group=hotel_group, room_prices=room_prices,
                ))
        else:
            self._rows.append(self._make_row(
                nb_pax=pax, hotel_group=hotel_group, room_prices=room_prices,
            ))

    def _default_room_prices(self, client: dict) -> dict:
        """Construit le dict room_prices depuis les compteurs du client (prix = 0)."""
        rp = {}
        for ck, rk, _ in _ROOM_TYPES:
            rp[rk] = {
                "count": _to_int(client.get(ck, 0)),
                "price": 0.0,
            }
        return rp

    @staticmethod
    def _make_row(ville="", nuits="", hotel="", hotel_group="standard",
                  room_prices=None, nb_pax="", marge="") -> dict:
        if room_prices is None:
            room_prices = {rk: {"count": 0, "price": 0.0} for _, rk, _ in _ROOM_TYPES}
        prix  = _compute_prix_unitaire(room_prices)
        n     = _to_float(nuits, 1.0)
        mg    = _to_float(marge)
        dep   = prix * n
        total = dep * (1 + mg / 100)
        return {
            "ville":       ville,
            "nuits":       nuits,
            "hotel":       hotel,
            "hotel_group": hotel_group,
            "room_prices": room_prices,
            "nb_pax":      nb_pax,
            "marge":       marge,
            "prix_unitaire": prix,
            "depense":     dep,
            "total":       total,
        }

    def _hotels_for_city(self, ville: str) -> list:
        """Retourne les labels d'hôtels correspondant à la ville (ou tous si pas de match)."""
        norm = _normalize(ville)
        if not norm:
            return self._hotel_labels
        filtered = self._hotels_by_city.get(norm, [])
        return filtered if filtered else self._hotel_labels

    def _hotel_currency(self, hotel_label: str) -> str:
        """Retourne la devise de l'hôtel ('MGA', 'EUR', 'USD', …)."""
        h = self._hotels_by_label.get(hotel_label)
        if not h:
            return "MGA"
        return (h.get("unite") or "MGA").strip().upper() or "MGA"

    def _get_hotel_prices(self, hotel_label: str, group_key: str) -> dict:
        """Retourne {room_key: price_in_MGA} pour l'hôtel et le groupe sélectionnés.
        Convertit automatiquement EUR/USD → MGA si besoin."""
        h = self._hotels_by_label.get(hotel_label)
        if not h:
            return {}
        room_rates = h.get("room_rates") or {}
        prices = dict(room_rates.get(group_key, {}))
        unite  = self._hotel_currency(hotel_label)
        # Si la devise n'est pas MGA/Ariary, on convertit chaque prix
        if unite not in ("MGA", "ARIARY", "AR", ""):
            converted = {}
            for rk, price in prices.items():
                if price:
                    try:
                        converted[rk] = convert_currency(
                            float(price), unite, "MGA", self._rates
                        )
                    except Exception:
                        converted[rk] = price
                else:
                    converted[rk] = price
            return converted
        return prices

    # ── Rafraîchissement du tableau ────────────────────────────────────────────

    def _refresh_tree(self):
        self._tree.delete(*self._tree.get_children())
        for i, rd in enumerate(self._rows):
            rooming_str = _build_rooming_str(rd["room_prices"])
            prix        = rd["prix_unitaire"]
            dep         = rd["depense"]
            total       = rd["total"]
            tag = "odd" if i % 2 else "even"
            self._tree.insert("", "end", iid=str(i), values=(
                rd["ville"],
                rd["nuits"],
                rd["hotel"],
                rooming_str,
                rd["nb_pax"],
                _fmt(prix)  if prix  else "",
                _fmt(dep)   if dep   else "",
                rd["marge"] + " %" if rd["marge"] else "",
                _fmt(total) if total else "",
            ), tags=(tag,))

    def _refresh_totals(self):
        total_dep  = sum(rd["depense"] for rd in self._rows)
        total_glob = sum(rd["total"]   for rd in self._rows)
        self._lbl_dep.configure(text=_fmt(total_dep))
        self._lbl_glob.configure(text=_fmt(total_glob))

    # ── Sauvegarde Excel ───────────────────────────────────────────────────────

    def _save_to_excel(self):
        if not self._rows:
            messagebox.showwarning("Aucune donnée", "Le tableau est vide. Rien à sauvegarder.")
            return
        result = save_client_hotel_cotation_to_excel(self.client, self._rows)
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

    # ── Actions sur les lignes ─────────────────────────────────────────────────

    def _add_row_dialog(self):
        initial = self._make_row(
            nb_pax      = str(self.client.get("nombre_participants") or
                              self.client.get("nombre_adultes") or ""),
            hotel_group = _client_group_key(self.client),
            room_prices = self._default_room_prices(self.client),
        )
        self._open_row_dialog(initial, row_index=None)

    def _edit_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Aucune sélection", "Sélectionnez une ligne à modifier.")
            return
        idx = int(sel[0])
        self._open_row_dialog(self._rows[idx], row_index=idx)

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

        # ── Variables ──────────────────────────────────────────────────────
        v_ville  = tk.StringVar(value=row["ville"])
        v_nuits  = tk.StringVar(value=row["nuits"])
        v_hotel  = tk.StringVar(value=row["hotel"])
        v_group  = tk.StringVar(value=_GROUP_BY_KEY.get(row["hotel_group"], "Standard"))
        v_pax    = tk.StringVar(value=row["nb_pax"])
        v_marge  = tk.StringVar(value=row["marge"])

        # room_prices : copie locale modifiable
        local_rp: dict = {
            rk: {
                "count": _to_int(row["room_prices"].get(rk, {}).get("count", 0)),
                "price": _to_float(row["room_prices"].get(rk, {}).get("price", 0)),
            }
            for _, rk, _ in _ROOM_TYPES
        }

        # StringVars pour chaque type de chambre
        rp_vars: dict = {
            rk: {
                "count": tk.StringVar(value=str(local_rp[rk]["count"])),
                "price": tk.StringVar(value=str(local_rp[rk]["price"])
                         if local_rp[rk]["price"] else ""),
            }
            for _, rk, _ in _ROOM_TYPES
        }

        # ── Layout principal ───────────────────────────────────────────────
        outer = tk.Frame(win, bg=MAIN_BG_COLOR)
        outer.pack(padx=24, pady=(20, 0))

        SEC = PANEL_BG_COLOR  # fond des sections

        def _lbl(parent, text, r, c=0, span=1, bg=None):
            tk.Label(
                parent, text=text,
                font=LABEL_FONT, fg=TEXT_COLOR, bg=bg or SEC, anchor="w",
            ).grid(row=r, column=c, columnspan=span, sticky="w", padx=(0, 10), pady=5)

        def _entry(parent, var, r, c=1, w=30, justify="left"):
            e = tk.Entry(
                parent, textvariable=var,
                font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                width=w, justify=justify,
                insertbackground=TEXT_COLOR, relief="flat",
            )
            e.grid(row=r, column=c, sticky="ew", pady=5)
            return e

        def _make_section(title):
            """Crée une carte de section avec titre et retourne le frame intérieur."""
            card = tk.Frame(outer, bg=SEC)
            card.pack(fill="x", pady=(0, 10))
            tk.Label(card, text=title, font=LABEL_FONT, fg=TEXT_COLOR, bg=SEC,
                     ).pack(anchor="w", padx=10, pady=(8, 2))
            inner = tk.Frame(card, bg=SEC)
            inner.pack(fill="x", padx=10, pady=(0, 8))
            return inner

        # ── Section 1 : Lieu / Hôtel ───────────────────────────────────────
        s1 = _make_section("Lieu & Hôtel")

        _lbl(s1, "Ville :", 0)
        _entry(s1, v_ville, 0, w=30)

        _lbl(s1, "Nuits :", 1)
        _entry(s1, v_nuits, 1, w=6, justify="center")

        _lbl(s1, "Hôtel :", 2)
        combo_hotel = ttk.Combobox(
            s1, textvariable=v_hotel,
            values=self._hotels_for_city(row["ville"]),
            font=ENTRY_FONT, width=35, state="normal",
        )
        combo_hotel.grid(row=2, column=1, sticky="ew", pady=5)

        # Indicateur de devise
        lbl_devise = tk.Label(
            s1, text="",
            font=ENTRY_FONT, fg=ACCENT_TEXT_COLOR, bg=SEC, anchor="w",
        )
        lbl_devise.grid(row=2, column=2, sticky="w", padx=(8, 0), pady=5)

        _lbl(s1, "Catégorie de chambre :", 3)
        combo_group = ttk.Combobox(
            s1, textvariable=v_group,
            values=_GROUP_LABELS, font=ENTRY_FONT, width=18, state="readonly",
        )
        combo_group.grid(row=3, column=1, sticky="w", pady=5)

        _lbl(s1, "Nb pax :", 4)
        _entry(s1, v_pax, 4, w=6, justify="center")

        # ── Section 2 : Rooming — prix par type de chambre ─────────────────
        s2 = _make_section("Rooming list — Prix par type de chambre")

        # En-têtes du mini-tableau
        for c_i, (hdr, w) in enumerate([
            ("Type", 8), ("Nb chambres", 11), ("Prix / chambre", 16), ("Sous-total", 14),
        ]):
            tk.Label(
                s2, text=hdr,
                font=LABEL_FONT, fg=MUTED_TEXT_COLOR, bg=SEC,
                width=w, anchor="center",
            ).grid(row=0, column=c_i, padx=4, pady=(0, 2))

        # Séparateur
        tk.Frame(s2, bg=MUTED_TEXT_COLOR, height=1).grid(
            row=1, column=0, columnspan=4, sticky="ew", padx=2, pady=(0, 4),
        )

        subtotal_labels: dict = {}  # rk → tk.Label

        def _update_subtotal(rk, *_):
            count = _to_int(rp_vars[rk]["count"].get())
            price = _to_float(rp_vars[rk]["price"].get())
            sub   = count * price
            if rk in subtotal_labels:
                subtotal_labels[rk].configure(
                    text=_fmt(sub) if sub else "0.00",
                    fg=ACCENT_TEXT_COLOR if sub > 0 else MUTED_TEXT_COLOR,
                )
            _update_preview()

        for row_i, (_, rk, lbl) in enumerate(_ROOM_TYPES):
            grid_r = row_i + 2

            tk.Label(
                s2, text=lbl,
                font=LABEL_FONT, fg=TEXT_COLOR, bg=SEC,
                width=8, anchor="center",
            ).grid(row=grid_r, column=0, padx=4, pady=3)

            e_count = tk.Entry(
                s2, textvariable=rp_vars[rk]["count"],
                font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                width=11, justify="center", relief="flat",
                insertbackground=TEXT_COLOR,
            )
            e_count.grid(row=grid_r, column=1, padx=4, pady=3)

            e_price = tk.Entry(
                s2, textvariable=rp_vars[rk]["price"],
                font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                width=16, justify="right", relief="flat",
                insertbackground=TEXT_COLOR,
            )
            e_price.grid(row=grid_r, column=2, padx=4, pady=3)

            lbl_sub = tk.Label(
                s2, text="0.00",
                font=ENTRY_FONT, fg=MUTED_TEXT_COLOR, bg=SEC,
                width=14, anchor="e",
            )
            lbl_sub.grid(row=grid_r, column=3, padx=4, pady=3)
            subtotal_labels[rk] = lbl_sub

            rp_vars[rk]["count"].trace_add("write", lambda *_, r=rk: _update_subtotal(r))
            rp_vars[rk]["price"].trace_add("write", lambda *_, r=rk: _update_subtotal(r))

        # Séparateur + ligne Prix unitaire total
        sep_r = len(_ROOM_TYPES) + 2
        tk.Frame(s2, bg=MUTED_TEXT_COLOR, height=1).grid(
            row=sep_r, column=0, columnspan=4, sticky="ew", padx=2, pady=(4, 2),
        )
        tk.Label(
            s2, text="Prix unitaire total :",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=SEC,
        ).grid(row=sep_r + 1, column=2, sticky="e", padx=(0, 4), pady=4)

        lbl_prix_total = tk.Label(
            s2, text="0.00",
            font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=SEC,
            width=14, anchor="e",
        )
        lbl_prix_total.grid(row=sep_r + 1, column=3, padx=4, pady=4)

        # ── Section 3 : Marge + Preview ───────────────────────────────────
        s3 = tk.Frame(outer, bg=MAIN_BG_COLOR)
        s3.pack(fill="x", pady=(0, 8))

        _lbl(s3, "Marge (%) :", 0, bg=MAIN_BG_COLOR)
        _entry(s3, v_marge, 0, w=8, justify="center")

        lbl_prev = tk.Label(
            s3, text="Dépense : —   |   Total : —",
            font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=MAIN_BG_COLOR,
        )
        lbl_prev.grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

        # ── Preview / calcul ───────────────────────────────────────────────
        def _update_preview(*_):
            prix  = sum(
                _to_int(rp_vars[rk]["count"].get()) * _to_float(rp_vars[rk]["price"].get())
                for _, rk, _ in _ROOM_TYPES
            )
            nuits = _to_float(v_nuits.get(), 1.0)
            marge = _to_float(v_marge.get())
            dep   = prix * nuits
            total = dep * (1 + marge / 100)
            lbl_prix_total.configure(text=_fmt(prix))
            lbl_prev.configure(
                text=f"Prix unitaire : {_fmt(prix)}   |   "
                     f"Dépense : {_fmt(dep)}   |   Total : {_fmt(total)}"
            )

        v_nuits.trace_add("write", _update_preview)
        v_marge.trace_add("write", _update_preview)

        # ── Filtre hôtel par ville ─────────────────────────────────────────
        def _update_hotel_filter(*_):
            """Met à jour la liste déroulante des hôtels selon la ville saisie."""
            filtered = self._hotels_for_city(v_ville.get())
            combo_hotel["values"] = filtered
            if v_hotel.get() and v_hotel.get() not in filtered:
                v_hotel.set("")
                lbl_devise.configure(text="")

        v_ville.trace_add("write", _update_hotel_filter)

        # ── Auto-remplissage depuis la BD ──────────────────────────────────
        def _autofill_prices(*_):
            """Remplit les prix depuis la BD quand hôtel ou catégorie change.
            Convertit EUR/USD → MGA automatiquement."""
            hotel_lbl = v_hotel.get().strip()
            group_lbl = v_group.get().strip()
            group_key = _GROUP_BY_LBL.get(group_lbl, "standard")

            unite = self._hotel_currency(hotel_lbl) if hotel_lbl else "MGA"
            if unite not in ("MGA", "ARIARY", "AR", ""):
                rate = self._rates.get(unite, None)
                if rate:
                    lbl_devise.configure(
                        text=f"⚠ Prix en {unite} → converti en MGA  (1 {unite} ≈ {rate:,.0f} Ar)"
                    )
                else:
                    lbl_devise.configure(text=f"⚠ Prix en {unite} → converti en MGA")
            else:
                lbl_devise.configure(text="")

            prices = self._get_hotel_prices(hotel_lbl, group_key)
            if not prices:
                return

            for _, rk, _ in _ROOM_TYPES:
                db_price = prices.get(rk)
                if db_price:
                    rp_vars[rk]["price"].set(f"{float(db_price):.2f}")

            _update_preview()

        combo_hotel.bind("<<ComboboxSelected>>", _autofill_prices)
        combo_group.bind("<<ComboboxSelected>>", _autofill_prices)

        # Initialisation des sous-totaux
        for _, rk, _ in _ROOM_TYPES:
            _update_subtotal(rk)
        _update_preview()

        # ── Boutons Valider / Annuler ──────────────────────────────────────
        btn_bar = tk.Frame(win, bg=MAIN_BG_COLOR)
        btn_bar.pack(pady=(8, 20), padx=24, fill="x")

        def _save():
            # Reconstituer room_prices depuis les vars
            new_rp = {
                rk: {
                    "count": _to_int(rp_vars[rk]["count"].get()),
                    "price": _to_float(rp_vars[rk]["price"].get()),
                }
                for _, rk, _ in _ROOM_TYPES
            }
            group_lbl = v_group.get().strip()
            new_row = self._make_row(
                ville       = v_ville.get().strip(),
                nuits       = v_nuits.get().strip(),
                hotel       = v_hotel.get().strip(),
                hotel_group = _GROUP_BY_LBL.get(group_lbl, "standard"),
                room_prices = new_rp,
                nb_pax      = v_pax.get().strip(),
                marge       = v_marge.get().strip(),
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

        # Centrer la fenêtre
        win.update_idletasks()
        pw = self.parent.winfo_rootx() + self.parent.winfo_width() // 2
        ph = self.parent.winfo_rooty() + self.parent.winfo_height() // 2
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        win.geometry(f"{w}x{h}+{pw - w // 2}+{ph - h // 2}")

