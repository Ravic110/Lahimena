"""
Cotation transport par client — une ligne par segment de trajet client.

Les segments sont générés automatiquement depuis l'itinéraire du client :
  ex. « Antananarivo → Antsirabe → Fianarantsoa »
  donne deux lignes : Antananarivo → Antsirabe, Antsirabe → Fianarantsoa.

Pour chaque ligne, l'utilisateur choisit le prestataire et le type de véhicule.
Le KM est pré-rempli depuis KM_MADA (ville d'arrivée) et reste éditable.

Total = (prix_jour × nb_jours) + carburant
Carburant = consommation (L/100 km) × km × prix_carburant / 100
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
    get_km_mada_km_for_repere,
    get_km_mada_reperes,
    get_segment_distance,
    get_transport_fuel_price,
    get_transport_prestataires,
    get_transport_vehicle_data,
    get_transport_vehicle_types,
    load_client_transport_cotation,
    normalize_city_name,
    save_client_transport_cotation_to_excel,
)

# ── Hover colors ───────────────────────────────────────────────────────────────
_HOVER_GREEN = "#0A6870"
_HOVER_BLUE  = "#0B6080"
_HOVER_RED   = "#A82020"


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


def _parse_cities(raw: str) -> list:
    """Extrait une liste ordonnée de villes depuis une chaîne d'itinéraire."""
    if not raw:
        return []
    seen, result = set(), []
    # Sépare sur virgule, point-virgule, tiret double, slash, pipe, retour ligne
    for city in re.split(r"[,;/|\n]|(?<!\w)-(?!\w)|--", str(raw)):
        city = city.strip(" -\t")
        # Retire un éventuel préfixe « J1 - », « Jour 1 : », « 1. » etc.
        city = re.sub(r"^(?:j(?:our)?\s*\d+\s*[-:.]?\s*|\d+\s*[-:.]\s*)", "",
                      city, flags=re.IGNORECASE).strip()
        if not city:
            continue
        norm = _normalize(city)
        if norm and norm not in seen:
            seen.add(norm)
            result.append(city)
    return result


def _make_segments(client: dict) -> list:
    """
    Construit la liste de segments (depart, arrivee) depuis l'itinéraire client.
    Retourne [(depart, arrivee), ...].
    """
    raw_itin = client.get("itineraire_circuit") or ""
    raw_va   = client.get("ville_arrivee") or ""
    raw_dep  = (client.get("ville_depart") or "").strip()

    cities = _parse_cities(raw_itin) or _parse_cities(raw_va)

    # Préfixe avec la ville de départ si absente
    if raw_dep:
        norm_dep = _normalize(raw_dep)
        if not cities or _normalize(cities[0]) != norm_dep:
            cities = [raw_dep] + cities

    if len(cities) < 2:
        return []

    return [(cities[i], cities[i + 1]) for i in range(len(cities) - 1)]


def _compute_carburant(consommation: float, km: float,
                       prix_carburant: float) -> float:
    """Budget carburant = conso (L/100 km) × km × prix_carburant / 100."""
    return consommation * km * prix_carburant / 100.0


def _make_row(depart="", arrivee="", km_distance="", prestataire="",
              type_voiture="", nb_places="", nb_vehicules="1",
              nb_jours="1", prix_jour=0.0, km="",
              consommation=0.0, energie="") -> dict:
    prix_carburant = get_transport_fuel_price(energie) if energie else 0.0
    km_f      = _to_float(km)
    carburant = _compute_carburant(consommation, km_f, prix_carburant)
    nv        = max(1, _to_int(nb_vehicules, 1))
    total     = nv * (_to_float(nb_jours, 1.0) * prix_jour + carburant)
    return {
        "depart":       depart,
        "arrivee":      arrivee,
        "km_distance":  km_distance,   # km affiché dans la colonne trajet
        "prestataire":  prestataire,
        "type_voiture": type_voiture,
        "nb_places":    nb_places,
        "nb_vehicules": str(nv),
        "nb_jours":     nb_jours,
        "prix_jour":    prix_jour,
        "km":           km,
        "consommation": consommation,
        "energie":      energie,
        "carburant":    carburant,
        "total":        total,
    }


# ── Classe principale ─────────────────────────────────────────────────────────

class ClientTransportCotation:
    """Cotation transport par trajet pour un client — une ligne par segment."""

    _COLS = [
        ("trajet",        "Trajet",         260),
        ("prestataire",   "Prestataire",    150),
        ("type_voiture",  "Type véhicule",  120),
        ("nb_vehicules",  "Nb véh.",         60),
        ("nb_jours",      "Nb jours",        65),
        ("prix_jour",     "Prix/jour",      110),
        ("km",            "KM route",        75),
        ("carburant",     "Carburant",      110),
        ("total",         "Total",          120),
    ]

    def __init__(self, parent: tk.Widget, client: dict, on_back=None):
        self.parent  = parent
        self.client  = client
        self.on_back = on_back
        self._rows: list = []

        self._build_ui()
        self._populate_initial_rows()
        self._refresh_tree()
        self._refresh_totals()

    # ── Interface ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        client  = self.client
        nom     = client.get("nom", "")
        prenom  = client.get("prenom", "")
        dossier = client.get("numero_dossier", "")
        pax     = str(client.get("nombre_participants") or
                      client.get("nombre_adultes") or "")
        sejour  = str(client.get("duree_sejour") or "")

        # Titre
        tk.Label(
            self.parent, text="COTATION TRANSPORT",
            font=TITLE_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR,
        ).pack(pady=(20, 4))

        client_name = f"{prenom} {nom}".strip() or "—"
        tk.Label(
            self.parent,
            text="Client : " + client_name +
                 (f"   |   Dossier : {dossier}" if dossier else ""),
            font=ENTRY_FONT, fg=MUTED_TEXT_COLOR, bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 2))

        if self.on_back:
            back_bar = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
            back_bar.pack(fill="x", padx=20, pady=(4, 0))
            ctk.CTkButton(
                back_bar, text="⬅  Retour à l'accueil",
                command=self.on_back,
                fg_color=BUTTON_BLUE, hover_color=_HOVER_BLUE,
                text_color="white", font=BUTTON_FONT,
                corner_radius=8, cursor="hand2",
            ).pack(side="left")

        root = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        root.pack(fill="both", expand=True, padx=20, pady=(8, 20))

        # ── Carte infos client ─────────────────────────────────────────────
        info_card = tk.Frame(root, bg=PANEL_BG_COLOR)
        info_card.pack(fill="x", pady=(0, 10))
        tk.Label(
            info_card, text="Informations client",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(anchor="w", padx=12, pady=(8, 2))
        info_inner = tk.Frame(info_card, bg=PANEL_BG_COLOR)
        info_inner.pack(fill="x", padx=12, pady=(0, 8))

        # Itinéraire résumé
        segments = _make_segments(client)
        itin_str = "  →  ".join(
            [seg[0] for seg in segments] + ([segments[-1][1]] if segments else [])
        ) if segments else (client.get("itineraire_circuit") or "—")

        for col_i, (lbl_text, val_text) in enumerate([
            ("Participants :", pax or "—"),
            ("Durée séjour :", f"{sejour} j" if sejour else "—"),
        ]):
            tk.Label(
                info_inner, text=lbl_text,
                font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
            ).grid(row=0, column=col_i * 2, sticky="w",
                   padx=(0 if col_i == 0 else 20, 4), pady=4)
            tk.Label(
                info_inner, text=val_text,
                font=ENTRY_FONT, fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR,
            ).grid(row=0, column=col_i * 2 + 1, sticky="w", pady=4)

        tk.Label(
            info_inner, text="Itinéraire :",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).grid(row=1, column=0, sticky="w", padx=(0, 4), pady=(0, 4))
        tk.Label(
            info_inner, text=itin_str,
            font=ENTRY_FONT, fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR,
            wraplength=700, justify="left",
        ).grid(row=1, column=1, columnspan=3, sticky="w", pady=(0, 4))

        # ── Barre d'actions ────────────────────────────────────────────────
        action_bar = tk.Frame(root, bg=PANEL_BG_COLOR)
        action_bar.pack(fill="x", pady=(0, 6))

        ctk.CTkButton(
            action_bar, text="＋  Ajouter un trajet",
            command=self._add_row_dialog,
            fg_color=BUTTON_GREEN, hover_color=_HOVER_GREEN,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2",
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="✏️  Modifier",
            command=self._edit_selected,
            fg_color=BUTTON_BLUE, hover_color=_HOVER_BLUE,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2",
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="🗑️  Supprimer",
            command=self._delete_selected,
            fg_color=BUTTON_RED, hover_color=_HOVER_RED,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2",
        ).pack(side="left")

        ctk.CTkButton(
            action_bar, text="💾  Sauvegarder",
            command=self._save_to_excel,
            fg_color=BUTTON_GREEN, hover_color=_HOVER_GREEN,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2",
        ).pack(side="right")

        # ── Treeview ───────────────────────────────────────────────────────
        style = ttk.Style()
        style.configure(
            "Transport.Treeview",
            background=INPUT_BG_COLOR, foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR, rowheight=30,
        )
        style.configure("Transport.Treeview.Heading", font=LABEL_FONT)
        style.map("Transport.Treeview", background=[("selected", BUTTON_BLUE)])

        col_ids   = [c[0] for c in self._COLS]
        tree_outer = tk.Frame(root, bg=PANEL_BG_COLOR)
        tree_outer.pack(fill="both", expand=True, pady=(0, 6))

        self._tree = ttk.Treeview(
            tree_outer, columns=col_ids, show="headings",
            height=10, style="Transport.Treeview",
        )
        for key, heading, width in self._COLS:
            self._tree.heading(key, text=heading)
            anchor = "e" if key in (
                "nb_vehicules", "nb_jours", "prix_jour", "km", "carburant", "total"
            ) else "w"
            self._tree.column(key, width=width, anchor=anchor, stretch=False)

        vsb = ttk.Scrollbar(tree_outer, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_outer, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_outer.grid_rowconfigure(0, weight=1)
        tree_outer.grid_columnconfigure(0, weight=1)
        self._tree.bind("<Double-1>", lambda e: self._edit_selected())
        self._tree.tag_configure("odd",  background="#E4F2F6")
        self._tree.tag_configure("even", background=INPUT_BG_COLOR)

        # ── Totaux ─────────────────────────────────────────────────────────
        totals_inner = tk.Frame(root, bg=PANEL_BG_COLOR)
        totals_inner.pack(fill="x", pady=(4, 6))

        tk.Label(
            totals_inner, text="Total carburant :",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        self._lbl_carb = tk.Label(
            totals_inner, text="0.00",
            font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR,
        )
        self._lbl_carb.grid(row=0, column=1, sticky="w", padx=(0, 40), pady=4)

        tk.Label(
            totals_inner, text="Total global :",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=4)
        self._lbl_total = tk.Label(
            totals_inner, text="0.00",
            font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR,
        )
        self._lbl_total.grid(row=0, column=3, sticky="w", pady=4)

    # ── Données initiales ──────────────────────────────────────────────────────

    def _populate_initial_rows(self):
        # Priorité : données sauvegardées
        saved = load_client_transport_cotation(self.client)
        if saved:
            self._rows = saved
            return

        # Sinon : générer depuis l'itinéraire client
        segments = _make_segments(self.client)
        for depart_raw, arrivee_raw in segments:
            depart  = normalize_city_name(depart_raw)
            arrivee = normalize_city_name(arrivee_raw)
            km_val  = get_segment_distance(depart, arrivee)
            km_str  = str(int(km_val)) if km_val else ""
            self._rows.append(_make_row(
                depart=depart, arrivee=arrivee,
                km_distance=km_str,
                km=km_str,
            ))

        if not self._rows:
            # Aucun itinéraire parsable : ligne vide
            self._rows.append(_make_row())

    # ── Rafraîchissement ───────────────────────────────────────────────────────

    def _refresh_tree(self):
        self._tree.delete(*self._tree.get_children())
        for i, rd in enumerate(self._rows):
            tag  = "odd" if i % 2 else "even"
            # Colonne Trajet : "Ville A → Ville B (169 km)"
            km_d = rd.get("km_distance") or rd.get("km") or ""
            km_suffix = f" ({km_d} km)" if km_d else ""
            if rd.get("depart") or rd.get("arrivee"):
                trajet = f"{rd['depart']}  →  {rd['arrivee']}{km_suffix}"
            else:
                trajet = "—"
            prix_j = rd["prix_jour"]
            carb   = rd["carburant"]
            total  = rd["total"]
            self._tree.insert("", "end", iid=str(i), values=(
                trajet,
                rd["prestataire"],
                rd["type_voiture"],
                rd.get("nb_vehicules", "1"),
                rd["nb_jours"],
                _fmt(prix_j) if prix_j else "",
                rd["km"],
                _fmt(carb)   if carb   else "",
                _fmt(total)  if total  else "",
            ), tags=(tag,))

    def _refresh_totals(self):
        self._lbl_carb.configure(
            text=_fmt(sum(rd["carburant"] for rd in self._rows)))
        self._lbl_total.configure(
            text=_fmt(sum(rd["total"] for rd in self._rows)))

    # ── Sauvegarde ─────────────────────────────────────────────────────────────

    def _save_to_excel(self):
        if not self._rows:
            messagebox.showwarning("Aucune donnée",
                                   "Le tableau est vide. Rien à sauvegarder.")
            return
        result = save_client_transport_cotation_to_excel(self.client, self._rows)
        if result > 0:
            messagebox.showinfo("Sauvegarde réussie",
                                f"{result} ligne(s) enregistrée(s).")
        elif result == -2:
            messagebox.showerror(
                "Fichier verrouillé",
                "Le fichier Excel est ouvert ailleurs.\n"
                "Fermez data.xlsx puis réessayez.",
            )
        else:
            messagebox.showerror(
                "Erreur",
                "La sauvegarde a échoué. Consultez les logs.",
            )

    # ── Actions ────────────────────────────────────────────────────────────────

    def _add_row_dialog(self):
        self._open_row_dialog(_make_row(), row_index=None)

    def _edit_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Aucune sélection",
                                   "Sélectionnez une ligne à modifier.")
            return
        self._open_row_dialog(self._rows[int(sel[0])], row_index=int(sel[0]))

    def _delete_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Aucune sélection",
                                   "Sélectionnez une ligne à supprimer.")
            return
        if not messagebox.askyesno("Supprimer",
                                   "Supprimer la ligne sélectionnée ?"):
            return
        del self._rows[int(sel[0])]
        self._refresh_tree()
        self._refresh_totals()

    # ── Dialog d'édition ──────────────────────────────────────────────────────

    def _open_row_dialog(self, row: dict, row_index):
        win = tk.Toplevel(self.parent)
        win.title("Modifier le trajet" if row_index is not None
                  else "Ajouter un trajet")
        win.configure(bg=MAIN_BG_COLOR)
        win.resizable(False, False)
        win.transient(self.parent)
        win.after(0, lambda: [win.lift(), win.focus_set()])

        # ── Variables ──────────────────────────────────────────────────────
        v_depart       = tk.StringVar(value=row.get("depart", ""))
        v_arrivee      = tk.StringVar(value=row.get("arrivee", ""))
        v_prestataire  = tk.StringVar(value=row.get("prestataire", ""))
        v_type_voiture = tk.StringVar(value=row.get("type_voiture", ""))
        v_nb_places    = tk.StringVar(value=row.get("nb_places", ""))
        v_nb_vehicules = tk.StringVar(value=row.get("nb_vehicules", "1"))
        v_nb_jours     = tk.StringVar(value=row.get("nb_jours", "1"))
        v_prix_jour    = tk.StringVar(
            value=str(row["prix_jour"]) if row.get("prix_jour") else "")
        v_km           = tk.StringVar(value=row.get("km", ""))
        v_consommation = tk.StringVar(
            value=str(row["consommation"]) if row.get("consommation") else "")
        v_energie      = tk.StringVar(value=row.get("energie", ""))
        v_carburant    = tk.StringVar(value="")
        v_total        = tk.StringVar(value="")

        # ── Layout ────────────────────────────────────────────────────────
        outer = tk.Frame(win, bg=MAIN_BG_COLOR)
        outer.pack(padx=24, pady=(20, 0))

        SEC = PANEL_BG_COLOR

        def _lbl(parent, text, r, c=0, bg=None):
            tk.Label(
                parent, text=text,
                font=LABEL_FONT, fg=TEXT_COLOR, bg=bg or SEC, anchor="w",
            ).grid(row=r, column=c, sticky="w", padx=(0, 10), pady=5)

        def _entry(parent, var, r, c=1, w=28, justify="left",
                   state="normal", readonly_bg=None):
            bg = readonly_bg or (INPUT_BG_COLOR if state == "normal"
                                 else PANEL_BG_COLOR)
            e = tk.Entry(
                parent, textvariable=var,
                font=ENTRY_FONT, bg=bg, fg=TEXT_COLOR,
                width=w, justify=justify,
                insertbackground=TEXT_COLOR, relief="flat",
                state=state,
            )
            e.grid(row=r, column=c, sticky="ew", pady=5)
            return e

        def _make_section(title):
            card = tk.Frame(outer, bg=SEC)
            card.pack(fill="x", pady=(0, 10))
            tk.Label(card, text=title, font=LABEL_FONT, fg=TEXT_COLOR,
                     bg=SEC).pack(anchor="w", padx=10, pady=(8, 2))
            inner = tk.Frame(card, bg=SEC)
            inner.pack(fill="x", padx=10, pady=(0, 8))
            return inner

        # Liste des villes depuis KM_MADA pour les comboboxes
        from utils.excel_handler import get_km_mada_reperes as _get_reperes
        _km_cities = _get_reperes()

        # ── Section 1 : Trajet ─────────────────────────────────────────────
        s1 = _make_section("Trajet")

        _lbl(s1, "Départ :", 0)
        combo_depart = ttk.Combobox(
            s1, textvariable=v_depart,
            values=_km_cities, font=ENTRY_FONT, width=26, state="normal",
        )
        combo_depart.grid(row=0, column=1, sticky="ew", pady=5)

        _lbl(s1, "Arrivée :", 1)
        combo_arrivee = ttk.Combobox(
            s1, textvariable=v_arrivee,
            values=_km_cities, font=ENTRY_FONT, width=26, state="normal",
        )
        combo_arrivee.grid(row=1, column=1, sticky="ew", pady=5)

        # Champ KM — lecture seule, auto depuis KM_MADA
        _lbl(s1, "Distance (KM) :", 2)
        _entry(s1, v_km, 2, w=10, justify="right", state="readonly")
        tk.Label(
            s1, text="automatique depuis la BD",
            font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=SEC,
        ).grid(row=2, column=2, sticky="w", padx=(6, 0), pady=5)

        # ── Section 2 : Véhicule ───────────────────────────────────────────
        s2 = _make_section("Véhicule")

        _lbl(s2, "Prestataire :", 0)
        combo_prest = ttk.Combobox(
            s2, textvariable=v_prestataire,
            values=get_transport_prestataires(),
            font=ENTRY_FONT, width=26, state="normal",
        )
        combo_prest.grid(row=0, column=1, sticky="ew", pady=5)

        _lbl(s2, "Type de véhicule :", 1)
        combo_type = ttk.Combobox(
            s2, textvariable=v_type_voiture,
            values=get_transport_vehicle_types(v_prestataire.get() or None),
            font=ENTRY_FONT, width=26, state="normal",
        )
        combo_type.grid(row=1, column=1, sticky="ew", pady=5)

        _lbl(s2, "Nombre de places :", 2)
        _entry(s2, v_nb_places, 2, w=8, justify="center", state="readonly")

        _lbl(s2, "Nb véhicules :", 3)
        e_nv = _entry(s2, v_nb_vehicules, 3, w=8, justify="center")
        # Petite note explicative
        tk.Label(
            s2, text="(si groupe nombreux, ex. 2 véhicules)",
            font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=SEC,
        ).grid(row=3, column=2, sticky="w", padx=(6, 0), pady=5)

        _lbl(s2, "Prix/jour (MGA) :", 4)
        _entry(s2, v_prix_jour, 4, w=16, justify="right", state="readonly")

        _lbl(s2, "Consommation (L/100) :", 5)
        _entry(s2, v_consommation, 5, w=10, justify="right", state="readonly")

        _lbl(s2, "Énergie :", 6)
        _entry(s2, v_energie, 6, w=16, state="readonly")

        # ── Section 3 : Utilisation ────────────────────────────────────────
        s3 = _make_section("Utilisation")

        _lbl(s3, "Nombre de jours :", 0)
        _entry(s3, v_nb_jours, 0, w=8, justify="center")

        # ── Aperçu ─────────────────────────────────────────────────────────
        prev_card = tk.Frame(outer, bg=SEC)
        prev_card.pack(fill="x", pady=(0, 10))
        tk.Label(prev_card, text="Calcul", font=LABEL_FONT,
                 fg=TEXT_COLOR, bg=SEC).pack(anchor="w", padx=10, pady=(8, 2))
        prev_inner = tk.Frame(prev_card, bg=SEC)
        prev_inner.pack(fill="x", padx=10, pady=(0, 8))

        for r, (lbl_t, var) in enumerate([
            ("Budget carburant (MGA) :", v_carburant),
            ("Total (MGA) :",            v_total),
        ]):
            tk.Label(prev_inner, text=lbl_t, font=LABEL_FONT,
                     fg=TEXT_COLOR, bg=SEC, anchor="w",
                     ).grid(row=r, column=0, sticky="w", padx=(0, 10), pady=3)
            tk.Label(prev_inner, textvariable=var, font=ENTRY_FONT,
                     fg=ACCENT_TEXT_COLOR, bg=SEC,
                     ).grid(row=r, column=1, sticky="w", pady=3)

        # ── Callbacks (définis après tous les widgets) ─────────────────────
        def _update_preview(*_):
            nv        = max(1, _to_int(v_nb_vehicules.get(), 1))
            nb_j      = _to_float(v_nb_jours.get(), 1.0)
            prix_j    = _to_float(v_prix_jour.get())
            km_v      = _to_float(v_km.get())
            conso     = _to_float(v_consommation.get())
            energie   = v_energie.get()
            prix_carb = get_transport_fuel_price(energie) if energie else 0.0
            carb      = _compute_carburant(conso, km_v, prix_carb)
            total     = nv * (prix_j * nb_j + carb)
            v_carburant.set(_fmt(carb))
            v_total.set(_fmt(total))

        def _on_arrivee_change(*_):
            """Mise à jour du KM = abs(km_arr - km_dep) depuis KM_MADA."""
            depart  = normalize_city_name(v_depart.get().strip())
            arrivee = normalize_city_name(v_arrivee.get().strip())
            km_val  = get_segment_distance(depart, arrivee) if arrivee else 0
            v_km.set(str(int(km_val)) if km_val else "")
            _update_preview()

        def _on_type_change(*_):
            prest = v_prestataire.get()
            tv    = v_type_voiture.get()
            if prest and tv:
                data = get_transport_vehicle_data(prest, tv)
                v_nb_places.set(
                    str(int(data["nombre_place"])) if data.get("nombre_place") else "")
                v_prix_jour.set(
                    str(data["location_par_jour"]) if data.get("location_par_jour") else "")
                v_consommation.set(
                    str(data["consommation"]) if data.get("consommation") else "")
                v_energie.set(data.get("energie", ""))
            _update_preview()

        def _on_prestataire_change(*_):
            prest = v_prestataire.get()
            types = get_transport_vehicle_types(prest or None)
            combo_type["values"] = types
            if v_type_voiture.get() not in types:
                v_type_voiture.set(types[0] if types else "")
            _on_type_change()

        # Traces (après définition des callbacks)
        v_depart.trace_add("write",  lambda *a: _on_arrivee_change())
        v_arrivee.trace_add("write", lambda *a: _on_arrivee_change())
        combo_arrivee.bind("<<ComboboxSelected>>", _on_arrivee_change)
        combo_depart.bind("<<ComboboxSelected>>",  _on_arrivee_change)
        v_prestataire.trace_add("write", _on_prestataire_change)
        v_type_voiture.trace_add("write", lambda *a: (_on_type_change(),))
        v_nb_vehicules.trace_add("write", _update_preview)
        v_nb_jours.trace_add("write", _update_preview)

        # Initialisation au chargement du dialog
        _on_arrivee_change()   # force le km depuis la BD selon l'arrivée actuelle
        _on_type_change()      # force le prix/jour, conso, énergie selon le véhicule
        _update_preview()      # calcule carburant + total

        # ── Boutons ────────────────────────────────────────────────────────
        btn_bar = tk.Frame(win, bg=MAIN_BG_COLOR)
        btn_bar.pack(fill="x", padx=24, pady=(8, 20))

        def _on_ok():
            depart  = v_depart.get().strip()
            arrivee = v_arrivee.get().strip()
            prest   = v_prestataire.get().strip()
            tv      = v_type_voiture.get().strip()
            if not prest or not tv:
                messagebox.showwarning(
                    "Champs manquants",
                    "Veuillez sélectionner un prestataire et un type de véhicule.",
                    parent=win,
                )
                return
            # km_distance : garde la valeur de référence pour l'affichage
            km_dist = row.get("km_distance") or v_km.get()
            new_row = _make_row(
                depart=depart, arrivee=arrivee,
                km_distance=km_dist,
                prestataire=prest, type_voiture=tv,
                nb_places=v_nb_places.get(),
                nb_vehicules=v_nb_vehicules.get(),
                nb_jours=v_nb_jours.get(),
                prix_jour=_to_float(v_prix_jour.get()),
                km=v_km.get(),
                consommation=_to_float(v_consommation.get()),
                energie=v_energie.get(),
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
            command=_on_ok,
            fg_color=BUTTON_GREEN, hover_color=_HOVER_GREEN,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2",
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_bar, text="✖  Annuler",
            command=win.destroy,
            fg_color=BUTTON_RED, hover_color=_HOVER_RED,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2",
        ).pack(side="left")

        # Centrage
        win.update_idletasks()
        pw = self.parent.winfo_toplevel()
        ww = win.winfo_reqwidth() + 48
        wh = win.winfo_reqheight() + 20
        px = pw.winfo_rootx() + (pw.winfo_width()  - ww) // 2
        py = pw.winfo_rooty() + (pw.winfo_height() - wh) // 2
        win.geometry(f"{ww}x{wh}+{px}+{py}")
