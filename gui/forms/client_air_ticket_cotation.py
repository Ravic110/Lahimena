"""
Cotation avion par client.

Pré-remplit depuis la fiche client (compagnie, villes, pax),
génère aller simple ou aller-retour, marge (%) par ligne,
total manuel optionnel, sauvegarde dans la feuille COTATION_AVION.

Améliorations :
- Compagnie depuis la BD avion (combobox)
- Auto-remplissage des tarifs adulte/enfant depuis la BD
- Date de vol par ligne
- Numéro de vol
- Classe de voyage (Économique / Affaires / Première)
- Sous-totaux adultes/enfants dans le pied de page
- Export PDF
"""

import os
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    ACCENT_TEXT_COLOR,
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_ORANGE,
    BUTTON_RED,
    DEVIS_FOLDER,
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
    get_avion_compagnies,
    get_avion_tarifs,
    load_client_air_ticket_cotation,
    save_client_air_ticket_cotation_to_excel,
)

_HOVER_GREEN  = "#0A6870"
_HOVER_BLUE   = "#0B6080"
_HOVER_RED    = "#A82020"
_HOVER_ORANGE = "#C8860A"

_CLASSES = ["Économique", "Affaires", "Première"]


# ── Helpers purs ──────────────────────────────────────────────────────────────

def _to_float(value, default=0.0):
    try:
        return float(str(value).replace(",", ".").strip() or default)
    except (TypeError, ValueError):
        return default


def _fmt(value):
    return f"{value:,.2f}"


def _make_row(
    date_vol="",
    numero_vol="",
    type_trajet="aller",
    compagnie="",
    ville_depart="",
    ville_arrivee="",
    classe="Économique",
    nb_adultes="",
    nb_enfants="",
    tarif_adulte="",
    tarif_enfant="",
    montant_adultes=0.0,
    montant_enfants=0.0,
    sous_total=0.0,
    marge_pct="",
    total=0.0,
    total_manuel=False,
):
    if total in ("", None):
        total_value = ""
    else:
        total_value = float(total or 0.0)
    return {
        "date_vol":        str(date_vol),
        "numero_vol":      str(numero_vol),
        "type_trajet":     type_trajet,
        "compagnie":       compagnie,
        "ville_depart":    ville_depart,
        "ville_arrivee":   ville_arrivee,
        "classe":          classe or "Économique",
        "nb_adultes":      str(nb_adultes),
        "nb_enfants":      str(nb_enfants),
        "tarif_adulte":    str(tarif_adulte),
        "tarif_enfant":    str(tarif_enfant),
        "montant_adultes": float(montant_adultes),
        "montant_enfants": float(montant_enfants),
        "sous_total":      float(sous_total),
        "marge_pct":       str(marge_pct),
        "total":           total_value,
        "total_manuel":    bool(total_manuel),
    }


def _recompute_row(row):
    updated = dict(row)
    montant_adultes = _to_float(updated.get("nb_adultes")) * _to_float(updated.get("tarif_adulte"))
    montant_enfants = _to_float(updated.get("nb_enfants")) * _to_float(updated.get("tarif_enfant"))
    sous_total = montant_adultes + montant_enfants
    marge_pct = _to_float(updated.get("marge_pct"))
    total_calcule = sous_total * (1 + marge_pct / 100.0)

    updated["montant_adultes"] = montant_adultes
    updated["montant_enfants"] = montant_enfants
    updated["sous_total"] = sous_total

    if updated.get("total_manuel"):
        manual_val = _to_float(updated.get("total", 0))
        if manual_val == 0.0:
            updated["total"] = total_calcule
            updated["total_manuel"] = False
        else:
            updated["total"] = manual_val
    else:
        updated["total"] = total_calcule

    return updated


def _build_initial_rows(client, trip_mode):
    compagnie  = str(client.get("compagnie") or "").strip()
    depart     = str(client.get("ville_depart") or "").strip()
    arrivee    = str(client.get("ville_arrivee") or "").strip()
    nb_adultes = str(client.get("nombre_adultes") or client.get("nombre_participants") or "")
    nb_enfants = str(client.get("nombre_enfants") or "0")

    # Tarifs depuis BD si compagnie connue
    tarif_a, tarif_e = 0.0, 0.0
    if compagnie:
        try:
            tarif_a, tarif_e = get_avion_tarifs({"compagnie": compagnie})
        except Exception:
            pass

    aller = _make_row(
        type_trajet="aller",
        compagnie=compagnie,
        ville_depart=depart,
        ville_arrivee=arrivee,
        nb_adultes=nb_adultes,
        nb_enfants=nb_enfants,
        tarif_adulte=str(tarif_a) if tarif_a else "",
        tarif_enfant=str(tarif_e) if tarif_e else "",
        marge_pct="0",
    )
    if trip_mode != "aller-retour":
        return [_recompute_row(aller)]

    retour = _make_row(
        type_trajet="retour",
        compagnie=compagnie,
        ville_depart=arrivee,
        ville_arrivee=depart,
        nb_adultes=nb_adultes,
        nb_enfants=nb_enfants,
        tarif_adulte=str(tarif_a) if tarif_a else "",
        tarif_enfant=str(tarif_e) if tarif_e else "",
        marge_pct="0",
    )
    return [_recompute_row(aller), _recompute_row(retour)]


def _validate_rows(rows):
    errors = []
    numeric_fields = ("nb_adultes", "nb_enfants", "tarif_adulte", "tarif_enfant", "marge_pct")

    for index, row in enumerate(rows, start=1):
        if not str(row.get("ville_depart") or "").strip() or not str(row.get("ville_arrivee") or "").strip():
            errors.append(f"Ligne {index} : ville_depart et ville_arrivee sont obligatoires.")

        for field in numeric_fields:
            value = str(row.get(field, "")).strip()
            if value and _to_float(value, -1) < 0:
                errors.append(f"Ligne {index} : {field} doit etre numerique et >= 0.")

        if row.get("total_manuel") and _to_float(row.get("total"), -1) < 0:
            errors.append(f"Ligne {index} : total doit etre numerique et >= 0.")

    return errors


# ── Écran principal ────────────────────────────────────────────────────────────

class ClientAirTicketCotation:

    _COLS = [
        ("date_vol",      "Date vol",     85),
        ("numero_vol",    "N° vol",       80),
        ("type_trajet",   "Trajet",       70),
        ("compagnie",     "Compagnie",   110),
        ("ville_depart",  "Départ",      115),
        ("ville_arrivee", "Arrivée",     115),
        ("classe",        "Classe",       80),
        ("nb_adultes",    "Adt",          45),
        ("nb_enfants",    "Enf",          45),
        ("sous_total",    "Sous-total",  100),
        ("marge_pct",     "Marge %",      65),
        ("total",         "Total",       105),
    ]

    def __init__(self, parent: tk.Widget, client: dict, on_back=None):
        self.parent   = parent
        self.client   = client
        self.on_back  = on_back
        self._rows    = []
        self.trip_mode_var = tk.StringVar(value="aller-retour")

        self._build_ui()
        self._populate_initial_rows()
        self._refresh_tree()
        self._refresh_totals()

    # ── Construction UI ────────────────────────────────────────────────────────

    def _build_ui(self):
        for w in self.parent.winfo_children():
            w.destroy()

        shell = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        self._shell = shell

        # En-tête
        hdr = tk.Frame(shell, bg=MAIN_BG_COLOR)
        hdr.pack(fill="x", pady=(0, 12))

        ctk.CTkButton(
            hdr, text="← Retour", width=90, height=30,
            fg_color=BUTTON_BLUE, hover_color=_HOVER_BLUE,
            text_color="white", font=("Poppins", 10, "bold"),
            corner_radius=6, cursor="hand2",
            command=self._go_back,
        ).pack(side="left")

        nom = (self.client.get("nom", "") + " " + self.client.get("prenom", "")).strip()
        tk.Label(
            hdr, text=f"Cotation Avion — {nom}",
            font=TITLE_FONT, bg=MAIN_BG_COLOR, fg=TEXT_COLOR,
        ).pack(side="left", padx=16)

        # Barre d'actions
        action_bar = tk.Frame(shell, bg=MAIN_BG_COLOR)
        action_bar.pack(fill="x", pady=(0, 8))

        tk.Label(action_bar, text="Mode :", font=LABEL_FONT,
                 bg=MAIN_BG_COLOR, fg=MUTED_TEXT_COLOR).pack(side="left", padx=(0, 4))

        ttk.Combobox(
            action_bar, textvariable=self.trip_mode_var,
            values=["aller simple", "aller-retour"],
            state="readonly", width=16,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="Régénérer", command=self._regenerate_rows,
            fg_color=BUTTON_BLUE, hover_color=_HOVER_BLUE,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2", width=110,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="+ Ajouter", command=self._open_add_dialog,
            fg_color=BUTTON_GREEN, hover_color=_HOVER_GREEN,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2", width=100,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="✎ Modifier", command=self._open_edit_dialog,
            fg_color=BUTTON_BLUE, hover_color=_HOVER_BLUE,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2", width=100,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="Supprimer", command=self._delete_selected,
            fg_color=BUTTON_RED, hover_color=_HOVER_RED,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2", width=100,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="📄 Devis PDF", command=self._export_pdf,
            fg_color=BUTTON_ORANGE, hover_color=_HOVER_ORANGE,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2", width=120,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            action_bar, text="💾 Sauvegarder", command=self._save_to_excel,
            fg_color=BUTTON_GREEN, hover_color=_HOVER_GREEN,
            text_color="white", font=BUTTON_FONT,
            corner_radius=8, cursor="hand2", width=140,
        ).pack(side="right")

        # Tableau
        tree_frame = tk.Frame(shell, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True)

        cols = [c[0] for c in self._COLS]
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                  selectmode="browse", height=14)

        for col_id, col_label, col_width in self._COLS:
            self._tree.heading(col_id, text=col_label)
            self._tree.column(col_id, width=col_width, minwidth=40, anchor="center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self._tree.tag_configure("even", background="#F5FBFD")
        self._tree.tag_configure("odd",  background="#DCEFF3")
        self._tree.bind("<Double-1>", lambda _e: self._open_edit_dialog())

        # Pied de page
        footer = tk.Frame(shell, bg=PANEL_BG_COLOR, pady=6)
        footer.pack(fill="x", pady=(8, 0))

        def _total_label(parent, title, var):
            tk.Label(parent, text=title, font=LABEL_FONT,
                     bg=PANEL_BG_COLOR, fg=MUTED_TEXT_COLOR).pack(side="left", padx=(16, 2))
            lbl = ctk.CTkLabel(parent, textvariable=var, font=("Poppins", 11, "bold"),
                               text_color=TEXT_COLOR)
            lbl.pack(side="left", padx=(0, 4))
            tk.Label(parent, text="Ar", font=("Poppins", 10),
                     bg=PANEL_BG_COLOR, fg=MUTED_TEXT_COLOR).pack(side="left", padx=(0, 12))

        self._var_total_adultes = tk.StringVar(value="0,00")
        self._var_total_enfants = tk.StringVar(value="0,00")
        self._var_grand_total   = tk.StringVar(value="0,00")

        _total_label(footer, "Adultes :", self._var_total_adultes)
        tk.Label(footer, text="|", bg=PANEL_BG_COLOR, fg=MUTED_TEXT_COLOR).pack(side="left")
        _total_label(footer, "Enfants :", self._var_total_enfants)
        tk.Label(footer, text="|", bg=PANEL_BG_COLOR, fg=MUTED_TEXT_COLOR).pack(side="left")

        tk.Label(footer, text="TOTAL GÉNÉRAL :", font=("Poppins", 12, "bold"),
                 bg=PANEL_BG_COLOR, fg=TEXT_COLOR).pack(side="left", padx=(16, 2))
        lbl_grand = ctk.CTkLabel(footer, textvariable=self._var_grand_total,
                                 font=("Poppins", 14, "bold"), text_color=ACCENT_TEXT_COLOR)
        lbl_grand.pack(side="left", padx=(0, 4))
        tk.Label(footer, text="Ar", font=("Poppins", 11),
                 bg=PANEL_BG_COLOR, fg=MUTED_TEXT_COLOR).pack(side="left")

    # ── Données ────────────────────────────────────────────────────────────────

    def _populate_initial_rows(self):
        saved = load_client_air_ticket_cotation(self.client)
        if saved:
            self._rows = [_recompute_row(row) for row in saved]
            return
        trip_mode = self.trip_mode_var.get().strip() or "aller-retour"
        self._rows = _build_initial_rows(self.client, trip_mode)

    def _refresh_tree(self):
        self._tree.delete(*self._tree.get_children())
        for idx, row in enumerate(self._rows):
            tag = "odd" if idx % 2 else "even"
            self._tree.insert(
                "", "end", iid=str(idx),
                values=(
                    row.get("date_vol", ""),
                    row.get("numero_vol", ""),
                    row.get("type_trajet", ""),
                    row.get("compagnie", ""),
                    row.get("ville_depart", ""),
                    row.get("ville_arrivee", ""),
                    row.get("classe", ""),
                    row.get("nb_adultes", ""),
                    row.get("nb_enfants", ""),
                    _fmt(row["sous_total"]) if row.get("sous_total") else "",
                    row.get("marge_pct", ""),
                    _fmt(row["total"]) if row.get("total") else "",
                ),
                tags=(tag,),
            )

    def _refresh_totals(self):
        total_a = sum(_to_float(r.get("montant_adultes", 0)) for r in self._rows)
        total_e = sum(_to_float(r.get("montant_enfants", 0)) for r in self._rows)
        grand   = sum(_to_float(r.get("total", 0)) for r in self._rows)
        self._var_total_adultes.set(_fmt(total_a))
        self._var_total_enfants.set(_fmt(total_e))
        self._var_grand_total.set(_fmt(grand))

    def _regenerate_rows(self):
        self._rows = _build_initial_rows(self.client, self.trip_mode_var.get())
        self._refresh_tree()
        self._refresh_totals()

    # ── Actions ────────────────────────────────────────────────────────────────

    def _get_selected_index(self):
        sel = self._tree.selection()
        return int(sel[0]) if sel else None

    def _delete_selected(self):
        idx = self._get_selected_index()
        if idx is None:
            messagebox.showwarning("Sélection", "Sélectionnez une ligne à supprimer.")
            return
        self._rows.pop(idx)
        self._refresh_tree()
        self._refresh_totals()

    def _open_add_dialog(self):
        self._open_row_dialog(None)

    def _open_edit_dialog(self):
        idx = self._get_selected_index()
        if idx is None:
            messagebox.showwarning("Sélection", "Sélectionnez une ligne à modifier.")
            return
        self._open_row_dialog(idx)

    def _apply_dialog_values(self, row_index, values):
        total_value  = str(values["total"]).strip()
        total_manuel = bool(total_value)
        row = _make_row(
            date_vol=values["date_vol"],
            numero_vol=values["numero_vol"],
            type_trajet=values["type_trajet"],
            compagnie=values["compagnie"],
            ville_depart=values["ville_depart"],
            ville_arrivee=values["ville_arrivee"],
            classe=values["classe"],
            nb_adultes=values["nb_adultes"],
            nb_enfants=values["nb_enfants"],
            tarif_adulte=values["tarif_adulte"],
            tarif_enfant=values["tarif_enfant"],
            marge_pct=values["marge_pct"],
            total=total_value or 0,
            total_manuel=total_manuel,
        )
        row = _recompute_row(row)
        if row_index is None:
            self._rows.append(row)
        else:
            self._rows[row_index] = row
        self._refresh_tree()
        self._refresh_totals()

    def _save_to_excel(self):
        errors = _validate_rows(self._rows)
        if errors:
            messagebox.showwarning("Validation", "\n".join(errors))
            return
        result = save_client_air_ticket_cotation_to_excel(self.client, self._rows)
        if result > 0:
            messagebox.showinfo("Sauvegarde réussie", f"{result} ligne(s) enregistrée(s).")
        elif result == -2:
            messagebox.showerror(
                "Fichier verrouillé",
                "Le fichier Excel est ouvert ailleurs.\nFermez data.xlsx puis réessayez.",
            )
        else:
            messagebox.showerror("Erreur", "La sauvegarde a échoué. Consultez les logs.")

    def _export_pdf(self):
        if not self._rows:
            messagebox.showwarning("Aucune donnée", "Ajoutez des lignes avant de générer le PDF.")
            return
        try:
            from utils.pdf_generator import generate_air_ticket_cotation_pdf
            path = generate_air_ticket_cotation_pdf(self.client, self._rows, DEVIS_FOLDER)
            messagebox.showinfo("PDF généré", f"Fichier créé :\n{path}")
            # Ouvre le PDF avec le lecteur par défaut
            try:
                if os.name == "nt":
                    os.startfile(path)
                else:
                    subprocess.Popen(["xdg-open", path])
            except Exception:
                pass
        except ImportError:
            messagebox.showerror(
                "Module manquant",
                "ReportLab est requis pour la génération PDF.\n"
                "Installez-le avec : pip install reportlab",
            )
        except Exception as exc:
            messagebox.showerror("Erreur PDF", str(exc))

    def _go_back(self):
        if self.on_back:
            self.on_back()

    # ── Dialogue d'édition ─────────────────────────────────────────────────────

    def _open_row_dialog(self, row_index):
        existing = self._rows[row_index] if row_index is not None else None

        win = tk.Toplevel(self.parent)
        win.title("Modifier le vol" if existing else "Ajouter un vol")
        win.configure(bg=MAIN_BG_COLOR)
        win.resizable(False, False)
        win.transient(self.parent)
        win.after(0, lambda: [win.lift(), win.focus_set()])

        # ── Variables ─────────────────────────────────────────────────────────
        def _v(key, default=""):
            return tk.StringVar(value=str(existing.get(key, default) if existing else default))

        v_date    = _v("date_vol")
        v_num     = _v("numero_vol")
        v_type    = _v("type_trajet", "aller")
        v_comp    = _v("compagnie")
        v_dep     = _v("ville_depart")
        v_arr     = _v("ville_arrivee")
        v_classe  = _v("classe", "Économique")
        v_nadult  = _v("nb_adultes")
        v_nenfant = _v("nb_enfants", "0")
        v_tadult  = _v("tarif_adulte")
        v_tenfant = _v("tarif_enfant")
        v_marge   = _v("marge_pct", "0")
        v_total   = tk.StringVar(
            value=str(existing["total"]) if (existing and existing.get("total_manuel")) else ""
        )
        # Aperçu
        v_mnt_a  = tk.StringVar(value="")
        v_mnt_e  = tk.StringVar(value="")
        v_st     = tk.StringVar(value="")
        v_tc     = tk.StringVar(value="")

        # ── Helpers mise en page ──────────────────────────────────────────────
        outer = tk.Frame(win, bg=MAIN_BG_COLOR)
        outer.pack(padx=24, pady=(20, 0))

        SEC = PANEL_BG_COLOR

        def _lbl(parent, text, r, c=0):
            tk.Label(parent, text=text, font=LABEL_FONT,
                     fg=TEXT_COLOR, bg=SEC, anchor="w",
                     ).grid(row=r, column=c, sticky="w", padx=(0, 10), pady=5)

        def _entry(parent, var, r, c=1, w=28, state="normal", justify="left"):
            bg = INPUT_BG_COLOR if state == "normal" else PANEL_BG_COLOR
            e = tk.Entry(parent, textvariable=var, font=ENTRY_FONT,
                         bg=bg, fg=TEXT_COLOR, width=w, justify=justify,
                         insertbackground=TEXT_COLOR, relief="flat", state=state)
            e.grid(row=r, column=c, sticky="ew", pady=5)
            return e

        def _make_section(title):
            card = tk.Frame(outer, bg=SEC)
            card.pack(fill="x", pady=(0, 10))
            tk.Label(card, text=title, font=LABEL_FONT, fg=TEXT_COLOR,
                     bg=SEC).pack(anchor="w", padx=10, pady=(8, 2))
            inner = tk.Frame(card, bg=SEC)
            inner.pack(fill="x", padx=10, pady=(0, 8))
            inner.columnconfigure(1, weight=1)
            return inner

        # ── Section 1 : Informations du vol ───────────────────────────────────
        s1 = _make_section("Informations du vol")

        _lbl(s1, "Date du vol :", 0)
        _entry(s1, v_date, 0, w=14)
        tk.Label(s1, text="JJ/MM/AAAA", font=("Poppins", 9),
                 fg=MUTED_TEXT_COLOR, bg=SEC,
                 ).grid(row=0, column=2, sticky="w", padx=(6, 0))

        _lbl(s1, "N° de vol :", 1)
        _entry(s1, v_num, 1, w=14)

        _lbl(s1, "Type de trajet :", 2)
        ttk.Combobox(s1, textvariable=v_type,
                     values=["aller", "retour", "transit"],
                     state="readonly", font=ENTRY_FONT, width=14,
                     ).grid(row=2, column=1, sticky="w", pady=5)

        # ── Section 2 : Trajet & compagnie ────────────────────────────────────
        s2 = _make_section("Trajet & compagnie")

        _lbl(s2, "Compagnie :", 0)
        compagnies = get_avion_compagnies()
        combo_comp = ttk.Combobox(s2, textvariable=v_comp,
                                  values=compagnies, font=ENTRY_FONT, width=26)
        combo_comp.grid(row=0, column=1, sticky="ew", pady=5)

        _lbl(s2, "Départ :", 1)
        _entry(s2, v_dep, 1)

        _lbl(s2, "Arrivée :", 2)
        _entry(s2, v_arr, 2)

        _lbl(s2, "Classe :", 3)
        ttk.Combobox(s2, textvariable=v_classe,
                     values=_CLASSES, state="readonly",
                     font=ENTRY_FONT, width=16,
                     ).grid(row=3, column=1, sticky="w", pady=5)

        # ── Section 3 : Passagers & tarifs ────────────────────────────────────
        s3 = _make_section("Passagers & tarifs")

        _lbl(s3, "Nb adultes :", 0)
        _entry(s3, v_nadult, 0, w=10, justify="center")

        _lbl(s3, "Nb enfants :", 1)
        _entry(s3, v_nenfant, 1, w=10, justify="center")

        _lbl(s3, "Tarif adulte (Ar) :", 2)
        _entry(s3, v_tadult, 2, w=18, justify="right")
        tk.Label(s3, text="auto depuis BD", font=("Poppins", 9),
                 fg=MUTED_TEXT_COLOR, bg=SEC,
                 ).grid(row=2, column=2, sticky="w", padx=(6, 0))

        _lbl(s3, "Tarif enfant (Ar) :", 3)
        _entry(s3, v_tenfant, 3, w=18, justify="right")

        # ── Section 4 : Marge & total manuel ──────────────────────────────────
        s4 = _make_section("Marge & total")

        _lbl(s4, "Marge (%) :", 0)
        _entry(s4, v_marge, 0, w=10, justify="right")

        _lbl(s4, "Total manuel (Ar) :", 1)
        _entry(s4, v_total, 1, w=18, justify="right")
        tk.Label(s4, text="laisser vide = automatique", font=("Poppins", 9),
                 fg=MUTED_TEXT_COLOR, bg=SEC,
                 ).grid(row=1, column=2, sticky="w", padx=(6, 0), pady=5)

        # ── Aperçu ────────────────────────────────────────────────────────────
        prev_card = tk.Frame(outer, bg=SEC)
        prev_card.pack(fill="x", pady=(0, 10))
        tk.Label(prev_card, text="Aperçu", font=LABEL_FONT,
                 fg=TEXT_COLOR, bg=SEC).pack(anchor="w", padx=10, pady=(8, 2))
        prev_inner = tk.Frame(prev_card, bg=SEC)
        prev_inner.pack(fill="x", padx=10, pady=(0, 8))

        for r_i, (lbl_t, var) in enumerate([
            ("Adultes (Ar) :",    v_mnt_a),
            ("Enfants (Ar) :",    v_mnt_e),
            ("Sous-total (Ar) :", v_st),
            ("Total (Ar) :",      v_tc),
        ]):
            tk.Label(prev_inner, text=lbl_t, font=LABEL_FONT,
                     fg=TEXT_COLOR, bg=SEC, anchor="w",
                     ).grid(row=r_i, column=0, sticky="w", padx=(0, 10), pady=3)
            tk.Label(prev_inner, textvariable=var, font=ENTRY_FONT,
                     fg=ACCENT_TEXT_COLOR, bg=SEC,
                     ).grid(row=r_i, column=1, sticky="w", pady=3)

        # ── Callbacks ─────────────────────────────────────────────────────────
        def _update_preview(*_):
            na  = _to_float(v_nadult.get())
            ne  = _to_float(v_nenfant.get())
            ta  = _to_float(v_tadult.get())
            te  = _to_float(v_tenfant.get())
            mg  = _to_float(v_marge.get())
            ma  = na * ta
            me  = ne * te
            st  = ma + me
            tc  = st * (1 + mg / 100.0)
            ts  = v_total.get().strip()
            tot = _to_float(ts) if ts else tc
            v_mnt_a.set(_fmt(ma))
            v_mnt_e.set(_fmt(me))
            v_st.set(_fmt(st))
            v_tc.set(_fmt(tot))

        def _autofill_tarifs(*_):
            comp = v_comp.get().strip()
            if not comp:
                return
            try:
                ta, te = get_avion_tarifs({"compagnie": comp})
                if ta and not v_tadult.get().strip():
                    v_tadult.set(str(ta))
                if te and not v_tenfant.get().strip():
                    v_tenfant.set(str(te))
            except Exception:
                pass
            _update_preview()

        combo_comp.bind("<<ComboboxSelected>>", _autofill_tarifs)
        for var in (v_nadult, v_nenfant, v_tadult, v_tenfant, v_marge, v_total):
            var.trace_add("write", _update_preview)

        _update_preview()

        # ── Boutons ───────────────────────────────────────────────────────────
        btn_bar = tk.Frame(win, bg=MAIN_BG_COLOR)
        btn_bar.pack(fill="x", padx=24, pady=(8, 20))

        def _on_ok():
            self._apply_dialog_values(row_index, {
                "date_vol":      v_date.get().strip(),
                "numero_vol":    v_num.get().strip(),
                "type_trajet":   v_type.get(),
                "compagnie":     v_comp.get().strip(),
                "ville_depart":  v_dep.get().strip(),
                "ville_arrivee": v_arr.get().strip(),
                "classe":        v_classe.get(),
                "nb_adultes":    v_nadult.get().strip(),
                "nb_enfants":    v_nenfant.get().strip(),
                "tarif_adulte":  v_tadult.get().strip(),
                "tarif_enfant":  v_tenfant.get().strip(),
                "marge_pct":     v_marge.get().strip(),
                "total":         v_total.get().strip(),
            })
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
