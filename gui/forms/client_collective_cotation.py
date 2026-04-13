"""
Cotation frais collectifs par client.

Colonnes : Prestataire | Désignation | Forfait | Quantité
           | Prix unitaire (BD) | Dépense | Marge (%) | Total

Le prix unitaire est lu depuis la base FRAIS_COLLECTIFS (data-hotel.xlsx).
"""

import tkinter as tk
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
    get_collective_expense_designations,
    get_collective_expense_forfait,
    get_collective_expense_montant,
    get_collective_expense_prestataires,
    load_client_collective_cotation,
    save_client_collective_cotation_to_excel,
)

_HOVER_GREEN = "#0A6870"
_HOVER_BLUE  = "#0B6080"
_HOVER_RED   = "#A82020"
_HOVER_GREY  = "#9EA7AA"


# ── Utilitaires ───────────────────────────────────────────────────────────────

def _to_float(s, default: float = 0.0) -> float:
    try:
        return float(str(s).replace(",", ".").strip() or default)
    except (ValueError, TypeError):
        return default


def _fmt(value: float) -> str:
    return f"{value:,.2f}"


def _make_row(prestataire="", designation="", forfait="",
              quantite="", prix_unitaire="", marge="") -> dict:
    prix  = _to_float(prix_unitaire)
    qty   = _to_float(quantite, 1.0)
    mg    = _to_float(marge)
    dep   = prix * qty
    total = dep * (1 + mg / 100)
    return {
        "prestataire":  prestataire,
        "designation":  designation,
        "forfait":      forfait,
        "quantite":     quantite,
        "prix_unitaire": prix_unitaire,
        "marge":        marge,
        "depense":      dep,
        "total":        total,
    }


# ── Classe principale ─────────────────────────────────────────────────────────

class ClientCollectiveCotation:
    """Tableau de cotation frais collectifs par prestataire pour un client."""

    _COLS = [
        ("prestataire",  "Prestataire",    160),
        ("forfait",      "Forfait",         90),
        ("designation",  "Désignation",    180),
        ("quantite",     "Quantité",        70),
        ("prix_unitaire","Prix unitaire",  120),
        ("depense",      "Dépense",        120),
        ("marge",        "Marge (%)",       80),
        ("total",        "Total",          120),
    ]

    def __init__(self, parent: tk.Widget, client: dict, on_back=None):
        self.parent  = parent
        self.client  = client
        self.on_back = on_back
        self._rows: list = []

        self._prestataires: list = get_collective_expense_prestataires()

        self._build_ui()
        saved = load_client_collective_cotation(self.client)
        if saved:
            self._rows = saved
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
            self.parent,
            text="COTATION FRAIS COLLECTIFS",
            font=TITLE_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR,
        ).pack(pady=(20, 4))

        client_name = f"{prenom} {nom}".strip() or "—"
        tk.Label(
            self.parent,
            text="Client : " + client_name +
                 (f"   |   Dossier : {dossier}" if dossier else ""),
            font=ENTRY_FONT, fg=MUTED_TEXT_COLOR, bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 2))

        # Bouton retour
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

        # ── Carte Infos client ─────────────────────────────────────────────────
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
        ]):
            tk.Label(
                info_inner, text=lbl_text,
                font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
            ).grid(row=0, column=col_i * 2,
                   sticky="w", padx=(0 if col_i == 0 else 20, 4), pady=4)
            tk.Label(
                info_inner, text=val_text,
                font=ENTRY_FONT, fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR,
            ).grid(row=0, column=col_i * 2 + 1, sticky="w", pady=4)

        # ── Barre d'actions ────────────────────────────────────────────────────
        action_bar = tk.Frame(root, bg=MAIN_BG_COLOR)
        action_bar.pack(fill="x", pady=(0, 8))

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

        # ── Carte Tableau ──────────────────────────────────────────────────────
        table_card = tk.Frame(root, bg=PANEL_BG_COLOR)
        table_card.pack(fill="both", expand=True, pady=(0, 10))
        tk.Label(
            table_card, text="Lignes de frais collectifs",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(anchor="w", padx=12, pady=(8, 2))
        table_inner = tk.Frame(table_card, bg=PANEL_BG_COLOR)
        table_inner.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        style = ttk.Style()
        style.configure(
            "Coll.Treeview",
            background=INPUT_BG_COLOR, foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR, rowheight=30,
        )
        style.configure("Coll.Treeview.Heading", font=LABEL_FONT)
        style.map("Coll.Treeview", background=[("selected", BUTTON_BLUE)])

        col_ids = [c[0] for c in self._COLS]
        self._tree = ttk.Treeview(
            table_inner, columns=col_ids, show="headings",
            height=12, style="Coll.Treeview",
        )
        for key, heading, width in self._COLS:
            self._tree.heading(key, text=heading)
            anchor = "e" if key in (
                "quantite", "prix_unitaire", "depense", "marge", "total"
            ) else "w"
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

        # Tags pour les lignes alternées
        self._tree.tag_configure("odd",  background="#E4F2F6")
        self._tree.tag_configure("even", background=INPUT_BG_COLOR)

        # ── Carte Totaux ───────────────────────────────────────────────────────
        totals_card = tk.Frame(root, bg=PANEL_BG_COLOR)
        totals_card.pack(fill="x")
        tk.Label(
            totals_card, text="Totaux",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(anchor="w", padx=12, pady=(8, 2))
        totals_inner = tk.Frame(totals_card, bg=PANEL_BG_COLOR)
        totals_inner.pack(fill="x", padx=12, pady=(0, 8))

        tk.Label(
            totals_inner, text="Total dépenses :",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)

        self._lbl_dep = tk.Label(
            totals_inner, text="0.00",
            font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR,
        )
        self._lbl_dep.grid(row=0, column=1, sticky="w", padx=(0, 40), pady=4)

        tk.Label(
            totals_inner, text="Total global (avec marges) :",
            font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=4)

        self._lbl_glob = tk.Label(
            totals_inner, text="0.00",
            font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=PANEL_BG_COLOR,
        )
        self._lbl_glob.grid(row=0, column=3, sticky="w", pady=4)

    # ── Tableau ────────────────────────────────────────────────────────────────

    def _refresh_tree(self):
        self._tree.delete(*self._tree.get_children())
        for i, rd in enumerate(self._rows):
            dep   = rd["depense"]
            total = rd["total"]
            prix  = _to_float(rd["prix_unitaire"])
            tag = "odd" if i % 2 else "even"
            self._tree.insert("", "end", iid=str(i), values=(
                rd["prestataire"],
                rd["forfait"],
                rd["designation"],
                rd["quantite"],
                _fmt(prix)  if prix  else "",
                _fmt(dep)   if dep   else "",
                rd["marge"] + " %" if rd["marge"] else "",
                _fmt(total) if total else "",
            ), tags=(tag,))

    def _refresh_totals(self):
        self._lbl_dep.configure(text=_fmt(sum(r["depense"] for r in self._rows)))
        self._lbl_glob.configure(text=_fmt(sum(r["total"]   for r in self._rows)))

    # ── Sauvegarde Excel ───────────────────────────────────────────────────────

    def _save_to_excel(self):
        if not self._rows:
            messagebox.showwarning("Aucune donnée", "Le tableau est vide. Rien à sauvegarder.")
            return
        result = save_client_collective_cotation_to_excel(self.client, self._rows)
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
        pax = str(self.client.get("nombre_participants") or
                  self.client.get("nombre_adultes") or "")
        self._open_row_dialog(_make_row(quantite=pax), row_index=None)

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

        # Variables
        v_presta  = tk.StringVar(value=row["prestataire"])
        v_desig   = tk.StringVar(value=row["designation"])
        v_forfait = tk.StringVar(value=row["forfait"])
        v_qty     = tk.StringVar(value=row["quantite"])
        v_prix    = tk.StringVar(value=row["prix_unitaire"])
        v_marge   = tk.StringVar(value=row["marge"])

        outer = tk.Frame(win, bg=MAIN_BG_COLOR)
        outer.pack(padx=24, pady=(20, 0))

        SEC = PANEL_BG_COLOR

        def _make_section(title):
            card = tk.Frame(outer, bg=SEC)
            card.pack(fill="x", pady=(0, 10))
            tk.Label(card, text=title, font=LABEL_FONT, fg=TEXT_COLOR, bg=SEC,
                     ).pack(anchor="w", padx=10, pady=(8, 2))
            inner = tk.Frame(card, bg=SEC)
            inner.pack(fill="x", padx=10, pady=(0, 8))
            return inner

        def _lbl(parent, text, r, c=0, bg=None):
            tk.Label(
                parent, text=text,
                font=LABEL_FONT, fg=TEXT_COLOR, bg=bg or SEC, anchor="w",
            ).grid(row=r, column=c, sticky="w", padx=(0, 10), pady=5)

        def _entry(parent, var, r, c=1, w=32, justify="left", readonly=False):
            e = tk.Entry(
                parent, textvariable=var,
                font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                width=w, justify=justify,
                insertbackground=TEXT_COLOR, relief="flat",
                state="readonly" if readonly else "normal",
            )
            e.grid(row=r, column=c, sticky="ew", pady=5)
            return e

        # ── Section 1 : Prestataire / Désignation ─────────────────────────
        s1 = _make_section("Service")

        _lbl(s1, "Prestataire :", 0)
        combo_presta = ttk.Combobox(
            s1, textvariable=v_presta,
            values=self._prestataires, font=ENTRY_FONT, width=30, state="readonly",
        )
        combo_presta.grid(row=0, column=1, sticky="ew", pady=5)

        _lbl(s1, "Désignation :", 1)
        combo_desig = ttk.Combobox(
            s1, textvariable=v_desig,
            values=[], font=ENTRY_FONT, width=30, state="readonly",
        )
        combo_desig.grid(row=1, column=1, sticky="ew", pady=5)

        _lbl(s1, "Forfait :", 2)
        _entry(s1, v_forfait, 2, readonly=True)

        # ── Section 2 : Quantité / Prix / Marge ───────────────────────────
        s2 = _make_section("Tarification")

        _lbl(s2, "Quantité :", 0)
        _entry(s2, v_qty, 0, w=10, justify="center")

        _lbl(s2, "Prix unitaire :", 1)
        _entry(s2, v_prix, 1, w=20, justify="right")

        tk.Label(
            s2, text="(auto-rempli depuis la base, modifiable)",
            font=ENTRY_FONT, fg=MUTED_TEXT_COLOR, bg=SEC,
        ).grid(row=1, column=2, sticky="w", padx=(8, 0))

        _lbl(s2, "Marge (%) :", 2)
        _entry(s2, v_marge, 2, w=10, justify="center")

        # Preview
        lbl_prev = tk.Label(
            s2, text="Dépense : —   |   Total : —",
            font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=SEC,
        )
        lbl_prev.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 0))

        # ── Calcul temps réel ──────────────────────────────────────────────
        def _update_preview(*_):
            prix  = _to_float(v_prix.get())
            qty   = _to_float(v_qty.get(), 1.0)
            marge = _to_float(v_marge.get())
            dep   = prix * qty
            total = dep * (1 + marge / 100)
            lbl_prev.configure(
                text=f"Dépense : {_fmt(dep)}   |   Total : {_fmt(total)}"
            )

        for var in (v_qty, v_prix, v_marge):
            var.trace_add("write", _update_preview)

        # ── Cascade prestataire → désignation → forfait/prix ──────────────
        def _on_prestataire(*_):
            presta = v_presta.get()
            desigs = get_collective_expense_designations(presta) if presta else []
            combo_desig["values"] = desigs
            v_desig.set("")
            v_forfait.set("")
            v_prix.set("")
            _update_preview()

        def _on_designation(*_):
            presta = v_presta.get()
            desig  = v_desig.get()
            if presta and desig:
                forfait = get_collective_expense_forfait(presta, desig)
                montant = get_collective_expense_montant(presta, desig)
                v_forfait.set(forfait)
                v_prix.set(f"{float(montant):.2f}" if montant else "")
            _update_preview()

        combo_presta.bind("<<ComboboxSelected>>", _on_prestataire)
        combo_desig.bind("<<ComboboxSelected>>",  _on_designation)

        if row["prestataire"]:
            combo_desig["values"] = get_collective_expense_designations(row["prestataire"])

        _update_preview()

        # ── Boutons ────────────────────────────────────────────────────────
        btn_bar = tk.Frame(win, bg=MAIN_BG_COLOR)
        btn_bar.pack(pady=(8, 20), padx=24, fill="x")

        def _save():
            new_row = _make_row(
                prestataire  = v_presta.get().strip(),
                designation  = v_desig.get().strip(),
                forfait      = v_forfait.get().strip(),
                quantite     = v_qty.get().strip(),
                prix_unitaire= v_prix.get().strip(),
                marge        = v_marge.get().strip(),
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
