"""Dedicated active client quote page."""

import tkinter as tk
from tkinter import messagebox, ttk

from config import (
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_ORANGE,
    BUTTON_RED,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.client_billing import apply_margin_to_quote_line, build_client_quote, convert_quote_to_invoice
from utils.excel_handler import (
    load_active_client_quote_from_excel,
    save_active_client_invoice_to_excel,
    save_active_client_quote_to_excel,
)


def _to_float(value, default=0.0):
    try:
        return float(str(value).replace(",", ".").strip() or default)
    except (TypeError, ValueError):
        return default


def _fmt(value):
    return f"{_to_float(value):,.2f}"


class ClientQuotePage:
    """Detailed quote page for one selected client."""

    def __init__(self, parent, client, on_back=None, on_open_invoice=None):
        self.parent = parent
        self.client = client or {}
        self.on_back = on_back
        self.on_open_invoice = on_open_invoice
        self.document = {}
        self.tree = None
        self.total_cost_label = None
        self.total_price_label = None
        self.total_margin_label = None

        self._build_ui()
        self._load_document()

    def _build_ui(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        root = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        root.pack(fill="both", expand=True, padx=16, pady=12)

        header = tk.Frame(root, bg=PANEL_BG_COLOR)
        header.pack(fill="x", pady=(0, 10))
        if self.on_back:
            tk.Button(
                header,
                text="← Retour",
                command=self.on_back,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=12,
                pady=6,
            ).pack(side="left", padx=12, pady=12)

        client_name = f"{self.client.get('prenom', '')} {self.client.get('nom', '')}".strip()
        tk.Label(
            header,
            text=f"DEVIS CLIENT · {client_name or 'Client'}",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=PANEL_BG_COLOR,
        ).pack(side="left", padx=(0, 12), pady=12)

        if self.client.get("numero_dossier"):
            tk.Label(
                header,
                text=f"Dossier : {self.client.get('numero_dossier')}",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=PANEL_BG_COLOR,
            ).pack(side="left", pady=12)

        actions = tk.Frame(root, bg=MAIN_BG_COLOR)
        actions.pack(fill="x", pady=(0, 10))
        for text, cmd, color in (
            ("🔄 Actualiser depuis les dépenses client", self._refresh_from_sources, BUTTON_BLUE),
            ("💾 Enregistrer le devis", self._save_quote, BUTTON_GREEN),
            ("🧾 Générer la facture", self._generate_invoice, BUTTON_ORANGE),
        ):
            tk.Button(
                actions,
                text=text,
                command=cmd,
                bg=color,
                fg="white",
                font=BUTTON_FONT,
                padx=12,
                pady=6,
            ).pack(side="left", padx=(0, 8))

        table_wrap = tk.Frame(root, bg=MAIN_BG_COLOR)
        table_wrap.pack(fill="both", expand=True)
        columns = (
            "category",
            "designation",
            "quantity",
            "cost_total",
            "margin_pct",
            "margin_amount",
            "total_price",
        )
        self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=14)
        headings = {
            "category": "Catégorie",
            "designation": "Désignation",
            "quantity": "Qté",
            "cost_total": "Coût total",
            "margin_pct": "Marge (%)",
            "margin_amount": "Marge",
            "total_price": "Prix de vente",
        }
        widths = {
            "category": 140,
            "designation": 320,
            "quantity": 60,
            "cost_total": 120,
            "margin_pct": 90,
            "margin_amount": 110,
            "total_price": 130,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda _e: self._edit_selected_line())
        scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")

        edit_row = tk.Frame(root, bg=MAIN_BG_COLOR)
        edit_row.pack(fill="x", pady=(8, 6))
        tk.Button(
            edit_row,
            text="✏️ Modifier la marge",
            command=self._edit_selected_line,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=10,
            pady=5,
        ).pack(side="left")

        totals = tk.Frame(root, bg=PANEL_BG_COLOR)
        totals.pack(fill="x")
        self.total_cost_label = tk.Label(totals, text="Coût: 0.00", font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR)
        self.total_cost_label.pack(side="left", padx=12, pady=10)
        self.total_margin_label = tk.Label(totals, text="Marge: 0.00", font=LABEL_FONT, fg=TEXT_COLOR, bg=PANEL_BG_COLOR)
        self.total_margin_label.pack(side="left", padx=12, pady=10)
        self.total_price_label = tk.Label(totals, text="Total devis: 0.00", font=TITLE_FONT, fg=BUTTON_GREEN, bg=PANEL_BG_COLOR)
        self.total_price_label.pack(side="right", padx=12, pady=10)

    def _load_document(self):
        loaded = load_active_client_quote_from_excel(self.client)
        self.document = loaded or build_client_quote(self.client)
        if self.document and self.document.get("lines") and not loaded:
            save_active_client_quote_to_excel(self.client, self.document)
        self._render_document()

    def _refresh_from_sources(self):
        self.document = build_client_quote(self.client)
        self._render_document()

    def _render_document(self):
        self.tree.delete(*self.tree.get_children())
        for index, line in enumerate(self.document.get("lines", [])):
            margin_text = "0 (verrouillé)" if not line.get("margin_editable", True) else _fmt(line.get("margin_pct", 0))
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    line.get("category", ""),
                    line.get("designation", ""),
                    line.get("quantity", 1),
                    _fmt(line.get("cost_total", 0)),
                    margin_text,
                    _fmt(line.get("margin_amount", 0)),
                    _fmt(line.get("total_price", 0)),
                ),
            )

        total_cost = sum(_to_float(line.get("cost_total", 0)) for line in self.document.get("lines", []))
        total_margin = sum(_to_float(line.get("margin_amount", 0)) for line in self.document.get("lines", []))
        total_price = sum(_to_float(line.get("total_price", 0)) for line in self.document.get("lines", []))
        self.total_cost_label.config(text=f"Coût total: {_fmt(total_cost)} Ar")
        self.total_margin_label.config(text=f"Marge totale: {_fmt(total_margin)} Ar")
        self.total_price_label.config(text=f"Total devis: {_fmt(total_price)} Ar")

    def _selected_index(self):
        selection = self.tree.selection()
        if not selection:
            return None
        return int(selection[0])

    def _edit_selected_line(self):
        index = self._selected_index()
        if index is None:
            messagebox.showwarning("Devis client", "Sélectionnez une ligne.")
            return
        line = self.document.get("lines", [])[index]
        if not line.get("margin_editable", True):
            messagebox.showinfo(
                "Restauration",
                "La marge restauration est visible mais verrouillée à 0.",
            )
            return

        win = tk.Toplevel(self.parent)
        win.title("Modifier la marge")
        win.configure(bg=MAIN_BG_COLOR)
        tk.Label(win, text=line.get("designation", ""), font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).pack(anchor="w", padx=12, pady=(12, 6))
        tk.Label(win, text="Marge (%)", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).pack(anchor="w", padx=12)
        margin_var = tk.StringVar(value=str(line.get("margin_pct", 0)))
        tk.Entry(win, textvariable=margin_var, font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR, width=24).pack(anchor="w", padx=12, pady=(0, 12))

        def _apply():
            updated = apply_margin_to_quote_line(line, margin_var.get())
            self.document["lines"][index] = updated
            self._render_document()
            win.destroy()

        btns = tk.Frame(win, bg=MAIN_BG_COLOR)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        tk.Button(btns, text="Appliquer", command=_apply, bg=BUTTON_GREEN, fg="white", font=BUTTON_FONT).pack(side="left", padx=(0, 8))
        tk.Button(btns, text="Annuler", command=win.destroy, bg=BUTTON_RED, fg="white", font=BUTTON_FONT).pack(side="left")

    def _save_quote(self):
        result = save_active_client_quote_to_excel(self.client, self.document)
        if result == -2:
            messagebox.showerror("Devis client", "Fermez data.xlsx puis réessayez.")
            return
        if result < 0:
            messagebox.showerror("Devis client", "La sauvegarde du devis a échoué.")
            return
        messagebox.showinfo("Devis client", "Devis actif enregistré.")

    def _generate_invoice(self):
        if not self.document.get("lines"):
            messagebox.showwarning("Devis client", "Aucune ligne de devis à facturer.")
            return
        invoice_document = convert_quote_to_invoice(self.document)
        result = save_active_client_invoice_to_excel(self.client, invoice_document)
        if result == -2:
            messagebox.showerror("Facture client", "Fermez data.xlsx puis réessayez.")
            return
        if result < 0:
            messagebox.showerror("Facture client", "La génération de la facture a échoué.")
            return
        if self.on_open_invoice:
            self.on_open_invoice()
            return
        messagebox.showinfo("Facture client", "Facture active générée.")
