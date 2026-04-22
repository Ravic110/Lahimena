"""Dedicated active client invoice page."""

import os
import subprocess
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

from config import (
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_RED,
    DEVIS_FOLDER,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.client_billing import convert_quote_to_invoice
from utils.excel_handler import (
    load_active_client_invoice_from_excel,
    load_active_client_quote_from_excel,
    save_active_client_invoice_to_excel,
)
from utils.pdf_generator import REPORTLAB_AVAILABLE, generate_invoice_pdf


def _to_float(value, default=0.0):
    try:
        return float(str(value).replace(",", ".").strip() or default)
    except (TypeError, ValueError):
        return default


def _fmt(value):
    return f"{_to_float(value):,.2f}"


class ClientInvoicePage:
    """Simplified invoice page for one selected client."""

    def __init__(self, parent, client, on_back=None):
        self.parent = parent
        self.client = client or {}
        self.on_back = on_back
        self.document = {}
        self.tree = None
        self.total_label = None

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
            text=f"FACTURE CLIENT · {client_name or 'Client'}",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=PANEL_BG_COLOR,
        ).pack(side="left", padx=(0, 12), pady=12)

        actions = tk.Frame(root, bg=MAIN_BG_COLOR)
        actions.pack(fill="x", pady=(0, 10))
        for text, cmd, color in (
            ("🔁 Régénérer depuis le devis actif", self._rebuild_from_quote, BUTTON_BLUE),
            ("💾 Enregistrer la facture", self._save_invoice, BUTTON_GREEN),
            ("📄 Générer la facture en PDF", self._generate_pdf, BUTTON_GREEN),
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
        columns = ("category", "designation", "quantity", "unit_price", "total_price")
        self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=12)
        headings = {
            "category": "Catégorie",
            "designation": "Désignation",
            "quantity": "Qté",
            "unit_price": "Prix unitaire",
            "total_price": "Total",
        }
        widths = {
            "category": 160,
            "designation": 260,
            "quantity": 60,
            "unit_price": 130,
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

        tk.Button(
            root,
            text="✏️ Modifier la ligne",
            command=self._edit_selected_line,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=10,
            pady=5,
        ).pack(anchor="w", pady=(8, 6))

        footer = tk.Frame(root, bg=PANEL_BG_COLOR)
        footer.pack(fill="x")
        self.total_label = tk.Label(footer, text="Total facture: 0.00", font=TITLE_FONT, fg=BUTTON_GREEN, bg=PANEL_BG_COLOR)
        self.total_label.pack(side="right", padx=12, pady=10)

    def _load_document(self):
        self.document = load_active_client_invoice_from_excel(self.client)
        if not self.document:
            quote_document = load_active_client_quote_from_excel(self.client)
            if quote_document:
                self.document = convert_quote_to_invoice(quote_document)
                save_active_client_invoice_to_excel(self.client, self.document)
        self._render_document()

    def _render_document(self):
        self.tree.delete(*self.tree.get_children())
        for index, line in enumerate(self.document.get("lines", [])):
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    line.get("category", ""),
                    line.get("designation", ""),
                    line.get("quantity", 1),
                    _fmt(line.get("unit_price", 0)),
                    _fmt(line.get("total_price", 0)),
                ),
            )
        total = sum(_to_float(line.get("total_price", 0)) for line in self.document.get("lines", []))
        self.total_label.config(text=f"Total facture: {_fmt(total)} Ar")

    def _selected_index(self):
        selection = self.tree.selection()
        if not selection:
            return None
        return int(selection[0])

    def _edit_selected_line(self):
        index = self._selected_index()
        if index is None:
            messagebox.showwarning("Facture client", "Sélectionnez une ligne.")
            return
        line = self.document.get("lines", [])[index]

        win = tk.Toplevel(self.parent)
        win.title("Modifier la ligne facture")
        win.configure(bg=MAIN_BG_COLOR)
        tk.Label(win, text=line.get("designation", ""), font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).pack(anchor="w", padx=12, pady=(12, 6))
        tk.Label(win, text="Prix unitaire", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).pack(anchor="w", padx=12)
        unit_var = tk.StringVar(value=str(line.get("unit_price", 0)))
        tk.Entry(win, textvariable=unit_var, font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR, width=24).pack(anchor="w", padx=12, pady=(0, 12))

        def _apply():
            unit_price = max(0.0, _to_float(unit_var.get()))
            quantity = max(1, int(_to_float(line.get("quantity", 1), 1)))
            self.document["lines"][index] = {
                **line,
                "unit_price": unit_price,
                "total_price": unit_price * quantity,
            }
            self._render_document()
            win.destroy()

        btns = tk.Frame(win, bg=MAIN_BG_COLOR)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        tk.Button(btns, text="Appliquer", command=_apply, bg=BUTTON_GREEN, fg="white", font=BUTTON_FONT).pack(side="left", padx=(0, 8))
        tk.Button(btns, text="Annuler", command=win.destroy, bg=BUTTON_RED, fg="white", font=BUTTON_FONT).pack(side="left")

    def _rebuild_from_quote(self):
        quote_document = load_active_client_quote_from_excel(self.client)
        if not quote_document:
            messagebox.showwarning("Facture client", "Aucun devis actif disponible.")
            return
        self.document = convert_quote_to_invoice(quote_document)
        self._render_document()

    def _save_invoice(self):
        result = save_active_client_invoice_to_excel(self.client, self.document)
        if result == -2:
            messagebox.showerror("Facture client", "Fermez data.xlsx puis réessayez.")
            return
        if result < 0:
            messagebox.showerror("Facture client", "La sauvegarde a échoué.")
            return
        messagebox.showinfo("Facture client", "Facture active enregistrée.")

    def _generate_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("PDF", "ReportLab n'est pas disponible pour générer le PDF.")
            return
        if not self.document.get("lines"):
            messagebox.showwarning("PDF", "Aucune ligne de facture à exporter.")
            return

        total = sum(_to_float(line.get("total_price", 0)) for line in self.document.get("lines", []))
        invoice_number = f"FAC-{self.client.get('ref_client', 'CLIENT')}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        items = [
            {
                "designation": line.get("designation", line.get("category", "")),
                "nights": max(1, int(_to_float(line.get("quantity", 1), 1))),
                "unit_price": _to_float(line.get("unit_price", 0)),
                "total": _to_float(line.get("total_price", 0)),
            }
            for line in self.document.get("lines", [])
        ]

        path = generate_invoice_pdf(
            invoice_number=invoice_number,
            invoice_date=datetime.now().strftime("%Y-%m-%d"),
            client_name=self.document.get(
                "client_name",
                f"{self.client.get('prenom', '')} {self.client.get('nom', '')}".strip(),
            ),
            client_email=self.client.get("email", ""),
            client_phone=self.client.get("telephone", ""),
            source_type="Client",
            source_ref=self.client.get("ref_client", ""),
            currency=self.document.get("currency", "Ariary"),
            montant_ht=total,
            marge_pct=0,
            marge_amount=0,
            tva_pct=0,
            tva_amount=0,
            total_ttc=total,
            acompte=0,
            reste_a_payer=total,
            statut="Non payée",
            items=items,
            base_taxable_ht=total,
            output_dir=DEVIS_FOLDER,
        )
        messagebox.showinfo("PDF", f"Facture PDF générée:\n{path}")
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception:
            pass
