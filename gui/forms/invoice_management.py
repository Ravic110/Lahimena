"""Invoice management GUI component."""

import os
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk

from config import (
    ACCENT_TEXT_COLOR,
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.excel_handler import (
    INVOICE_STATUS_PAID,
    INVOICE_STATUS_PARTIAL,
    INVOICE_STATUS_UNPAID,
    calculate_invoice_totals,
    load_all_air_ticket_quotations,
    load_all_clients,
    load_all_collective_expense_quotations,
    load_all_hotel_quotations,
    load_all_invoices,
    load_all_transport_quotations,
    load_all_visite_excursion_quotations,
    load_financial_state_snapshot,
    refresh_financial_state_from_invoices,
    save_invoice_to_excel,
    update_invoice_in_excel,
)
from utils.pdf_generator import REPORTLAB_AVAILABLE, generate_invoice_pdf


class InvoiceManagement:
    """Create and track invoices with VAT, margins, deposits and statuses."""

    SOURCE_TYPES = ["Client"]

    def __init__(self, parent, on_back_to_hub=None):
        self.parent = parent
        self.on_back_to_hub = on_back_to_hub
        self.source_rows = []
        self.invoices = []
        self.source_map = {}
        self.invoice_tree = None
        self.acompte_entry = None
        self.acompte_pct_entry = None
        self._syncing_acompte_fields = False
        self.preview_labels = {}

        self.source_type_var = tk.StringVar(value=self.SOURCE_TYPES[0])
        self.source_ref_var = tk.StringVar()
        self.montant_ht_var = tk.StringVar(value="0")
        self.cout_ht_var = tk.StringVar(value="0")
        self.marge_pct_var = tk.StringVar(value="0")
        self.tva_pct_var = tk.StringVar(value="20")
        self.acompte_pct_var = tk.StringVar(value="0")
        self.acompte_var = tk.StringVar(value="0")
        self.statut_var = tk.StringVar(value=INVOICE_STATUS_UNPAID)

        self._build_ui()
        self._refresh_all()

    def _build_ui(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        tk.Label(
            self.parent,
            text="FACTURATION CLIENTS",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(20, 10), fill="x")
        tk.Label(
            self.parent,
            text="Sélectionnez une source puis ajustez marge/TVA/acompte. Le total et le reste à payer se mettent à jour automatiquement.",
            font=ENTRY_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 8), fill="x")

        if self.on_back_to_hub:
            top_actions = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
            top_actions.pack(fill="x", padx=16, pady=(0, 8))
            tk.Button(
                top_actions,
                text="⬅ Retour vers Factures / Devis",
                command=self._go_back_to_hub,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=12,
                pady=5,
            ).pack(side="left")

        root = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        root.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        form = tk.LabelFrame(
            root,
            text="Nouvelle facture",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        form.pack(fill="x", pady=(0, 10))

        self._field(form, 0, "Source", self._source_type_combo(form))
        self._field(form, 1, "Référence", self._source_ref_combo(form))
        self._field(form, 2, "Montant HT", self._entry(form, self.montant_ht_var))
        self._field(form, 3, "Coût HT", self._entry(form, self.cout_ht_var))
        self._field(form, 4, "Marge (%)", self._entry(form, self.marge_pct_var))
        self._field(form, 5, "TVA (%)", self._entry(form, self.tva_pct_var))
        self.acompte_pct_entry = self._entry(form, self.acompte_pct_var)
        self._field(form, 6, "Acompte (%)", self.acompte_pct_entry)
        self.acompte_entry = self._entry(form, self.acompte_var)
        self._field(form, 7, "Acompte", self.acompte_entry)
        self._field(form, 8, "Statut", self._status_combo(form))

        actions = tk.Frame(form, bg=MAIN_BG_COLOR)
        actions.grid(row=9, column=0, columnspan=2, sticky="w", pady=(10, 0))

        tk.Button(
            actions,
            text="➕ Générer facture",
            command=self._create_invoice,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            actions,
            text="💾 Mettre à jour statut/acompte",
            command=self._update_selected_invoice,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            actions,
            text="🔄 Rafraîchir",
            command=self._refresh_all,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left")

        tk.Button(
            actions,
            text="📄 Générer PDF",
            command=self._generate_selected_invoice_pdf,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left", padx=(8, 0))

        tk.Button(
            actions,
            text="🧹 Réinitialiser",
            command=self._reset_form,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left", padx=(8, 0))

        preview = tk.Frame(form, bg=INPUT_BG_COLOR, bd=1, relief="ridge")
        preview.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        self.preview_labels["base"] = tk.Label(
            preview, text="Base taxable HT: 0.00", font=LABEL_FONT, fg=TEXT_COLOR, bg=INPUT_BG_COLOR
        )
        self.preview_labels["base"].grid(row=0, column=0, sticky="w", padx=8, pady=(6, 2))
        self.preview_labels["tva"] = tk.Label(
            preview, text="TVA: 0.00", font=LABEL_FONT, fg=TEXT_COLOR, bg=INPUT_BG_COLOR
        )
        self.preview_labels["tva"].grid(row=0, column=1, sticky="w", padx=8, pady=(6, 2))
        self.preview_labels["ttc"] = tk.Label(
            preview, text="Total TTC: 0.00", font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=INPUT_BG_COLOR
        )
        self.preview_labels["ttc"].grid(row=1, column=0, sticky="w", padx=8, pady=(0, 6))
        self.preview_labels["reste"] = tk.Label(
            preview, text="Reste à payer: 0.00", font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=INPUT_BG_COLOR
        )
        self.preview_labels["reste"].grid(row=1, column=1, sticky="w", padx=8, pady=(0, 6))

        self.state_frame = tk.Frame(root, bg=MAIN_BG_COLOR)
        self.state_frame.pack(fill="x", pady=(0, 10))

        table = tk.LabelFrame(
            root,
            text="Factures",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=8,
            pady=8,
        )
        table.pack(fill="both", expand=True)

        cols = (
            "id",
            "date",
            "source",
            "client",
            "ht",
            "tva",
            "ttc",
            "acompte",
            "reste",
            "statut",
        )
        self.invoice_tree = ttk.Treeview(table, columns=cols, show="headings", selectmode="browse")
        headings = {
            "id": "Facture",
            "date": "Date",
            "source": "Source",
            "client": "Client",
            "ht": "HT",
            "tva": "TVA",
            "ttc": "TTC",
            "acompte": "Acompte",
            "reste": "Reste",
            "statut": "Statut",
        }
        for key in cols:
            self.invoice_tree.heading(key, text=headings[key])
            self.invoice_tree.column(key, width=125, anchor="w")
        self.invoice_tree.column("id", width=150)
        self.invoice_tree.column("date", width=135)
        self.invoice_tree.column("source", width=135)
        self.invoice_tree.column("client", width=180)

        style = ttk.Style()
        style.configure(
            "Treeview",
            background=INPUT_BG_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR,
        )

        scroll = ttk.Scrollbar(table, orient="vertical", command=self.invoice_tree.yview)
        self.invoice_tree.configure(yscrollcommand=scroll.set)
        self.invoice_tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.invoice_tree.bind("<<TreeviewSelect>>", self._on_invoice_selected)
        self._bind_preview_updates()

    def _field(self, parent, row, label, widget):
        tk.Label(
            parent,
            text=f"{label} :",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        widget.grid(row=row, column=1, sticky="w", pady=4)

    def _entry(self, parent, variable):
        return tk.Entry(parent, textvariable=variable, font=ENTRY_FONT, width=24, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)

    def _source_type_combo(self, parent):
        combo = ttk.Combobox(
            parent,
            textvariable=self.source_type_var,
            values=self.SOURCE_TYPES,
            font=ENTRY_FONT,
            width=22,
            state="readonly",
        )
        combo.bind("<<ComboboxSelected>>", self._on_source_type_changed)
        return combo

    def _source_ref_combo(self, parent):
        combo = ttk.Combobox(
            parent,
            textvariable=self.source_ref_var,
            values=[],
            font=ENTRY_FONT,
            width=42,
            state="readonly",
        )
        combo.bind("<<ComboboxSelected>>", self._on_source_ref_changed)
        self.source_ref_combo = combo
        return combo

    def _status_combo(self, parent):
        combo = ttk.Combobox(
            parent,
            textvariable=self.statut_var,
            values=[INVOICE_STATUS_UNPAID, INVOICE_STATUS_PARTIAL, INVOICE_STATUS_PAID],
            font=ENTRY_FONT,
            width=24,
            state="readonly",
        )
        combo.bind("<<ComboboxSelected>>", self._on_status_changed)
        return combo

    def _go_back_to_hub(self):
        if self.on_back_to_hub:
            self.on_back_to_hub()

    def _bind_preview_updates(self):
        for var in (self.montant_ht_var, self.cout_ht_var, self.marge_pct_var, self.tva_pct_var):
            var.trace_add("write", lambda *_args: self._refresh_preview())
        self.acompte_pct_var.trace_add("write", lambda *_args: self._on_acompte_pct_changed())
        self.acompte_var.trace_add("write", lambda *_args: self._on_acompte_amount_changed())
        self._apply_acompte_lock()
        self._refresh_preview()

    def _apply_acompte_lock(self):
        if not self.acompte_entry or not self.acompte_pct_entry:
            return
        status = self.statut_var.get()
        if status in (INVOICE_STATUS_UNPAID, INVOICE_STATUS_PAID):
            self.acompte_pct_entry.config(state="disabled")
            self.acompte_entry.config(state="disabled")
        else:
            self.acompte_pct_entry.config(state="normal")
            self.acompte_entry.config(state="normal")

    def _on_status_changed(self, _event=None):
        self._apply_acompte_lock()
        self._refresh_preview()

    def _current_total_ttc(self):
        totals = calculate_invoice_totals(
            montant_ht=self.montant_ht_var.get(),
            cout_ht=self.cout_ht_var.get(),
            marge_pct=self.marge_pct_var.get(),
            tva_pct=self.tva_pct_var.get(),
            acompte=0,
            statut=INVOICE_STATUS_PARTIAL,
        )
        return self._to_number(totals.get("Total_TTC", 0))

    def _on_acompte_pct_changed(self):
        if self._syncing_acompte_fields:
            return
        self._syncing_acompte_fields = True
        try:
            total_ttc = self._current_total_ttc()
            pct = max(0.0, self._to_number(self.acompte_pct_var.get()))
            if pct > 100:
                pct = 100.0
            amount = (total_ttc * pct) / 100.0 if total_ttc > 0 else 0.0
            self.acompte_var.set(f"{amount:.2f}")
        finally:
            self._syncing_acompte_fields = False
        self._refresh_preview()

    def _on_acompte_amount_changed(self):
        if self._syncing_acompte_fields:
            return
        self._syncing_acompte_fields = True
        try:
            total_ttc = self._current_total_ttc()
            amount = max(0.0, self._to_number(self.acompte_var.get()))
            if total_ttc > 0:
                pct = (amount / total_ttc) * 100.0
            else:
                pct = 0.0
            if pct > 100:
                pct = 100.0
            self.acompte_pct_var.set(f"{pct:.2f}")
        finally:
            self._syncing_acompte_fields = False
        self._refresh_preview()

    def _refresh_preview(self):
        try:
            totals = calculate_invoice_totals(
                montant_ht=self.montant_ht_var.get(),
                cout_ht=self.cout_ht_var.get(),
                marge_pct=self.marge_pct_var.get(),
                tva_pct=self.tva_pct_var.get(),
                acompte=self.acompte_var.get(),
                statut=self.statut_var.get(),
            )
        except Exception:
            return

        status = self.statut_var.get()
        if status == INVOICE_STATUS_UNPAID and self._to_number(self.acompte_var.get()) != 0:
            self.acompte_var.set("0")
            return
        if status == INVOICE_STATUS_PAID:
            expected = self._to_number(totals.get("Total_TTC", 0))
            if abs(self._to_number(self.acompte_var.get()) - expected) > 0.005:
                self.acompte_var.set(f"{expected:.2f}")
                return

        self.preview_labels["base"].config(
            text=f"Base taxable HT: {self._to_number(totals.get('Base_Taxable_HT', 0)):,.2f}"
        )
        self.preview_labels["tva"].config(
            text=f"TVA: {self._to_number(totals.get('TVA_Montant', 0)):,.2f}"
        )
        self.preview_labels["ttc"].config(
            text=f"Total TTC: {self._to_number(totals.get('Total_TTC', 0)):,.2f}"
        )
        self.preview_labels["reste"].config(
            text=f"Reste à payer: {self._to_number(totals.get('Reste_A_Payer', 0)):,.2f}"
        )

    def _to_number(self, value):
        try:
            text = str(value or "").replace(" ", "").replace(",", ".")
            cleaned = ""
            used_dot = False
            for char in text:
                if char.isdigit() or char == "-":
                    cleaned += char
                elif char == "." and not used_dot:
                    cleaned += char
                    used_dot = True
            return float(cleaned) if cleaned not in ("", "-", ".") else 0.0
        except Exception:
            return 0.0

    def _normalize(self, value):
        return str(value or "").strip().lower()

    def _find_first_number(self, data, keywords):
        for key, value in data.items():
            norm = self._normalize(key)
            if any(keyword in norm for keyword in keywords):
                number = self._to_number(value)
                if number > 0:
                    return number
        return 0.0

    def _display_source(self, row):
        source_type = row.get("source_type", "")
        row_number = row.get("row_number", "")
        client = row.get("client", "")
        total = row.get("montant_ht", 0)
        if source_type == "Client":
            client_id = row.get("client_id", "")
            if client_id:
                return f"{client_id} - {client} - {total:,.2f}"
            return f"{client} - {total:,.2f}"
        return f"{source_type}#{row_number} - {client} - {total:,.2f}"

    def _rows_for_source_type(self, source_type):
        rows = []

        if source_type == "Hôtel":
            for row in load_all_hotel_quotations():
                rows.append(
                    {
                        "source_type": source_type,
                        "row_number": row.get("row_number"),
                        "source_ref": f"{source_type}#{row.get('row_number')}",
                        "client_id": row.get("client_id", ""),
                        "client": row.get("client_name", ""),
                        "devise": row.get("currency", "Ariary"),
                        "montant_ht": self._to_number(row.get("total_price", 0)),
                        "cout_ht": 0.0,
                    }
                )
        elif source_type == "Frais collectifs":
            for row in load_all_collective_expense_quotations():
                rows.append(
                    {
                        "source_type": source_type,
                        "row_number": row.get("row_number"),
                        "source_ref": f"{source_type}#{row.get('row_number')}",
                        "client_id": row.get("ID_CLIENT", ""),
                        "client": row.get("Nom", ""),
                        "devise": row.get("Devise", "Ariary"),
                        "montant_ht": self._to_number(row.get("Total", 0)),
                        "cout_ht": self._to_number(row.get("Montant", 0)),
                    }
                )
        elif source_type == "Visite & Excursion":
            for row in load_all_visite_excursion_quotations():
                rows.append(
                    {
                        "source_type": source_type,
                        "row_number": row.get("row_number"),
                        "source_ref": f"{source_type}#{row.get('row_number')}",
                        "client_id": row.get("ID_CLIENT", "") or row.get("Référence", ""),
                        "client": row.get("Nom", ""),
                        "devise": row.get("Devise", "Ariary"),
                        "montant_ht": self._to_number(row.get("Total", 0)),
                        "cout_ht": self._to_number(row.get("Montant", 0)),
                    }
                )
        elif source_type == "Billet avion":
            for row in load_all_air_ticket_quotations():
                total = self._find_first_number(row, ["total"])
                cout = self._find_first_number(row, ["montant adultes", "montant enfants", "cout", "coût"])
                client_id = row.get("ID_CLIENT", "") or row.get("ID", "") or row.get("Ref", "")
                client_name = row.get("Nom", "") or row.get("Client", "")
                rows.append(
                    {
                        "source_type": source_type,
                        "row_number": row.get("row_number"),
                        "source_ref": f"{source_type}#{row.get('row_number')}",
                        "client_id": client_id,
                        "client": client_name,
                        "devise": row.get("Devise", "Ariary"),
                        "montant_ht": total,
                        "cout_ht": cout,
                    }
                )
        else:
            for row in load_all_transport_quotations():
                total = self._find_first_number(row, ["budget", "total", "montant"])
                cout = self._find_first_number(row, ["carburant", "cout", "coût"])
                client_id = row.get("ID_CLIENT", "") or row.get("ID", "") or row.get("Référence", "")
                client_name = row.get("Nom", "") or row.get("Client", "")
                rows.append(
                    {
                        "source_type": source_type,
                        "row_number": row.get("row_number"),
                        "source_ref": f"{source_type}#{row.get('row_number')}",
                        "client_id": client_id,
                        "client": client_name,
                        "devise": row.get("Devise", "Ariary"),
                        "montant_ht": total,
                        "cout_ht": cout,
                    }
                )
        return rows

    def _build_client_source_rows(self):
        source_rows = []
        for source_type in (
            "Hôtel",
            "Frais collectifs",
            "Visite & Excursion",
            "Billet avion",
            "Transport",
        ):
            source_rows.extend(self._rows_for_source_type(source_type))

        grouped = {}
        for row in source_rows:
            montant = self._to_number(row.get("montant_ht", 0))
            if montant <= 0:
                continue

            client_id = str(row.get("client_id") or "").strip()
            client_name = str(row.get("client") or "").strip() or "Client inconnu"
            key = client_id if client_id else f"name:{self._normalize(client_name)}"
            if key not in grouped:
                grouped[key] = {
                    "source_type": "Client",
                    "row_number": "",
                    "source_ref": f"Client#{client_id}" if client_id else f"Client#{client_name}",
                    "client_id": client_id,
                    "client": client_name,
                    "devise": str(row.get("devise") or "Ariary"),
                    "montant_ht": 0.0,
                    "cout_ht": 0.0,
                }
            grouped[key]["montant_ht"] += montant
            grouped[key]["cout_ht"] += self._to_number(row.get("cout_ht", 0))

        return list(grouped.values())

    def _load_source_rows(self):
        source_type = self.source_type_var.get()
        if source_type == "Client":
            rows = self._build_client_source_rows()
        else:
            rows = self._rows_for_source_type(source_type)

        self.source_rows = [row for row in rows if self._to_number(row.get("montant_ht", 0)) > 0]
        self.source_map = {self._display_source(row): row for row in self.source_rows}
        refs = sorted(self.source_map.keys())
        self.source_ref_combo["values"] = refs
        if refs:
            self.source_ref_var.set(refs[0])
            self._apply_source(self.source_map[refs[0]])
        else:
            self.source_ref_var.set("")
            self.montant_ht_var.set("0")
            self.cout_ht_var.set("0")

    def _apply_source(self, source_row):
        self.montant_ht_var.set(f"{self._to_number(source_row.get('montant_ht', 0)):.2f}")
        self.cout_ht_var.set(f"{self._to_number(source_row.get('cout_ht', 0)):.2f}")

    def _on_source_type_changed(self, _event=None):
        self._load_source_rows()

    def _on_source_ref_changed(self, _event=None):
        selected = self.source_ref_var.get()
        source_row = self.source_map.get(selected)
        if source_row:
            self._apply_source(source_row)

    def _create_invoice(self):
        selected = self.source_ref_var.get()
        source_row = self.source_map.get(selected)
        if not source_row:
            messagebox.showwarning("Source", "Sélectionnez une ligne source pour générer la facture.")
            return
        if not self._validate_invoice_fields():
            return

        payload = {
            "Source_Type": source_row.get("source_type", ""),
            "Source_Ref": source_row.get("source_ref", ""),
            "Client_ID": source_row.get("client_id", ""),
            "Client_Nom": source_row.get("client", ""),
            "Devise": source_row.get("devise", "Ariary"),
            "Montant_HT": self.montant_ht_var.get(),
            "Cout_HT": self.cout_ht_var.get(),
            "Marge_%": self.marge_pct_var.get(),
            "TVA_%": self.tva_pct_var.get(),
            "Acompte": self.acompte_var.get(),
            "Statut": self.statut_var.get(),
        }

        row = save_invoice_to_excel(payload)
        if row == -2:
            messagebox.showerror("Erreur", "Impossible d'écrire dans Excel. Fermez data.xlsx puis réessayez.")
            return
        if row < 0:
            messagebox.showerror("Erreur", "La facture n'a pas pu être créée.")
            return

        messagebox.showinfo("Succès", "Facture créée et état financier mis à jour automatiquement.")
        self._refresh_all()

    def _load_invoices(self):
        self.invoices = load_all_invoices()

    def _render_invoices(self):
        for item in self.invoice_tree.get_children():
            self.invoice_tree.delete(item)

        for row in self.invoices:
            values = (
                row.get("ID_Facture", ""),
                row.get("Date", ""),
                row.get("Source_Type", ""),
                row.get("Client_Nom", ""),
                f"{self._to_number(row.get('Montant_HT', 0)):,.2f}",
                f"{self._to_number(row.get('TVA_Montant', 0)):,.2f}",
                f"{self._to_number(row.get('Total_TTC', 0)):,.2f}",
                f"{self._to_number(row.get('Acompte', 0)):,.2f}",
                f"{self._to_number(row.get('Reste_A_Payer', 0)):,.2f}",
                row.get("Statut", ""),
            )
            iid = str(row.get("row_number", ""))
            self.invoice_tree.insert("", "end", iid=iid if iid else None, values=values)

    def _render_financial_state(self):
        for widget in self.state_frame.winfo_children():
            widget.destroy()

        state = load_financial_state_snapshot()
        if not state:
            return

        card = tk.Frame(self.state_frame, bg=INPUT_BG_COLOR, bd=2, relief="ridge")
        card.pack(fill="x")

        left = (
            f"Factures: {int(self._to_number(state.get('Nb_Factures', 0)))}"
            f" | Payées: {int(self._to_number(state.get('Nb_Payees', 0)))}"
            f" | Acompte: {int(self._to_number(state.get('Nb_Payees_Avec_Acompte', 0)))}"
            f" | Non payées: {int(self._to_number(state.get('Nb_Non_Payees', 0)))}"
        )
        right = (
            f"CA TTC: {self._to_number(state.get('CA_TTC', 0)):,.2f}"
            f" | Acomptes reçus: {self._to_number(state.get('Acomptes_Recus', 0)):,.2f}"
            f" | Reste: {self._to_number(state.get('Restes_A_Encaisser', 0)):,.2f}"
        )

        tk.Label(card, text=left, font=LABEL_FONT, fg=TEXT_COLOR, bg=INPUT_BG_COLOR).pack(anchor="w", padx=10, pady=(8, 2))
        tk.Label(card, text=right, font=LABEL_FONT, fg=ACCENT_TEXT_COLOR, bg=INPUT_BG_COLOR).pack(anchor="w", padx=10, pady=(0, 8))

    def _on_invoice_selected(self, _event=None):
        selection = self.invoice_tree.selection()
        if not selection:
            return
        row_number = int(selection[0])
        selected = next((row for row in self.invoices if row.get("row_number") == row_number), None)
        if not selected:
            return

        self.montant_ht_var.set(str(selected.get("Montant_HT", 0)))
        self.cout_ht_var.set(str(selected.get("Cout_HT", 0)))
        self.marge_pct_var.set(str(selected.get("Marge_%", 0)))
        self.tva_pct_var.set(str(selected.get("TVA_%", 0)))
        self.acompte_var.set(str(selected.get("Acompte", 0)))
        self.statut_var.set(selected.get("Statut", INVOICE_STATUS_UNPAID))
        self._apply_acompte_lock()
        self._refresh_preview()

    def _update_selected_invoice(self):
        selection = self.invoice_tree.selection()
        if not selection:
            messagebox.showwarning("Facture", "Sélectionnez une facture à mettre à jour.")
            return
        if not self._validate_invoice_fields():
            return

        row_number = int(selection[0])
        data = {
            "Montant_HT": self.montant_ht_var.get(),
            "Cout_HT": self.cout_ht_var.get(),
            "Marge_%": self.marge_pct_var.get(),
            "TVA_%": self.tva_pct_var.get(),
            "Acompte": self.acompte_var.get(),
            "Statut": self.statut_var.get(),
        }

        check = calculate_invoice_totals(
            montant_ht=data["Montant_HT"],
            cout_ht=data["Cout_HT"],
            marge_pct=data["Marge_%"],
            tva_pct=data["TVA_%"],
            acompte=data["Acompte"],
            statut=data["Statut"],
        )
        data.update(check)

        result = update_invoice_in_excel(row_number, data)
        if result == -2:
            messagebox.showerror("Erreur", "Impossible d'écrire dans Excel. Fermez data.xlsx puis réessayez.")
            return
        if result < 0:
            messagebox.showerror("Erreur", "La mise à jour de la facture a échoué.")
            return

        messagebox.showinfo("Succès", "Facture mise à jour et état financier recalculé.")
        self._refresh_all()

    def _generate_selected_invoice_pdf(self):
        selection = self.invoice_tree.selection()
        if not selection:
            messagebox.showwarning("Facture", "Sélectionnez une facture à exporter en PDF.")
            return

        if not REPORTLAB_AVAILABLE:
            messagebox.showwarning(
                "PDF indisponible",
                "ReportLab n'est pas installé. Installez-le avec: pip install reportlab",
            )
            return

        row_number = int(selection[0])
        invoice = next(
            (row for row in self.invoices if row.get("row_number") == row_number), None
        )
        if not invoice:
            messagebox.showerror("Erreur", "Impossible de récupérer la facture sélectionnée.")
            return

        invoice_number = str(invoice.get("ID_Facture") or f"FACTURE_{row_number}")
        invoice_number = invoice_number.replace("/", "-").replace("\\", "-").strip()
        client_name = str(invoice.get("Client_Nom") or "Client")
        invoice_date = str(invoice.get("Date") or "")
        source_type = str(invoice.get("Source_Type") or "")
        source_ref = str(invoice.get("Source_Ref") or "")
        currency = str(invoice.get("Devise") or "Ariary")
        client_email, client_phone = self._resolve_client_contact(invoice)
        line_items = self._build_invoice_pdf_items(invoice)

        try:
            filepath = generate_invoice_pdf(
                invoice_number=invoice_number,
                invoice_date=invoice_date,
                client_name=client_name,
                client_email=client_email,
                client_phone=client_phone,
                source_type=source_type,
                source_ref=source_ref,
                currency=currency,
                montant_ht=self._to_number(invoice.get("Montant_HT", 0)),
                marge_pct=self._to_number(invoice.get("Marge_%", 0)),
                marge_amount=self._to_number(invoice.get("Marge_Montant", 0)),
                tva_pct=self._to_number(invoice.get("TVA_%", 0)),
                tva_amount=self._to_number(invoice.get("TVA_Montant", 0)),
                total_ttc=self._to_number(invoice.get("Total_TTC", 0)),
                acompte=self._to_number(invoice.get("Acompte", 0)),
                reste_a_payer=self._to_number(invoice.get("Reste_A_Payer", 0)),
                statut=str(invoice.get("Statut") or ""),
                items=line_items,
                base_taxable_ht=self._to_number(invoice.get("Base_Taxable_HT", 0)),
            )
        except Exception as e:
            messagebox.showerror("Erreur PDF", f"Impossible de générer la facture PDF.\n\n{e}")
            return

        messagebox.showinfo("Succès", f"Facture PDF générée:\n{filepath}")
        if os.path.exists(filepath):
            try:
                if os.name == "nt":
                    os.startfile(filepath)
                elif os.name == "posix":
                    subprocess.run(["xdg-open", filepath], check=False)
            except Exception:
                pass

    def _resolve_client_contact(self, invoice):
        client_id = str(invoice.get("Client_ID") or "").strip()
        client_name = str(invoice.get("Client_Nom") or "").strip().lower()
        client_email = ""
        client_phone = ""

        for client in load_all_clients():
            ref = str(client.get("ref_client") or "").strip()
            name = str(client.get("nom") or "").strip().lower()
            first = str(client.get("prenom") or "").strip().lower()
            full_name = f"{name} {first}".strip()

            id_match = client_id and ref and ref == client_id
            name_match = client_name and (client_name == name or client_name == full_name)
            if not id_match and not name_match:
                continue

            client_email = str(client.get("email") or "").strip()
            client_phone = str(client.get("telephone") or "").strip()
            break

        return client_email, client_phone

    def _build_invoice_pdf_items(self, invoice):
        """Build detailed PDF lines, especially hotels reserved."""
        source_type = str(invoice.get("Source_Type") or "")
        source_ref = str(invoice.get("Source_Ref") or "")
        client_id = str(invoice.get("Client_ID") or "").strip()
        client_name = str(invoice.get("Client_Nom") or "").strip()
        currency = str(invoice.get("Devise") or "Ariary")

        hotel_quotes = load_all_hotel_quotations()
        items = []

        if source_type in ("Devis client", "Client"):
            for q in hotel_quotes:
                q_client_id = str(q.get("client_id") or "").strip()
                q_client_name = str(q.get("client_name") or "").strip()
                if client_id and q_client_id != client_id:
                    continue
                if not client_id and client_name and q_client_name != client_name:
                    continue

                nights = int(self._to_number(q.get("nights", 0))) or 1
                total_price = self._to_number(q.get("total_price", 0))
                unit_price = total_price / nights if nights else total_price
                designation = (
                    f"{q.get('hotel_name', '')} - {q.get('city', '')} - {q.get('room_type', '')}"
                ).strip(" -")
                items.append(
                    {
                        "designation": designation or "Hôtel réservé",
                        "nights": nights,
                        "unit_price": unit_price,
                        "total": total_price,
                    }
                )

        elif source_type == "Hôtel" and "#" in source_ref:
            try:
                row_number = int(source_ref.split("#", 1)[1])
            except Exception:
                row_number = None
            if row_number is not None:
                quote = next(
                    (q for q in hotel_quotes if int(q.get("row_number", -1)) == row_number),
                    None,
                )
                if quote:
                    nights = int(self._to_number(quote.get("nights", 0))) or 1
                    total_price = self._to_number(quote.get("total_price", 0))
                    unit_price = total_price / nights if nights else total_price
                    designation = (
                        f"{quote.get('hotel_name', '')} - {quote.get('city', '')} - {quote.get('room_type', '')}"
                    ).strip(" -")
                    items.append(
                        {
                            "designation": designation or "Hôtel réservé",
                            "nights": nights,
                            "unit_price": unit_price,
                            "total": total_price,
                        }
                    )

        if not items:
            items = [
                {
                    "designation": f"{source_type} - {source_ref}",
                    "nights": 1,
                    "unit_price": self._to_number(invoice.get("Montant_HT", 0)),
                    "total": self._to_number(invoice.get("Montant_HT", 0)),
                }
            ]

        # Normalize totals if any line has empty/zero total by recomputing from unit*qty.
        for line in items:
            qty = int(self._to_number(line.get("nights", 1))) or 1
            unit = self._to_number(line.get("unit_price", 0))
            total = self._to_number(line.get("total", 0))
            if total <= 0 and unit > 0:
                line["total"] = unit * qty
            line["nights"] = qty
            line["unit_price"] = unit
            line["currency"] = currency
        return items

    def _validate_invoice_fields(self):
        try:
            montant = self._to_number(self.montant_ht_var.get())
            cout = self._to_number(self.cout_ht_var.get())
            marge = self._to_number(self.marge_pct_var.get())
            tva = self._to_number(self.tva_pct_var.get())
            acompte = self._to_number(self.acompte_var.get())
        except (TypeError, ValueError):
            messagebox.showerror("Erreur", "Veuillez saisir des nombres valides.")
            return False

        if montant <= 0:
            messagebox.showwarning("Montant HT", "Le montant HT doit être supérieur à 0.")
            return False
        if cout < 0 or marge < 0 or tva < 0 or acompte < 0:
            messagebox.showwarning("Valeurs invalides", "Les valeurs ne peuvent pas être négatives.")
            return False
        return True

    def _reset_form(self):
        self.montant_ht_var.set("0")
        self.cout_ht_var.set("0")
        self.marge_pct_var.set("0")
        self.tva_pct_var.set("20")
        self.acompte_pct_var.set("0")
        self.acompte_var.set("0")
        self.statut_var.set(INVOICE_STATUS_UNPAID)
        self._apply_acompte_lock()
        self._refresh_preview()

    def _refresh_all(self):
        refresh_financial_state_from_invoices()
        self._load_source_rows()
        self._load_invoices()
        self._render_invoices()
        self._render_financial_state()
