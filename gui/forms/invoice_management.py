"""Invoice management GUI component."""

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


class InvoiceManagement:
    """Create and track invoices with VAT, margins, deposits and statuses."""

    SOURCE_TYPES = [
        "Hôtel",
        "Frais collectifs",
        "Visite & Excursion",
        "Billet avion",
        "Transport",
    ]

    def __init__(self, parent):
        self.parent = parent
        self.source_rows = []
        self.invoices = []
        self.source_map = {}
        self.invoice_tree = None

        self.source_type_var = tk.StringVar(value=self.SOURCE_TYPES[0])
        self.source_ref_var = tk.StringVar()
        self.montant_ht_var = tk.StringVar(value="0")
        self.cout_ht_var = tk.StringVar(value="0")
        self.marge_pct_var = tk.StringVar(value="0")
        self.tva_pct_var = tk.StringVar(value="20")
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
        self._field(form, 6, "Acompte", self._entry(form, self.acompte_var))
        self._field(form, 7, "Statut", self._status_combo(form))

        actions = tk.Frame(form, bg=MAIN_BG_COLOR)
        actions.grid(row=8, column=0, columnspan=2, sticky="w", pady=(10, 0))

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
        return ttk.Combobox(
            parent,
            textvariable=self.statut_var,
            values=[INVOICE_STATUS_UNPAID, INVOICE_STATUS_PARTIAL, INVOICE_STATUS_PAID],
            font=ENTRY_FONT,
            width=24,
            state="readonly",
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
        return f"{source_type}#{row_number} - {client} - {total:,.2f}"

    def _load_source_rows(self):
        source_type = self.source_type_var.get()
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

    def _update_selected_invoice(self):
        selection = self.invoice_tree.selection()
        if not selection:
            messagebox.showwarning("Facture", "Sélectionnez une facture à mettre à jour.")
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

    def _refresh_all(self):
        refresh_financial_state_from_invoices()
        self._load_source_rows()
        self._load_invoices()
        self._render_invoices()
        self._render_financial_state()
