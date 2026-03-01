"""
Client quotation GUI component
"""

import datetime
import os
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_RED,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
    DEVIS_FOLDER,
)
from utils.excel_handler import (
    load_all_clients,
    load_all_collective_expense_quotations,
    load_all_hotel_quotations,
    load_all_transport_quotations,
    save_invoice_to_excel,
)
from utils.logger import logger
from utils.pdf_generator import REPORTLAB_AVAILABLE, generate_client_quotation_pdf


class ClientQuotation:
    """
    Client quotation component for generating a consolidated client quote
    """

    def __init__(self, parent):
        self.parent = parent
        self.quotations = []
        self.clients = []
        self.client_display_map = {}
        self.client_info_map = {}
        self.current_currency = "Ariary"
        self.current_items = []
        self.client_quotes = []
        self.all_client_lines = []

        self._load_data()
        self._create_interface()

    def _load_data(self):
        """Load quotations and client info"""
        try:
            self.quotations = load_all_hotel_quotations()
            self.collective_quotes = load_all_collective_expense_quotations()
            self.transport_quotes = load_all_transport_quotations()
        except Exception as e:
            logger.error(f"Failed to load hotel quotations: {e}", exc_info=True)
            self.quotations = []
            self.collective_quotes = []
            self.transport_quotes = []

        self.clients = []
        self.client_display_map = {}
        self.client_info_map = {}

        try:
            for client in load_all_clients():
                ref = (client.get("ref_client") or "").strip()
                name = f"{(client.get('nom') or '').strip()} {(client.get('prenom') or '').strip()}".strip()
                if ref:
                    self.client_info_map[ref] = {
                        "name": name or ref,
                        "email": (client.get("email") or "").strip(),
                        "phone": (client.get("telephone") or "").strip(),
                    }
        except Exception as e:
            logger.warning(f"Failed to load client info: {e}")

        for ref, info in sorted(self.client_info_map.items(), key=lambda kv: kv[0].lower()):
            display = f"{ref} - {info.get('name', ref)}"
            self.clients.append(display)
            self.client_display_map[display] = ref

    def _create_interface(self):
        """Create the client quotation interface"""
        for widget in self.parent.winfo_children():
            widget.destroy()

        title = tk.Label(
            self.parent,
            text="DEVIS CLIENT",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        title.pack(pady=(20, 10))

        tk.Label(
            self.parent,
            text="1) Sélectionnez le client  2) Ajustez marge/TVA  3) Générez le devis ou la facture",
            font=ENTRY_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 8))

        main_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        style = ttk.Style()
        style.configure(
            "Treeview",
            background=INPUT_BG_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR,
        )
        style.map("Treeview", background=[("selected", BUTTON_GREEN)])

        # Client selection
        client_frame = tk.LabelFrame(
            main_frame,
            text="Sélection client",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        client_frame.pack(fill="x", pady=(0, 10))

        tk.Label(
            client_frame,
            text="Client:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=0, column=0, sticky="w", pady=5)

        self.client_var = tk.StringVar()
        self.client_combo = ttk.Combobox(
            client_frame,
            textvariable=self.client_var,
            values=[""] + self.clients,
            font=ENTRY_FONT,
            width=40,
            state="readonly",
        )
        self.client_combo.grid(row=0, column=1, padx=(10, 20), pady=5, sticky="w")
        self.client_combo.bind("<<ComboboxSelected>>", self._on_client_selected)

        # Client contact info
        tk.Label(
            client_frame,
            text="Email:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=1, column=0, sticky="w", pady=5)

        self.email_var = tk.StringVar()
        self.email_entry = tk.Entry(
            client_frame,
            textvariable=self.email_var,
            font=ENTRY_FONT,
            width=30,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        )
        self.email_entry.grid(row=1, column=1, padx=(10, 20), pady=5, sticky="w")

        tk.Label(
            client_frame,
            text="Téléphone:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=1, column=2, sticky="w", pady=5)

        self.phone_var = tk.StringVar()
        self.phone_entry = tk.Entry(
            client_frame,
            textvariable=self.phone_var,
            font=ENTRY_FONT,
            width=20,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        )
        self.phone_entry.grid(row=1, column=3, padx=(10, 0), pady=5, sticky="w")

        # Pricing parameters
        pricing_frame = tk.LabelFrame(
            main_frame,
            text="Paramètres de prix",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        pricing_frame.pack(fill="x", pady=(0, 10))

        tk.Label(
            pricing_frame,
            text="Devise:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=0, column=0, sticky="w", pady=5)

        self.currency_var = tk.StringVar(value="Toutes")
        self.currency_combo = ttk.Combobox(
            pricing_frame,
            textvariable=self.currency_var,
            values=["Toutes"],
            font=ENTRY_FONT,
            width=12,
            state="readonly",
        )
        self.currency_combo.grid(row=0, column=1, padx=(10, 20), pady=5, sticky="w")
        self.currency_combo.bind("<<ComboboxSelected>>", self._on_currency_changed)

        tk.Label(
            pricing_frame,
            text="Marge (%):",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=0, column=2, sticky="w", pady=5)

        self.margin_var = tk.StringVar(value="0")
        self.margin_entry = tk.Entry(
            pricing_frame,
            textvariable=self.margin_var,
            font=ENTRY_FONT,
            width=10,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        )
        self.margin_entry.grid(row=0, column=3, padx=(10, 20), pady=5, sticky="w")
        self.margin_entry.bind("<KeyRelease>", self._on_pricing_changed)

        tk.Label(
            pricing_frame,
            text="TVA (%):",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=0, column=4, sticky="w", pady=5)

        self.tva_var = tk.StringVar(value="20")
        self.tva_entry = tk.Entry(
            pricing_frame,
            textvariable=self.tva_var,
            font=ENTRY_FONT,
            width=10,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        )
        self.tva_entry.grid(row=0, column=5, padx=(10, 0), pady=5, sticky="w")
        self.tva_entry.bind("<KeyRelease>", self._on_pricing_changed)

        # Quotations table
        table_frame = tk.LabelFrame(
            main_frame,
            text="Lignes de devis",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        table_frame.pack(fill="both", expand=True, pady=(0, 10))

        columns = ("designation", "nights", "unit_price", "total")
        self.items_tree = ttk.Treeview(
            table_frame, columns=columns, height=10, show="headings"
        )
        self.items_tree.heading("designation", text="Désignation")
        self.items_tree.heading("nights", text="Nuits")
        self.items_tree.heading("unit_price", text="Prix Unitaire")
        self.items_tree.heading("total", text="Total")

        self.items_tree.column("designation", width=320)
        self.items_tree.column("nights", width=60)
        self.items_tree.column("unit_price", width=120)
        self.items_tree.column("total", width=120)

        self.items_tree.pack(fill="both", expand=True)
        self.items_tree.bind("<Double-1>", self._edit_selected_line)

        line_action_frame = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
        line_action_frame.pack(fill="x", pady=(0, 10))
        tk.Button(
            line_action_frame,
            text="✏️ Modifier ligne",
            command=self._edit_selected_line,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=10,
            pady=6,
        ).pack(side="left", padx=(0, 6))
        tk.Button(
            line_action_frame,
            text="🗑️ Supprimer ligne",
            command=self._remove_selected_line,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            padx=10,
            pady=6,
        ).pack(side="left")

        # Totals section
        totals_frame = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
        totals_frame.pack(fill="x", pady=(0, 10))

        self.subtotal_label = tk.Label(
            totals_frame,
            text="Sous-total: 0",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        self.subtotal_label.pack(side="left", padx=(0, 20))

        self.margin_label = tk.Label(
            totals_frame,
            text="Marge: 0",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        self.margin_label.pack(side="left", padx=(0, 20))

        self.tva_label = tk.Label(
            totals_frame,
            text="TVA: 0",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        self.tva_label.pack(side="left", padx=(0, 20))

        self.total_label = tk.Label(
            totals_frame,
            text="TOTAL: 0",
            font=("Arial", 11, "bold"),
            fg=BUTTON_GREEN,
            bg=MAIN_BG_COLOR,
        )
        self.total_label.pack(side="right")

        # Action buttons
        action_frame = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
        action_frame.pack(fill="x", pady=(0, 10))

        tk.Button(
            action_frame,
            text="📄 Générer devis client",
            command=self._generate_client_quote,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8,
        ).pack(side="left", padx=5)

        tk.Button(
            action_frame,
            text="💰 Générer la facture",
            command=self._generate_invoice_from_quote,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8,
        ).pack(side="left", padx=5)

        tk.Button(
            action_frame,
            text="🔄 Rafraîchir",
            command=self._refresh_data,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8,
        ).pack(side="left", padx=5)

        tk.Button(
            action_frame,
            text="🧹 Réinitialiser",
            command=self._reset_form,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8,
        ).pack(side="left", padx=5)

        if not self.clients:
            messagebox.showinfo(
                "Aucune cotation",
                "Aucune cotation hôtel n'est disponible. Créez d'abord une cotation hôtel.",
            )

    def _reset_form(self):
        self.client_var.set("")
        self.email_var.set("")
        self.phone_var.set("")
        self.margin_var.set("0")
        self.tva_var.set("20")
        self.current_currency = "Ariary"
        self.current_items = []
        self.client_quotes = []
        self.all_client_lines = []
        self.currency_var.set("Toutes")
        self.currency_combo["values"] = ["Toutes"]
        self._clear_table()
        self._update_totals(0, 0, 0, 0)

    def _refresh_data(self):
        self._load_data()
        self.client_combo["values"] = [""] + self.clients
        self._reset_form()

    def _clear_table(self):
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)

    def _quote_key(self, quotation):
        client_id = (quotation.get("client_id") or "").strip()
        client_name = (quotation.get("client_name") or "").strip()
        return client_id or client_name

    def _on_client_selected(self, event=None):
        display = self.client_var.get()
        key = self.client_display_map.get(display, "")
        if not key:
            self._reset_form()
            return

        info = self.client_info_map.get(key, {})
        self.email_var.set(info.get("email", ""))
        self.phone_var.set(info.get("phone", ""))

        selected_name = str(info.get("name") or "").strip()
        self.all_client_lines = self._build_client_lines(key, selected_name)
        self.client_quotes = list(self.all_client_lines)

        if not self.all_client_lines:
            self._clear_table()
            self.current_items = []
            self._update_totals(0, 0, 0, 0)
            return

        currencies = sorted({line.get("currency") or "Ariary" for line in self.all_client_lines})
        self.currency_combo["values"] = ["Toutes"] + currencies
        if len(currencies) == 1:
            self.currency_var.set(currencies[0])
        else:
            self.currency_var.set("Toutes")

        self._apply_currency_filter()

    def _apply_currency_filter(self):
        filter_value = self.currency_var.get()
        if filter_value == "Toutes":
            filtered = list(self.all_client_lines)
        else:
            filtered = [
                q for q in self.all_client_lines if (q.get("currency") or "Ariary") == filter_value
            ]

        if not filtered:
            self._clear_table()
            self.current_items = []
            self.current_currency = "Ariary"
            self._update_totals(0, 0, 0, 0)
            return

        currencies = {q.get("currency") or "Ariary" for q in filtered}
        if len(currencies) > 1:
            messagebox.showwarning(
                "Devis multi-devises",
                "Sélectionnez une devise pour générer un devis.",
            )
            self._clear_table()
            self.current_items = []
            self._update_totals(0, 0, 0, 0)
            return

        self.current_currency = currencies.pop()

        self.current_items = []
        for line in filtered:
            qty = max(1, int(self._to_number(line.get("nights", line.get("quantity", 1)))))
            total = float(self._to_number(line.get("total", line.get("total_price", 0))))
            unit = float(line.get("unit_price", 0)) or (total / qty if qty else 0)
            designation = str(line.get("designation") or "").strip()
            self.current_items.append(
                {
                    "designation": designation,
                    "nights": qty,
                    "unit_price": unit,
                    "total": unit * qty,
                }
            )

        self._render_items_table()
        self._recalculate_totals()

    def _normalize_key(self, value):
        if value is None:
            return ""
        text = str(value).strip().lower()
        for src, dst in (
            ("é", "e"), ("è", "e"), ("ê", "e"), ("ë", "e"),
            ("à", "a"), ("â", "a"), ("ä", "a"),
            ("î", "i"), ("ï", "i"),
            ("ô", "o"), ("ö", "o"),
            ("ù", "u"), ("û", "u"), ("ü", "u"),
            ("ç", "c"), ("_", " "), ("-", " "),
        ):
            text = text.replace(src, dst)
        return " ".join(text.split())

    def _to_number(self, value):
        try:
            if isinstance(value, (int, float)):
                return float(value)
            text = str(value or "").replace(" ", "").replace(",", ".")
            cleaned = ""
            used_dot = False
            for ch in text:
                if ch.isdigit() or ch == "-":
                    cleaned += ch
                elif ch == "." and not used_dot:
                    cleaned += ch
                    used_dot = True
            return float(cleaned) if cleaned not in ("", "-", ".") else 0.0
        except Exception:
            return 0.0

    def _find_value(self, row, aliases, contains=False):
        normalized_aliases = [self._normalize_key(alias) for alias in aliases]
        for key, value in row.items():
            nk = self._normalize_key(key)
            for alias in normalized_aliases:
                if nk == alias or (contains and alias in nk):
                    if value not in (None, ""):
                        return value
        return ""

    def _build_client_lines(self, client_ref, client_name):
        lines = []
        client_ref_norm = self._normalize_key(client_ref)
        client_name_norm = self._normalize_key(client_name)

        for quotation in self.quotations:
            key = self._quote_key(quotation)
            if self._normalize_key(key) != client_ref_norm:
                continue
            nights = max(1, int(self._to_number(quotation.get("nights", 1))))
            total_price = self._to_number(quotation.get("total_price", 0))
            designation = (
                f"Hôtel - {quotation.get('hotel_name', '')}"
                f" - {quotation.get('city', '')}"
                f" - {quotation.get('room_type', '')}"
            ).strip(" -")
            lines.append(
                {
                    "designation": designation,
                    "nights": nights,
                    "unit_price": total_price / nights if nights else total_price,
                    "total": total_price,
                    "currency": quotation.get("currency") or "Ariary",
                }
            )

        for row in self.collective_quotes:
            row_id = self._normalize_key(self._find_value(row, ["ID_CLIENT", "id client", "ref client", "reference"], contains=True))
            row_name = self._normalize_key(self._find_value(row, ["Nom", "nom client", "client"], contains=True))
            if client_ref_norm and row_id and row_id != client_ref_norm:
                continue
            if client_ref_norm and not row_id and client_name_norm and row_name and row_name not in client_name_norm:
                continue
            qty = max(1, int(self._to_number(self._find_value(row, ["Quantité", "quantite", "qte", "nombre"], contains=True) or 1)))
            total = self._to_number(self._find_value(row, ["Total", "montant total"], contains=True))
            unit = self._to_number(self._find_value(row, ["Montant", "prix unitaire"], contains=True))
            if total <= 0 and unit > 0:
                total = unit * qty
            if unit <= 0 and qty > 0:
                unit = total / qty
            if total <= 0:
                continue
            presta = self._find_value(row, ["Prestataire"], contains=True)
            design = self._find_value(row, ["Désignation", "designation", "description"], contains=True)
            designation = f"Frais collectifs - {presta} - {design}".strip(" -")
            lines.append(
                {
                    "designation": designation,
                    "nights": qty,
                    "unit_price": unit,
                    "total": total,
                    "currency": self._find_value(row, ["Devise", "currency"], contains=True) or "Ariary",
                }
            )

        for row in self.transport_quotes:
            row_id = self._normalize_key(self._find_value(row, ["ID_CLIENT", "id client", "ref client", "reference"], contains=True))
            row_name = self._normalize_key(self._find_value(row, ["Nom", "nom client", "client"], contains=True))
            if client_ref_norm and row_id and row_id != client_ref_norm:
                continue
            if client_ref_norm and not row_id and client_name_norm and row_name and row_name not in client_name_norm:
                continue

            total = 0.0
            for key, value in row.items():
                nk = self._normalize_key(key)
                if "budget" in nk or ("total" in nk and "km" not in nk) or "montant" in nk:
                    total = max(total, self._to_number(value))
            if total <= 0:
                continue
            depart = self._find_value(row, ["ville depart", "depart"], contains=True)
            arrivee = self._find_value(row, ["ville arrivee", "arrivee"], contains=True)
            vehicule = self._find_value(row, ["type voiture", "voiture", "transport"], contains=True)
            designation = f"Transport - {depart} > {arrivee} - {vehicule}".strip(" -")
            lines.append(
                {
                    "designation": designation,
                    "nights": 1,
                    "unit_price": total,
                    "total": total,
                    "currency": self._find_value(row, ["Devise", "currency"], contains=True) or "Ariary",
                }
            )

        return lines

    def _render_items_table(self):
        self._clear_table()
        for idx, item in enumerate(self.current_items):
            self.items_tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    item.get("designation", ""),
                    int(item.get("nights", 0)),
                    f"{float(item.get('unit_price', 0)):,.2f} {self.current_currency}",
                    f"{float(item.get('total', 0)):,.2f} {self.current_currency}",
                ),
            )

    def _edit_selected_line(self, event=None):
        selection = self.items_tree.selection()
        if not selection:
            messagebox.showwarning("Ligne", "Veuillez sélectionner une ligne à modifier.")
            return
        idx = int(selection[0])
        if idx < 0 or idx >= len(self.current_items):
            return
        item = self.current_items[idx]

        win = tk.Toplevel(self.parent)
        win.title("Modifier ligne")
        win.configure(bg=MAIN_BG_COLOR)
        win.resizable(False, False)
        win.transient(self.parent)
        win.grab_set()

        tk.Label(win, text="Désignation:", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=0, column=0, sticky="w", padx=10, pady=8)
        designation_var = tk.StringVar(value=str(item.get("designation", "")))
        tk.Entry(win, textvariable=designation_var, font=ENTRY_FONT, width=40, bg=INPUT_BG_COLOR, fg=TEXT_COLOR).grid(row=0, column=1, padx=10, pady=8)

        tk.Label(win, text="Nuits/Quantité:", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=1, column=0, sticky="w", padx=10, pady=8)
        nights_var = tk.StringVar(value=str(int(item.get("nights", 1))))
        tk.Entry(win, textvariable=nights_var, font=ENTRY_FONT, width=20, bg=INPUT_BG_COLOR, fg=TEXT_COLOR).grid(row=1, column=1, sticky="w", padx=10, pady=8)

        tk.Label(win, text="Prix unitaire:", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=2, column=0, sticky="w", padx=10, pady=8)
        unit_var = tk.StringVar(value=str(float(item.get("unit_price", 0))))
        tk.Entry(win, textvariable=unit_var, font=ENTRY_FONT, width=20, bg=INPUT_BG_COLOR, fg=TEXT_COLOR).grid(row=2, column=1, sticky="w", padx=10, pady=8)

        def apply_changes():
            qty = max(1, int(self._to_number(nights_var.get())))
            unit = max(0.0, self._to_number(unit_var.get()))
            self.current_items[idx] = {
                "designation": designation_var.get().strip(),
                "nights": qty,
                "unit_price": unit,
                "total": qty * unit,
            }
            self._render_items_table()
            self._recalculate_totals()
            win.destroy()

        btns = tk.Frame(win, bg=MAIN_BG_COLOR)
        btns.grid(row=3, column=0, columnspan=2, pady=(8, 10))
        tk.Button(btns, text="✅ Appliquer", command=apply_changes, bg=BUTTON_GREEN, fg="white", font=BUTTON_FONT, padx=10, pady=4).pack(side="left", padx=5)
        tk.Button(btns, text="❌ Annuler", command=win.destroy, bg=BUTTON_RED, fg="white", font=BUTTON_FONT, padx=10, pady=4).pack(side="left", padx=5)

    def _remove_selected_line(self):
        selection = self.items_tree.selection()
        if not selection:
            messagebox.showwarning("Ligne", "Veuillez sélectionner une ligne à supprimer.")
            return
        idx = int(selection[0])
        if idx < 0 or idx >= len(self.current_items):
            return
        self.current_items.pop(idx)
        self._render_items_table()
        self._recalculate_totals()

    def _parse_percent(self, value, label):
        if not value:
            return 0.0
        try:
            return float(value.replace(",", "."))
        except Exception:
            raise ValueError(f"{label} invalide")

    def _recalculate_totals(self, subtotal=None):
        if subtotal is None:
            subtotal = sum(item["total"] for item in self.current_items)

        try:
            margin_pct = self._parse_percent(self.margin_var.get().strip(), "Marge")
            tva_pct = self._parse_percent(self.tva_var.get().strip(), "TVA")
        except ValueError as e:
            messagebox.showerror("Erreur", str(e))
            return

        margin_amount = subtotal * (margin_pct / 100)
        subtotal_with_margin = subtotal + margin_amount
        tva_amount = subtotal_with_margin * (tva_pct / 100)
        total = subtotal_with_margin + tva_amount

        self._update_totals(subtotal, margin_amount, tva_amount, total)

    def _update_totals(self, subtotal, margin_amount, tva_amount, total):
        currency = self.current_currency
        self.subtotal_label.config(
            text=f"Sous-total: {subtotal:,.2f} {currency}"
        )
        self.margin_label.config(text=f"Marge: {margin_amount:,.2f} {currency}")
        self.tva_label.config(text=f"TVA: {tva_amount:,.2f} {currency}")
        self.total_label.config(text=f"TOTAL: {total:,.2f} {currency}")

    def _on_pricing_changed(self, event=None):
        if not self.current_items:
            return
        self._recalculate_totals()

    def _on_currency_changed(self, event=None):
        if not self.client_quotes:
            return
        self._apply_currency_filter()

    def _generate_client_quote(self):
        if not self.current_items:
            messagebox.showwarning("Aucune ligne", "Aucune ligne de devis à générer.")
            return

        if not REPORTLAB_AVAILABLE:
            messagebox.showwarning(
                "⚠️ Génération PDF non disponible",
                "ReportLab n'est pas installé. Veuillez installer le package:\n\npip install reportlab",
            )
            return

        try:
            margin_pct = self._parse_percent(self.margin_var.get().strip(), "Marge")
            tva_pct = self._parse_percent(self.tva_var.get().strip(), "TVA")
        except ValueError as e:
            messagebox.showerror("Erreur", str(e))
            return

        client_display = self.client_var.get().strip()
        key = self.client_display_map.get(client_display, "")
        client_id = key if (key and key in self.client_info_map) else ""
        client_name = client_display.replace(f"{client_id} - ", "").strip() or key
        client_email = self.email_var.get().strip()
        client_phone = self.phone_var.get().strip()

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        quote_number = f"DEVIS_CLIENT_{client_id}_{timestamp}" if client_id else f"DEVIS_CLIENT_{timestamp}"

        subtotal = sum(item["total"] for item in self.current_items)
        margin_amount = subtotal * (margin_pct / 100)
        subtotal_with_margin = subtotal + margin_amount
        tva_amount = subtotal_with_margin * (tva_pct / 100)
        total = subtotal_with_margin + tva_amount

        filename = generate_client_quotation_pdf(
            client_name=client_name,
            client_email=client_email,
            client_phone=client_phone,
            quote_number=quote_number,
            quote_date=datetime.datetime.now().strftime("%d/%m/%Y"),
            items=self.current_items,
            currency=self.current_currency,
            margin_pct=margin_pct,
            margin_amount=margin_amount,
            tva_pct=tva_pct,
            tva_amount=tva_amount,
            subtotal=subtotal,
            total=total,
            output_dir=DEVIS_FOLDER,
        )

        messagebox.showinfo(
            "✅ Devis généré",
            f"Le devis PDF a été sauvegardé :\n{filename}",
        )

        try:
            if os.name == "nt":
                os.startfile(filename)
            elif os.name == "posix":
                subprocess.run(["xdg-open", filename], check=False)
        except Exception as e:
            logger.warning(f"Could not open quotation file automatically: {e}")

    def _generate_invoice_from_quote(self):
        if not self.current_items:
            messagebox.showwarning("Aucune ligne", "Aucune ligne de devis à facturer.")
            return

        try:
            margin_pct = self._parse_percent(self.margin_var.get().strip(), "Marge")
            tva_pct = self._parse_percent(self.tva_var.get().strip(), "TVA")
        except ValueError as e:
            messagebox.showerror("Erreur", str(e))
            return

        client_display = self.client_var.get().strip()
        key = self.client_display_map.get(client_display, "")
        client_id = key if (key and key in self.client_info_map) else ""
        client_name = client_display.replace(f"{client_id} - ", "").strip() or key

        if not client_name:
            messagebox.showwarning("Client", "Veuillez sélectionner un client valide.")
            return

        subtotal = sum(item["total"] for item in self.current_items)

        invoice_data = {
            "Source_Type": "Devis client",
            "Source_Ref": f"DEVIS_CLIENT:{client_id or client_name}",
            "Client_ID": client_id,
            "Client_Nom": client_name,
            "Devise": self.current_currency or "Ariary",
            "Montant_HT": subtotal,
            "Marge_%": margin_pct,
            "TVA_%": tva_pct,
            "Acompte": 0,
            "Statut": "non payée",
        }

        row = save_invoice_to_excel(invoice_data)
        if row == -2:
            messagebox.showerror(
                "Erreur",
                "Impossible d'écrire dans Excel. Fermez data.xlsx puis réessayez.",
            )
            return
        if row < 0:
            messagebox.showerror("Erreur", "La génération de la facture a échoué.")
            return

        messagebox.showinfo(
            "Facture générée",
            "La facture a été créée et intégrée automatiquement dans l'état financier.",
        )
