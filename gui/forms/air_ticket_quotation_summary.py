"""
Air ticket quotations summary GUI component
"""

import tkinter as tk
from tkinter import messagebox, ttk

from config import (
    ACCENT_TEXT_COLOR,
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_RED,
    CARD_BG_COLOR,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.excel_handler import (
    delete_air_ticket_from_excel,
    load_all_air_ticket_quotations,
)
from utils.logger import logger


class AirTicketQuotationSummary:
    """Display list/summary for air ticket quotations."""

    def __init__(self, parent, callback_edit=None, callback_add=None):
        self.parent = parent
        self.rows = []
        self.search_var = tk.StringVar()
        self.view_var = tk.StringVar(value="Toutes les lignes")
        self.tree = None
        self.callback_edit = callback_edit
        self.callback_add = callback_add

        self._load_rows()
        self._create_interface()

    def _load_rows(self):
        try:
            self.rows = load_all_air_ticket_quotations()
        except Exception as e:
            logger.error(f"Error loading air ticket quotations: {e}", exc_info=True)
            self.rows = []

    def _normalize(self, value):
        return str(value or "").strip().lower()

    def _display_header(self, header):
        norm = self._normalize(header)
        norm = (
            norm.replace("é", "e")
            .replace("è", "e")
            .replace("ê", "e")
            .replace("'", " ")
            .replace("’", " ")
        )
        if "tarif" in norm and "bebe" in norm:
            return "Tarif Enfant"
        return header

    def _ordered_headers(self, headers):
        priority_keywords = [
            "date",
            "id_client",
            "id client",
            "ref client",
            "nom",
            "ville de depart",
            "ville depart",
            "ville d arrive",
            "ville d arrivee",
            "nombre adultes",
            "nombre adulte",
            "nombres adultes",
            "nombre enfants",
            "nombre enfant",
            "nombres enfants",
            "tarif adulte",
            "tarif adultes",
            "tarif bebe",
            "tarif enfant",
            "tarifs enfants",
            "montant adultes",
            "montant adulte",
            "montant enfants",
            "montant enfant",
            "total",
            "observation",
        ]

        normalized_map = {h: self._normalize(h).replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("â", "a").replace("ô", "o").replace("ï", "i").replace("î", "i").replace("'", " ").replace("’", " ") for h in headers}

        ordered = []
        used = set()
        for key in priority_keywords:
            for header in headers:
                if header in used:
                    continue
                if key == normalized_map.get(header):
                    ordered.append(header)
                    used.add(header)
                    break

        for header in headers:
            if header not in used:
                ordered.append(header)

        return ordered

    def _guess_amount_headers(self):
        if not self.rows:
            return []
        headers = [key for key in self.rows[0].keys() if key != "row_number"]
        keywords = ["montant", "total", "prix", "tarif", "amount", "price", "cost"]
        return [h for h in headers if any(k in self._normalize(h) for k in keywords)]

    def _to_number(self, value):
        try:
            if value is None or value == "":
                return 0.0
            if isinstance(value, (int, float)):
                return float(value)
            text = str(value).replace(" ", "").replace(",", ".")
            cleaned = ""
            has_dot = False
            for char in text:
                if char.isdigit() or char == "-":
                    cleaned += char
                elif char == "." and not has_dot:
                    cleaned += char
                    has_dot = True
            return float(cleaned) if cleaned not in ("", "-", ".") else 0.0
        except Exception:
            return 0.0

    def _get_reference_header(self):
        if not self.rows:
            return None
        headers = [key for key in self.rows[0].keys() if key != "row_number"]
        preferred = ["Référence", "Reference", "ID", "Ref", "ID_CLIENT", "Nom"]
        for pref in preferred:
            for header in headers:
                if self._normalize(pref) == self._normalize(header):
                    return header
        return headers[0] if headers else None

    def _create_interface(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        title = tk.Label(
            self.parent,
            text="RÉSUMÉ RÉSERVATIONS BILLET AVION",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        title.pack(pady=(20, 10), fill="x")

        self.main_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        top_controls = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        top_controls.pack(fill="x", padx=8, pady=(0, 10))

        tk.Label(top_controls, text="Vue:", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).pack(side="left")
        view_combo = ttk.Combobox(
            top_controls,
            textvariable=self.view_var,
            values=["Toutes les lignes", "Par référence"],
            font=ENTRY_FONT,
            width=20,
            state="readonly",
        )
        view_combo.pack(side="left", padx=(8, 16))
        view_combo.bind("<<ComboboxSelected>>", self._refresh_display)

        tk.Label(top_controls, text="Recherche:", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).pack(side="left")
        search_entry = tk.Entry(
            top_controls,
            textvariable=self.search_var,
            font=ENTRY_FONT,
            width=28,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        )
        search_entry.pack(side="left", padx=(8, 16))
        self.search_var.trace("w", self._refresh_display)

        tk.Button(
            top_controls,
            text="🔄 Rafraîchir",
            command=self._on_refresh_clicked,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left")

        action_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        action_frame.pack(fill="x", padx=8, pady=(0, 10))

        tk.Button(
            action_frame,
            text="➕ Ajouter",
            command=self._on_add_clicked,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            action_frame,
            text="✏️ Modifier",
            command=self._on_edit_clicked,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            action_frame,
            text="🗑️ Supprimer",
            command=self._on_delete_clicked,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left")

        self.kpi_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        self.kpi_frame.pack(fill="x", padx=8, pady=(0, 10))

        self.content_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        self.content_frame.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure(
            "Treeview",
            background=INPUT_BG_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR,
        )
        style.map("Treeview", background=[("selected", BUTTON_GREEN)])

        self._refresh_display()

    def _on_refresh_clicked(self):
        self._load_rows()
        self._refresh_display()

    def _on_add_clicked(self):
        if self.callback_add:
            self.callback_add()

    def _on_edit_clicked(self):
        if self.tree is None:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à modifier.")
            return

        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à modifier.")
            return

        selected_item = selection[0]
        try:
            row_number = int(selected_item)
        except (TypeError, ValueError):
            return

        selected_row = next((row for row in self.rows if row.get("row_number") == row_number), None)
        if selected_row is None:
            return

        if self.callback_edit:
            self.callback_edit(selected_row, row_number)

    def _on_delete_clicked(self):
        if self.tree is None:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à supprimer.")
            return

        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à supprimer.")
            return

        confirm = messagebox.askyesno(
            "Confirmation",
            "Êtes-vous sûr de vouloir supprimer cette réservation ?\n\nCette action ne peut pas être annulée.",
        )
        if not confirm:
            return

        selected_item = selection[0]
        try:
            row_number = int(selected_item)
        except (TypeError, ValueError):
            return

        success = delete_air_ticket_from_excel(row_number)
        if success:
            messagebox.showinfo("Succès", "Réservation supprimée avec succès.")
            self._on_refresh_clicked()
        else:
            messagebox.showerror("Erreur", "Impossible de supprimer. Vérifiez que data.xlsx n'est pas ouvert.")

    def _matches_query(self, row, query):
        if not query:
            return True
        for key, value in row.items():
            if key == "row_number":
                continue
            if query in self._normalize(value):
                return True
        return False

    def _filtered_rows(self):
        query = self._normalize(self.search_var.get())
        return [row for row in self.rows if self._matches_query(row, query)]

    def _render_kpis(self, filtered_rows):
        for widget in self.kpi_frame.winfo_children():
            widget.destroy()

        total_rows = len(filtered_rows)
        amount_headers = self._guess_amount_headers()
        total_amount = 0.0
        for row in filtered_rows:
            for header in amount_headers:
                total_amount += self._to_number(row.get(header, 0))

        card = tk.Frame(self.kpi_frame, bg=CARD_BG_COLOR, bd=2, relief="ridge")
        card.pack(fill="x")

        tk.Label(
            card,
            text=f"Nombre de lignes: {total_rows}",
            font=("Arial", 11, "bold"),
            fg=TEXT_COLOR,
            bg=CARD_BG_COLOR,
        ).pack(side="left", padx=12, pady=8)

        tk.Label(
            card,
            text=f"Total montants: {total_amount:,.2f}",
            font=("Arial", 11, "bold"),
            fg=ACCENT_TEXT_COLOR,
            bg=CARD_BG_COLOR,
        ).pack(side="right", padx=12, pady=8)

    def _clear_content(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def _refresh_display(self, *args):
        filtered_rows = self._filtered_rows()
        self._render_kpis(filtered_rows)

        if self.view_var.get() == "Par référence":
            self._render_grouped_view(filtered_rows)
        else:
            self._render_lines_view(filtered_rows)

    def _render_lines_view(self, rows):
        self._clear_content()

        if not rows:
            tk.Label(
                self.content_frame,
                text="Aucune réservation trouvée",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            self.tree = None
            return

        headers = self._ordered_headers([key for key in rows[0].keys() if key != "row_number"])
        self.tree = ttk.Treeview(self.content_frame, columns=headers, show="headings", selectmode="browse")

        for header in headers:
            self.tree.heading(header, text=self._display_header(header))
            self.tree.column(header, width=150, anchor="w")

        scrollbar_y = ttk.Scrollbar(self.content_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(self.content_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        for row in rows:
            values = [row.get(header, "") for header in headers]
            iid = str(row.get("row_number") or "")
            if not iid:
                iid = None
            self.tree.insert("", "end", iid=iid, values=values)

        self.tree.pack(fill="both", expand=True, side="left")
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

    def _render_grouped_view(self, rows):
        self._clear_content()

        if not rows:
            tk.Label(
                self.content_frame,
                text="Aucune réservation trouvée",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            self.tree = None
            return

        ref_header = self._get_reference_header()
        if not ref_header:
            self._render_lines_view(rows)
            return

        grouped = {}
        for row in rows:
            key = row.get(ref_header, "Sans référence")
            grouped.setdefault(key, []).append(row)

        canvas = tk.Canvas(self.content_frame, bg=MAIN_BG_COLOR, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.content_frame, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=MAIN_BG_COLOR)

        inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )

        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        for ref, group_rows in grouped.items():
            frame = tk.LabelFrame(
                inner,
                text=f"{ref_header}: {ref} ({len(group_rows)} ligne(s))",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
                padx=8,
                pady=6,
            )
            frame.pack(fill="x", padx=8, pady=6)

            headers = self._ordered_headers([key for key in group_rows[0].keys() if key != "row_number"])
            for row in group_rows:
                line = " | ".join(f"{h}: {row.get(h, '')}" for h in headers[:5])
                tk.Label(
                    frame,
                    text=line,
                    font=("Arial", 9),
                    fg=TEXT_COLOR,
                    bg=MAIN_BG_COLOR,
                    anchor="w",
                    justify="left",
                ).pack(fill="x", pady=1)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree = None
