"""
Visite & Excursion quotation GUI component
"""

from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

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
)
from utils.excel_handler import (
    delete_visite_excursion_from_excel,
    get_visite_excursion_designations,
    get_visite_excursion_headers,
    get_visite_excursion_montant,
    get_visite_excursion_prestataires,
    load_all_clients,
    save_visite_excursion_quotation_to_excel,
    update_visite_excursion_quotation_in_excel,
)
from utils.logger import logger


class VisiteExcursionQuotation:
    """
    Form for visite & excursion quotation with cascading dropdowns and auto-calculations.
    Saved in data.xlsx/VISITE_EXCURSION.
    """

    DEFAULT_HEADERS = [
        "Date",
        "ID_CLIENT",
        "Nom",
        "Prénom",
        "Quantité",
        "Prestation",
        "Désignation",
        "Montant",
        "Total",
        "Observation",
    ]

    def __init__(self, parent, edit_data=None, row_number=None, callback_on_done=None):
        self.parent = parent
        self.headers = []
        self.field_vars = {}
        self.field_widgets = {}
        self.clients = []
        self.client_map = {}
        self.prestations = []
        self.edit_data = edit_data
        self.row_number = row_number
        self.callback_on_done = callback_on_done

        self._load_base_data()
        self._load_headers()
        self._create_form()
        if self.edit_data:
            self._load_edit_data()

    def _load_base_data(self):
        self.clients = load_all_clients()
        self.client_map = {}
        for client in self.clients:
            ref_client = str(client.get("ref_client") or "").strip()
            if ref_client:
                self.client_map[ref_client] = client

        self.prestations = get_visite_excursion_prestataires()

    def _load_headers(self):
        headers = get_visite_excursion_headers()
        self.headers = headers if headers else self.DEFAULT_HEADERS.copy()
        if "Date" not in self.headers:
            self.headers.insert(0, "Date")

    def _normalize_header(self, header):
        normalized = str(header).strip().lower().replace("_", " ").replace("-", " ")
        for char, replacement in [
            ("é", "e"),
            ("è", "e"),
            ("ê", "e"),
            ("à", "a"),
            ("â", "a"),
            ("ï", "i"),
            ("î", "i"),
            ("ô", "o"),
        ]:
            normalized = normalized.replace(char, replacement)
        return normalized

    def _find_header_by_type(self, field_type):
        for header, (ftype, _) in self.field_widgets.items():
            if ftype == field_type:
                return header
        return None

    def _get_field_type(self, header):
        norm = self._normalize_header(header)

        if norm in {"nom", "nom client", "client nom", "nom du client"}:
            return "client_name"
        if norm in {"prenom", "prénom", "prenom client", "client prenom"}:
            return "client_firstname"
        if norm in {"id", "id client", "id_client", "ref client", "reference", "ref"}:
            return "client_id"
        if norm in {"quantite", "quantité", "qty", "quantity", "nombre", "participants", "nombre participants"}:
            return "quantity"
        if norm in {"prestation", "prestations", "prestataire", "prestataires", "fournisseur", "provider"}:
            return "prestataire"
        if norm in {"designation", "désignation", "libelle", "description", "service"}:
            return "designation"
        if norm in {"montant", "price", "prix", "fee", "montant unitaire", "tarif", "tarif par pax"}:
            return "montant"
        if norm in {"total", "montant total", "total prix"}:
            return "total"
        if norm == "date":
            return "date"
        return "text"

    def _create_form(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        title_text = "MODIFIER VISITE & EXCURSION" if self.edit_data else "VISITE & EXCURSION"
        title = tk.Label(
            self.parent,
            text=title_text,
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        title.pack(pady=(20, 10), fill="x")

        info = tk.Label(
            self.parent,
            text="Sélectionnez ID client/Nom/Prénom pour auto-remplir. Prestation + Désignation remplissent Montant, puis Total = Montant × Quantité.",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        info.pack(pady=(0, 10))

        form_frame = tk.LabelFrame(
            self.parent,
            text="Informations visite & excursion",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        form_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.field_vars = {}
        self.field_widgets = {}

        for index, header in enumerate(self.headers):
            row = index // 2
            col_group = index % 2

            label_col = col_group * 2
            entry_col = label_col + 1

            tk.Label(
                form_frame,
                text=f"{header} :",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=row, column=label_col, sticky="w", padx=(0, 8), pady=6)

            field_var = tk.StringVar()
            field_type = self._get_field_type(header)

            if field_type == "date":
                field_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                widget = tk.Entry(
                    form_frame,
                    textvariable=field_var,
                    font=ENTRY_FONT,
                    width=32,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                    state="readonly",
                )
            elif field_type == "client_id":
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[""] + sorted(self.client_map.keys()),
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_client_id_changed)
            elif field_type == "client_name":
                client_names = sorted({str(c.get("nom") or "").strip() for c in self.clients if str(c.get("nom") or "").strip()})
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[""] + client_names,
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_client_name_or_firstname_changed)
            elif field_type == "client_firstname":
                first_names = sorted({str(c.get("prenom") or "").strip() for c in self.clients if str(c.get("prenom") or "").strip()})
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[""] + first_names,
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_client_name_or_firstname_changed)
            elif field_type == "prestataire":
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[""] + self.prestations,
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_prestataire_changed)
            elif field_type == "designation":
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[],
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_designation_changed)
            elif field_type in {"montant", "total"}:
                widget = tk.Entry(
                    form_frame,
                    textvariable=field_var,
                    font=ENTRY_FONT,
                    width=32,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                    state="readonly",
                )
            elif field_type == "quantity":
                widget = tk.Entry(
                    form_frame,
                    textvariable=field_var,
                    font=ENTRY_FONT,
                    width=32,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                )
                widget.bind("<KeyRelease>", self._on_quantity_changed)
            else:
                widget = tk.Entry(
                    form_frame,
                    textvariable=field_var,
                    font=ENTRY_FONT,
                    width=32,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                )

            widget.grid(row=row, column=entry_col, sticky="we", padx=(0, 16), pady=6)
            self.field_vars[header] = field_var
            self.field_widgets[header] = (field_type, widget)

        form_frame.grid_columnconfigure(1, weight=1)
        form_frame.grid_columnconfigure(3, weight=1)

        button_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        button_frame.pack(fill="x", padx=20, pady=(0, 20))

        tk.Button(
            button_frame,
            text="💾 Modifications" if self.edit_data else "💾 Enregistrer",
            command=self._save,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=16,
            pady=6,
        ).pack(side="left", padx=(0, 8))

        if self.edit_data:
            tk.Button(
                button_frame,
                text="🗑️ Supprimer",
                command=self._delete,
                bg=BUTTON_RED,
                fg="white",
                font=BUTTON_FONT,
                padx=16,
                pady=6,
            ).pack(side="left", padx=(0, 8))

            tk.Button(
                button_frame,
                text="❌ Annuler",
                command=self._cancel,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=16,
                pady=6,
            ).pack(side="left", padx=(0, 8))
        else:
            tk.Button(
                button_frame,
                text="🧹 Réinitialiser",
                command=self._clear,
                bg=BUTTON_RED,
                fg="white",
                font=BUTTON_FONT,
                padx=16,
                pady=6,
            ).pack(side="left", padx=(0, 8))

            tk.Button(
                button_frame,
                text="🔄 Recharger",
                command=self._reload_from_sheet,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=16,
                pady=6,
            ).pack(side="left")

    def _load_edit_data(self):
        try:
            for header, var in self.field_vars.items():
                var.set(str(self.edit_data.get(header, "")).strip())

            prestataire_header = self._find_header_by_type("prestataire")
            designation_header = self._find_header_by_type("designation")
            if prestataire_header and designation_header:
                prestataire = self.field_vars[prestataire_header].get().strip()
                if prestataire and designation_header in self.field_widgets:
                    _, widget = self.field_widgets[designation_header]
                    widget["values"] = [""] + get_visite_excursion_designations(prestataire, self._build_source_filters())

            self._update_montant_and_total()
        except Exception as e:
            logger.error(f"Error loading edit data: {e}", exc_info=True)

    def _find_matching_client(self):
        id_header = self._find_header_by_type("client_id")
        nom_header = self._find_header_by_type("client_name")
        prenom_header = self._find_header_by_type("client_firstname")

        ref = self.field_vars.get(id_header, tk.StringVar()).get().strip() if id_header else ""
        nom = self.field_vars.get(nom_header, tk.StringVar()).get().strip() if nom_header else ""
        prenom = self.field_vars.get(prenom_header, tk.StringVar()).get().strip() if prenom_header else ""

        if ref and ref in self.client_map:
            return self.client_map[ref]

        matches = self.clients
        if nom:
            matches = [c for c in matches if str(c.get("nom") or "").strip() == nom]
        if prenom:
            matches = [c for c in matches if str(c.get("prenom") or "").strip() == prenom]

        return matches[0] if matches else None

    def _apply_client_to_fields(self, client):
        if not client:
            return

        id_header = self._find_header_by_type("client_id")
        nom_header = self._find_header_by_type("client_name")
        prenom_header = self._find_header_by_type("client_firstname")
        qty_header = self._find_header_by_type("quantity")

        if id_header:
            self.field_vars[id_header].set(str(client.get("ref_client") or "").strip())
        if nom_header:
            self.field_vars[nom_header].set(str(client.get("nom") or "").strip())
        if prenom_header:
            self.field_vars[prenom_header].set(str(client.get("prenom") or "").strip())
        if qty_header:
            q = str(client.get("nombre_participants") or "").strip()
            if q:
                try:
                    self.field_vars[qty_header].set(str(int(float(q))))
                except Exception:
                    self.field_vars[qty_header].set(q)

    def _on_client_id_changed(self, event=None):
        try:
            client = self._find_matching_client()
            self._apply_client_to_fields(client)
            self._update_montant_and_total()
        except Exception as e:
            logger.error(f"Error on client id change: {e}", exc_info=True)

    def _on_client_name_or_firstname_changed(self, event=None):
        try:
            client = self._find_matching_client()
            self._apply_client_to_fields(client)
            self._update_montant_and_total()
        except Exception as e:
            logger.error(f"Error on client name/prenom change: {e}", exc_info=True)

    def _on_prestataire_changed(self, event=None):
        try:
            prestataire_header = self._find_header_by_type("prestataire")
            designation_header = self._find_header_by_type("designation")
            if not prestataire_header or not designation_header:
                return

            prestataire = self.field_vars[prestataire_header].get().strip()
            filters = self._build_source_filters()
            designations = get_visite_excursion_designations(prestataire, filters) if prestataire else get_visite_excursion_designations(filters=filters)

            _, widget = self.field_widgets[designation_header]
            widget["values"] = [""] + designations
            self.field_vars[designation_header].set("")

            self._update_montant_and_total()
        except Exception as e:
            logger.error(f"Error on prestataire change: {e}", exc_info=True)

    def _on_designation_changed(self, event=None):
        self._refresh_prestation_options()
        self._update_montant_and_total()

    def _on_quantity_changed(self, event=None):
        self._update_montant_and_total()

    def _build_source_filters(self):
        filters = {}
        for header, var in self.field_vars.items():
            value = var.get().strip()
            if not value:
                continue
            field_type = self._get_field_type(header)
            if field_type in {"date", "client_id", "client_name", "client_firstname", "quantity", "montant", "total"}:
                continue
            filters[header] = value
        return filters

    def _refresh_prestation_options(self):
        prestataire_header = self._find_header_by_type("prestataire")
        if not prestataire_header or prestataire_header not in self.field_widgets:
            return

        current_value = self.field_vars[prestataire_header].get().strip()
        filters = self._build_source_filters()
        options = get_visite_excursion_prestataires(filters)

        _, widget = self.field_widgets[prestataire_header]
        widget["values"] = [""] + options

        if current_value and current_value not in options:
            self.field_vars[prestataire_header].set("")

    def _update_montant_and_total(self):
        try:
            prestataire_header = self._find_header_by_type("prestataire")
            designation_header = self._find_header_by_type("designation")
            quantity_header = self._find_header_by_type("quantity")
            montant_header = self._find_header_by_type("montant")
            total_header = self._find_header_by_type("total")

            prestataire = self.field_vars.get(prestataire_header, tk.StringVar()).get().strip() if prestataire_header else ""
            designation = self.field_vars.get(designation_header, tk.StringVar()).get().strip() if designation_header else ""
            quantity_str = self.field_vars.get(quantity_header, tk.StringVar()).get().strip() if quantity_header else ""

            montant = 0
            if prestataire and designation:
                montant = get_visite_excursion_montant(prestataire, designation, self._build_source_filters())

            if montant_header:
                self.field_vars[montant_header].set(str(montant) if montant else "")

            total = 0
            if montant and quantity_str:
                try:
                    total = float(montant) * float(quantity_str)
                except Exception:
                    total = 0

            if total_header:
                self.field_vars[total_header].set(str(total) if total else "")
        except Exception as e:
            logger.error(f"Error updating montant/total: {e}", exc_info=True)

    def _collect_form_data(self):
        return {header: var.get().strip() for header, var in self.field_vars.items()}

    def _save(self):
        form_data = self._collect_form_data()

        has_data = any(str(v).strip() for k, v in form_data.items() if self._normalize_header(k) != "date")
        if not has_data:
            messagebox.showwarning("Validation", "Veuillez renseigner au moins un champ avant l'enregistrement.")
            return

        if self.edit_data and self.row_number is not None:
            result = update_visite_excursion_quotation_in_excel(self.row_number, form_data)
            if result == -2:
                messagebox.showerror("Fichier verrouillé", "Fermez data.xlsx puis réessayez.")
                return
            if result == -1:
                messagebox.showerror("Erreur", "Échec de la modification dans VISITE_EXCURSION.")
                return
            messagebox.showinfo("Succès", "Cotation modifiée avec succès.")
            if self.callback_on_done:
                self.callback_on_done()
            return

        row = save_visite_excursion_quotation_to_excel(form_data)
        if row == -2:
            messagebox.showerror("Fichier verrouillé", "Fermez data.xlsx puis réessayez.")
            return
        if row == -1:
            messagebox.showerror("Erreur", "Échec de l'enregistrement dans VISITE_EXCURSION.")
            return

        messagebox.showinfo("Succès", f"Cotation enregistrée avec succès à la ligne {row}.")
        self._clear()

    def _delete(self):
        if self.row_number is None:
            return
        confirm = messagebox.askyesno(
            "Confirmation",
            "Êtes-vous sûr de vouloir supprimer cette cotation ?\n\nCette action ne peut pas être annulée.",
        )
        if not confirm:
            return

        success = delete_visite_excursion_from_excel(self.row_number)
        if not success:
            messagebox.showerror("Erreur", "Impossible de supprimer. Vérifiez que data.xlsx n'est pas ouvert.")
            return

        messagebox.showinfo("Succès", "Cotation supprimée avec succès.")
        if self.callback_on_done:
            self.callback_on_done()

    def _cancel(self):
        if self.callback_on_done:
            self.callback_on_done()

    def _clear(self):
        for header, var in self.field_vars.items():
            if self._normalize_header(header) == "date":
                var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            else:
                var.set("")

    def _reload_from_sheet(self):
        self._load_base_data()
        self._load_headers()
        self._create_form()
