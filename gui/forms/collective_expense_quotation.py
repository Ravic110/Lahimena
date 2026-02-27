"""
Collective expense quotation GUI component
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
    get_collective_expense_headers,
    get_collective_expense_prestataires,
    get_collective_expense_designations,
    get_collective_expense_montant,
    get_collective_expense_forfait,
    load_all_clients,
    save_collective_expense_quotation_to_excel,
    update_collective_expense_quotation_in_excel,
)
from utils.logger import logger


class CollectiveExpenseQuotation:
    """
    Form for collective expense quotation with cascading dropdowns and auto-calculations.
    Saved in data.xlsx/COTATION_FRAIS_COL.
    """

    DEFAULT_HEADERS = [
        "Date",
        "ID_CLIENT",
        "Nom",
        "Numéro dossier",
        "Quantité",
        "Prestataire",
        "Désignation",
        "Montant",
        "Total",
        "Observation",
    ]

    def __init__(self, parent, edit_data=None, row_number=None, callback_on_save=None):
        self.parent = parent
        self.headers = []
        self.field_vars = {}
        self.field_widgets = {}
        self.clients = []
        self.client_map = {}
        self.dossier_numbers = []
        self.prestataires = []
        self.designations_map = {}
        self.edit_data = edit_data  # Data to edit, None for new
        self.row_number = row_number  # Row number for editing
        self.callback_on_save = callback_on_save  # Callback after successful save

        self._load_base_data()
        self._load_headers()
        self._create_form()
        if self.edit_data:
            self._load_edit_data()

    def _load_base_data(self):
        """Load clients and collective expense data"""
        self.clients = load_all_clients()
        self.client_map = {}
        for client in self.clients:
            ref_client = str(client.get("ref_client") or "").strip()
            if ref_client:
                self.client_map[ref_client] = client

        self.dossier_numbers = sorted(
            {
                str(client.get("numero_dossier") or "").strip()
                for client in self.clients
                if str(client.get("numero_dossier") or "").strip()
            },
            key=lambda value: value.lower(),
        )

        self.prestataires = get_collective_expense_prestataires()

    def _load_headers(self):
        headers = get_collective_expense_headers()
        self.headers = headers if headers else self.DEFAULT_HEADERS.copy()

        required_fields = [
            ("date", "Date"),
            ("client_id", "ID_CLIENT"),
            ("client_name", "Nom"),
        ]

        for field_type, fallback_label in reversed(required_fields):
            has_field = any(
                self._get_field_type(header) == field_type for header in self.headers
            )
            if not has_field:
                self.headers.insert(0, fallback_label)

        client_name_index = next(
            (
                index
                for index, header in enumerate(self.headers)
                if self._get_field_type(header) == "client_name"
            ),
            None,
        )

        dossier_index = next(
            (
                index
                for index, header in enumerate(self.headers)
                if self._get_field_type(header) == "dossier_number"
            ),
            None,
        )

        if client_name_index is not None:
            target_index = client_name_index + 1
            if dossier_index is None:
                self.headers.insert(target_index, "Numéro dossier")
            elif dossier_index != target_index:
                dossier_header = self.headers.pop(dossier_index)
                if dossier_index < target_index:
                    target_index -= 1
                self.headers.insert(target_index, dossier_header)

    def _normalize_header(self, header):
        normalized = (
            str(header)
            .strip()
            .lower()
            .replace("_", " ")
            .replace("-", " ")
            .replace("°", "o")
            .replace("º", "o")
            .replace(".", " ")
            .replace("/", " ")
        )
        for char, replacement in [
            ("é", "e"), ("è", "e"), ("ê", "e"), ("à", "a"),
            ("â", "a"), ("ï", "i"), ("î", "i"),
        ]:
            normalized = normalized.replace(char, replacement)
        normalized = " ".join(normalized.split())
        return normalized

    def _find_header_by_type(self, field_type):
        """Find the actual header name for a given field type"""
        for header, (ftype, _) in self.field_widgets.items():
            if ftype == field_type:
                return header
        return None

    def _get_field_type(self, header):
        """Determine field type based on header name"""
        norm = self._normalize_header(header)
        
        if norm in {"nom", "nom client", "client nom", "nom du client"}:
            return "client_name"
        elif norm in {"id", "id client", "id_client", "ref client", "reference", "ref"}:
            return "client_id"
        elif norm in {
            "numero dossier",
            "numero de dossier",
            "no dossier",
            "no de dossier",
            "num dossier",
            "num de dossier",
            "n dossier",
            "no dossier",
            "dossier",
            "numero du dossier",
            "num du dossier",
            "n o dossier",
        }:
            return "dossier_number"
        elif "dossier" in norm:
            return "dossier_number"
        elif norm in {"quantite", "quantité", "qty", "quantity", "nombre"}:
            return "quantity"
        elif norm in {"prestataire", "fournisseur", "provider"}:
            return "prestataire"
        elif norm in {"designation", "désignation", "libelle", "description", "service"}:
            return "designation"
        elif norm in {"forfait", "forfai"}:
            return "forfait"
        elif norm in {"montant", "price", "prix", "fee", "montant unitaire"}:
            return "montant"
        elif norm in {"total", "montant total", "total prix"}:
            return "total"
        elif norm == "date":
            return "date"
        return "text"

    def _load_edit_data(self):
        """Load and populate form with existing data for editing"""
        if not self.edit_data:
            return
        
        try:
            for header, var in self.field_vars.items():
                value = self.edit_data.get(header, "")
                var.set(str(value).strip())
            
            # Manually trigger cascading updates
            prestataire_header = self._find_header_by_type("prestataire")
            if prestataire_header:
                self._on_prestataire_changed()
            
            designat_header = self._find_header_by_type("designation")
            if designat_header:
                self._on_designation_changed()
            
            logger.info("Loaded edit data into form")
        except Exception as e:
            logger.error(f"Error loading edit data: {e}", exc_info=True)

    def _create_form(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        title_text = "MODIFIER COTATION FRAIS COLLECTIFS" if self.edit_data else "COTATION FRAIS COLLECTIFS"
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
            text="1) Sélectionnez le client  2) Choisissez prestataire/désignation  3) Vérifiez quantité/total puis Enregistrer.",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        info.pack(pady=(0, 10))

        form_frame = tk.LabelFrame(
            self.parent,
            text="Informations cotation",
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
                client_ids = [""] + sorted(
                    self.client_map.keys(), 
                    key=lambda x: (x is None, x)
                )
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=client_ids,
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_client_id_changed)
            elif field_type == "client_name":
                client_names = [""] + sorted(
                    set(c.get("nom", "") for c in self.clients if c.get("nom", "").strip()),
                    key=lambda x: (x == "", x)
                )
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=client_names,
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_client_name_changed)
            elif field_type == "dossier_number":
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[""] + self.dossier_numbers,
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_dossier_number_changed)
            elif field_type == "prestataire":
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[""] + self.prestataires,
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
            elif field_type == "forfait":
                widget = tk.Entry(
                    form_frame,
                    textvariable=field_var,
                    font=ENTRY_FONT,
                    width=32,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                    state="readonly",
                )
            elif field_type == "montant":
                widget = tk.Entry(
                    form_frame,
                    textvariable=field_var,
                    font=ENTRY_FONT,
                    width=32,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                    state="readonly",
                )
            elif field_type == "total":
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

        button_text = "💾 Modifications" if self.edit_data else "💾 Enregistrer"
        tk.Button(
            button_frame,
            text=button_text,
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

    def _on_client_id_changed(self, event=None):
        """Auto-fill when client ID selected"""
        try:
            client_id_header = self._find_header_by_type("client_id")
            if not client_id_header:
                logger.warning("client_id header not found")
                return
            
            client_id = self.field_vars.get(client_id_header, tk.StringVar()).get().strip()
            logger.info(f"Client ID selected: {client_id}")
            
            if client_id and client_id in self.client_map:
                client = self.client_map[client_id]
                nom = str(client.get("nom") or "").strip()
                
                # Set Nom
                nom_header = self._find_header_by_type("client_name")
                if nom_header:
                    self.field_vars[nom_header].set(nom)
                    logger.info(f"Set {nom_header} to {nom}")

                # Set Numéro dossier
                dossier_header = self._find_header_by_type("dossier_number")
                if dossier_header:
                    numero_dossier = str(client.get("numero_dossier") or "").strip()
                    self.field_vars[dossier_header].set(numero_dossier)
                    logger.info(f"Set {dossier_header} to {numero_dossier}")
                
                # Set Quantité
                quantite_str = str(client.get("nombre_participants") or "").strip()
                if quantite_str:
                    qty_header = self._find_header_by_type("quantity")
                    if qty_header:
                        try:
                            qty_value = int(float(quantite_str))
                            self.field_vars[qty_header].set(str(qty_value))
                            logger.info(f"Set {qty_header} to {qty_value}")
                        except ValueError:
                            pass
        except Exception as e:
            logger.error(f"Error in _on_client_id_changed: {e}", exc_info=True)

    def _on_client_name_changed(self, event=None):
        """Auto-fill when client name selected"""
        try:
            nom_header = self._find_header_by_type("client_name")
            if not nom_header:
                logger.warning("client_name header not found")
                return
            
            nom = self.field_vars.get(nom_header, tk.StringVar()).get().strip()
            logger.info(f"Client name selected: {nom}")
            
            for ref, client in self.client_map.items():
                if str(client.get("nom") or "").strip() == nom:
                    # Set ID_CLIENT
                    client_id_header = self._find_header_by_type("client_id")
                    if client_id_header:
                        self.field_vars[client_id_header].set(ref)
                        logger.info(f"Set {client_id_header} to {ref}")

                    # Set Numéro dossier
                    dossier_header = self._find_header_by_type("dossier_number")
                    if dossier_header:
                        numero_dossier = str(client.get("numero_dossier") or "").strip()
                        self.field_vars[dossier_header].set(numero_dossier)
                        logger.info(f"Set {dossier_header} to {numero_dossier}")
                    
                    # Set Quantité
                    quantite_str = str(client.get("nombre_participants") or "").strip()
                    if quantite_str:
                        qty_header = self._find_header_by_type("quantity")
                        if qty_header:
                            try:
                                qty_value = int(float(quantite_str))
                                self.field_vars[qty_header].set(str(qty_value))
                                logger.info(f"Set {qty_header} to {qty_value}")
                            except ValueError:
                                pass
                    break
        except Exception as e:
            logger.error(f"Error in _on_client_name_changed: {e}", exc_info=True)

    def _on_dossier_number_changed(self, event=None):
        """Auto-fill client fields when dossier number selected"""
        try:
            dossier_header = self._find_header_by_type("dossier_number")
            if not dossier_header:
                logger.warning("dossier_number header not found")
                return

            numero_dossier = self.field_vars.get(dossier_header, tk.StringVar()).get().strip()
            logger.info(f"Dossier selected: {numero_dossier}")

            if not numero_dossier:
                return

            for ref, client in self.client_map.items():
                client_dossier = str(client.get("numero_dossier") or "").strip()
                if client_dossier == numero_dossier:
                    client_id_header = self._find_header_by_type("client_id")
                    if client_id_header:
                        self.field_vars[client_id_header].set(ref)
                        logger.info(f"Set {client_id_header} to {ref}")

                    self._on_client_id_changed()
                    break
        except Exception as e:
            logger.error(f"Error in _on_dossier_number_changed: {e}", exc_info=True)

    def _on_prestataire_changed(self, event=None):
        """Update designation dropdown when prestataire changes"""
        try:
            prestataire_header = self._find_header_by_type("prestataire")
            if not prestataire_header:
                return
            
            prestataire = self.field_vars.get(prestataire_header, tk.StringVar()).get().strip()
            designations = (
                get_collective_expense_designations(prestataire)
                if prestataire else []
            )
            
            designation_header = self._find_header_by_type("designation")
            if designation_header and designation_header in self.field_widgets:
                _, widget = self.field_widgets[designation_header]
                widget["values"] = [""] + designations
                self.field_vars[designation_header].set("")
                logger.info(f"Updated designations for {prestataire}: {designations}")
                
            self._update_montant_and_total()
        except Exception as e:
            logger.error(f"Error in _on_prestataire_changed: {e}", exc_info=True)

    def _on_designation_changed(self, event=None):
        """Update montant when designation changes"""
        self._update_montant_and_total()

    def _on_quantity_changed(self, event=None):
        """Recalculate total when quantity changes"""
        self._update_montant_and_total()

    def _update_montant_and_total(self):
        """Calculate and update forfait + montant + total"""
        try:
            prestataire_header = self._find_header_by_type("prestataire")
            designation_header = self._find_header_by_type("designation")
            quantity_header = self._find_header_by_type("quantity")
            montant_header = self._find_header_by_type("montant")
            total_header = self._find_header_by_type("total")
            
            # Find forfait header by looking for it in field_widgets
            forfait_header = None
            for header in self.field_vars.keys():
                if self._normalize_header(header) == "forfait":
                    forfait_header = header
                    break
            
            prestataire = (
                self.field_vars.get(prestataire_header, tk.StringVar()).get().strip()
                if prestataire_header else ""
            )
            designation = (
                self.field_vars.get(designation_header, tk.StringVar()).get().strip()
                if designation_header else ""
            )
            quantite_str = (
                self.field_vars.get(quantity_header, tk.StringVar()).get().strip()
                if quantity_header else ""
            )

            # Update forfait
            if prestataire and designation:
                forfait = get_collective_expense_forfait(prestataire, designation)
                if forfait_header:
                    self.field_vars[forfait_header].set(forfait)
                    logger.info(f"Set forfait to {forfait}")

            # Update montant
            montant = 0
            if prestataire and designation:
                montant = get_collective_expense_montant(prestataire, designation)
                if montant_header:
                    self.field_vars[montant_header].set(str(montant))
                    logger.info(f"Set montant to {montant}")

            # Update total
            total = 0
            if montant > 0 and quantite_str:
                try:
                    quantite = float(quantite_str)
                    total = montant * quantite
                except ValueError:
                    pass

            if total_header:
                self.field_vars[total_header].set(str(total) if total > 0 else "")
                logger.info(f"Set total to {total}")
        except Exception as e:
            logger.error(f"Error in _update_montant_and_total: {e}", exc_info=True)

    def _collect_form_data(self):
        data = {}
        for header, var in self.field_vars.items():
            data[header] = var.get().strip()
        return data

    def _save(self):
        form_data = self._collect_form_data()

        non_empty_values = [
            value
            for key, value in form_data.items()
            if key.strip().lower() != "date" and str(value).strip()
        ]
        if not non_empty_values:
            messagebox.showwarning(
                "Validation",
                "Veuillez renseigner au moins un champ avant l'enregistrement.",
            )
            return

        client_id_header = self._find_header_by_type("client_id")
        client_name_header = self._find_header_by_type("client_name")
        client_id = form_data.get(client_id_header, "").strip() if client_id_header else ""
        client_name = form_data.get(client_name_header, "").strip() if client_name_header else ""
        if not client_id and not client_name:
            messagebox.showwarning(
                "Validation",
                "Veuillez sélectionner un client avant d'enregistrer la cotation.",
            )
            return

        if self.edit_data and self.row_number is not None:
            # Update existing row
            result = update_collective_expense_quotation_in_excel(self.row_number, form_data)
            if result == -2:
                messagebox.showerror(
                    "Fichier verrouillé",
                    "Impossible de modifier car data.xlsx est ouvert/verrouillé.\n\n"
                    "Fermez le fichier dans Excel puis réessayez.",
                )
                return
            if result == -1:
                messagebox.showerror(
                    "Erreur", "Échec de la modification dans COTATION_FRAIS_COL."
                )
                return
            messagebox.showinfo(
                "Succès",
                f"Cotation modifiée avec succès.",
            )
            # Call callback if provided, otherwise cancel
            if self.callback_on_save:
                self.callback_on_save()
            else:
                self._cancel()
        else:
            # Add new row
            row = save_collective_expense_quotation_to_excel(form_data)
            if row == -2:
                messagebox.showerror(
                    "Fichier verrouillé",
                    "Impossible d'enregistrer car data.xlsx est ouvert/verrouillé.\n\n"
                    "Fermez le fichier dans Excel puis réessayez.",
                )
                return
            if row == -1:
                messagebox.showerror(
                    "Erreur", "Échec de l'enregistrement dans COTATION_FRAIS_COL."
                )
                return

            messagebox.showinfo(
                "Succès",
                f"Cotation enregistrée avec succès à la ligne {row}.",
            )
            if self.callback_on_save:
                self.callback_on_save()
            else:
                self._clear()

    def _delete(self):
        """Delete the current row being edited"""
        if not self.edit_data or self.row_number is None:
            messagebox.showwarning("Info", "Aucune cotation à supprimer.")
            return
        
        confirm = messagebox.askyesno(
            "Confirmation",
            "Êtes-vous sûr de vouloir supprimer cette cotation ?\n\nCette action ne peut pas être annulée."
        )
        if not confirm:
            return
        
        from utils.excel_handler import delete_collective_expense_from_excel
        
        success = delete_collective_expense_from_excel(self.row_number)
        if not success:
            messagebox.showerror(
                "Erreur",
                "Impossible de supprimer. Vérifiez que data.xlsx n'est pas ouvert."
            )
            return
        
        messagebox.showinfo(
            "Succès",
            "Cotation supprimée avec succès."
        )
        self._cancel()
    
    def _cancel(self):
        """Cancel edit and navigate back"""
        logger.info("Edit form cancelled")
    
    def _clear(self):
        for header, var in self.field_vars.items():
            if header.strip().lower() == "date":
                var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            else:
                var.set("")

    def _reload_from_sheet(self):
        self._load_base_data()
        self._load_headers()
        self._create_form()
