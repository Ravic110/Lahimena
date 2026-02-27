"""
Air ticket quotation GUI component
"""

from datetime import datetime
import re
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
    delete_air_ticket_from_excel,
    get_avion_arrival_cities,
    get_avion_departure_cities,
    get_avion_headers,
    get_avion_tarifs,
    load_all_clients,
    save_air_ticket_quotation_to_excel,
    update_air_ticket_quotation_in_excel,
)
from utils.logger import logger


class AirTicketQuotation:
    """Form for air ticket quotation with client auto-fill and montant calculations."""

    DEFAULT_HEADERS = [
        "Date",
        "ID_CLIENT",
        "Nom",
        "Ville de départ",
        "Ville d'arrivée",
        "Nombre Adultes",
        "Nombre Enfants",
        "Tarif Adultes",
        "Tarifs Enfants",
        "Montant Adultes",
        "Montant Enfants",
        "Total",
        "Observation",
    ]

    def __init__(
        self,
        parent,
        edit_data=None,
        row_number=None,
        callback_on_done=None,
        callback_on_save=None,
    ):
        self.parent = parent
        self.headers = []
        self.field_vars = {}
        self.field_widgets = {}
        self.clients = []
        self.client_map = {}
        self.clients_by_name = {}
        self.edit_data = edit_data
        self.row_number = row_number
        self.callback_on_done = callback_on_done or callback_on_save

        self._load_base_data()
        self._load_headers()
        self._create_form()
        if self.edit_data:
            self._load_edit_data()

    def _load_base_data(self):
        self.clients = load_all_clients()
        self.client_map = {}
        self.clients_by_name = {}
        for client in self.clients:
            ref_client = str(client.get("ref_client") or "").strip()
            nom = str(client.get("nom") or "").strip()
            if ref_client:
                self.client_map[ref_client] = client
            if nom and nom not in self.clients_by_name:
                self.clients_by_name[nom] = client

    def _load_headers(self):
        headers = get_avion_headers()
        self.headers = headers if headers else self.DEFAULT_HEADERS.copy()

        required = [
            ("date", "Date"),
            ("client_id", "ID_CLIENT"),
            ("client_name", "Nom"),
            ("adult_count", "Nombre Adultes"),
            ("child_count", "Nombre Enfants"),
        ]
        for field_type, fallback_header in reversed(required):
            has_field = any(self._get_field_type(header) == field_type for header in self.headers)
            if not has_field:
                self.headers.insert(0, fallback_header)

        self._reorder_priority_headers()

    def _reorder_priority_headers(self):
        priority_types = [
            "date",
            "client_id",
            "client_name",
            "departure_city",
            "arrival_city",
            "adult_count",
            "child_count",
        ]

        type_to_header = {}
        for header in self.headers:
            field_type = self._get_field_type(header)
            if field_type in priority_types and field_type not in type_to_header:
                type_to_header[field_type] = header

        ordered_priority_headers = [
            type_to_header[field_type]
            for field_type in priority_types
            if field_type in type_to_header
        ]

        remaining_headers = [
            header for header in self.headers if header not in ordered_priority_headers
        ]

        self.headers = ordered_priority_headers + remaining_headers

    def _build_form_layout_headers(self):
        left_types = [
            "date",
            "client_id",
            "client_name",
            "departure_city",
            "arrival_city",
            "child_count",
        ]
        right_types = [
            "adult_count",
            "child_tarif",
            "adult_tarif",
            "child_amount",
            "adult_amount",
            "total",
        ]

        type_to_header = {}
        for header in self.headers:
            field_type = self._get_field_type(header)
            if field_type not in type_to_header:
                type_to_header[field_type] = header

        left_headers = [type_to_header[t] for t in left_types if t in type_to_header]
        right_headers = [type_to_header[t] for t in right_types if t in type_to_header]

        ordered = []
        used = set()
        row_count = max(len(left_headers), len(right_headers))
        for idx in range(row_count):
            if idx < len(left_headers):
                ordered.append(left_headers[idx])
                used.add(left_headers[idx])
            if idx < len(right_headers):
                ordered.append(right_headers[idx])
                used.add(right_headers[idx])

        remaining_headers = [header for header in self.headers if header not in used]
        ordered.extend(remaining_headers)
        return ordered

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
            ("ù", "u"),
            ("û", "u"),
            ("ü", "u"),
            ("ç", "c"),
            ("'", " "),
            ("’", " "),
        ]:
            normalized = normalized.replace(char, replacement)
        normalized = re.sub(r"[^a-z0-9\s]", " ", normalized)
        normalized = re.sub(r"\s+", " ", normalized).strip()
        return normalized

    def _display_label(self, header):
        norm = self._normalize_header(header)
        if "tarif" in norm and "bebe" in norm:
            return "Tarif Enfant"
        return str(header)

    def _format_city_label(self, value):
        text = str(value or "").strip()
        if not text:
            return ""
        parts = [part for part in text.split() if part]
        return " ".join(part[:1].upper() + part[1:].lower() for part in parts)

    def _set_city_options(self, widget, cities):
        display_to_raw = {self._format_city_label(city): city for city in cities if str(city).strip()}
        widget._display_to_raw = display_to_raw
        widget._raw_to_display = {
            raw: display for display, raw in display_to_raw.items()
        }
        options = [""] + sorted(display_to_raw.keys())
        widget["values"] = options

    def _resolve_city_raw(self, widget, value):
        text = str(value or "").strip()
        if not text:
            return ""

        display_to_raw = getattr(widget, "_display_to_raw", {})
        raw_to_display = getattr(widget, "_raw_to_display", {})

        if text in display_to_raw:
            return display_to_raw[text]
        if text in raw_to_display:
            return text

        normalized_display = self._format_city_label(text)
        if normalized_display in display_to_raw:
            return display_to_raw[normalized_display]

        lowered = text.lower()
        for raw in raw_to_display.keys():
            if str(raw).strip().lower() == lowered:
                return raw
        for display, raw in display_to_raw.items():
            if str(display).strip().lower() == lowered:
                return raw

        return text

    def _normalize_city_selection(self, header):
        if not header or header not in self.field_widgets:
            return
        _, widget = self.field_widgets[header]
        current = self.field_vars[header].get().strip()
        display_to_raw = getattr(widget, "_display_to_raw", {})
        raw_to_display = getattr(widget, "_raw_to_display", {})

        if current in display_to_raw:
            raw = display_to_raw[current]
            display = raw_to_display.get(raw, current)
            self.field_vars[header].set(display)
            return

        display = raw_to_display.get(current)
        if display:
            self.field_vars[header].set(display)
            return

        if current:
            normalized = self._format_city_label(current)
            if normalized in display_to_raw:
                self.field_vars[header].set(normalized)

    def _get_field_type(self, header):
        norm = self._normalize_header(header)

        if norm in {"date"}:
            return "date"
        if norm in {"id", "id client", "id_client", "ref client", "reference", "ref", "ref. client"}:
            return "client_id"
        if norm in {"nom", "nom client", "client nom", "nom du client"}:
            return "client_name"
        if norm.startswith("ville") and "depart" in norm:
            return "departure_city"
        if norm.startswith("ville") and "arriv" in norm:
            return "arrival_city"
        if norm in {"nombre adulte", "nombre adultes", "adultes", "adulte", "nb adultes", "nb adulte", "nombres adultes"}:
            return "adult_count"
        if norm in {"nombre enfant", "nombre enfants", "nombres enfants", "enfant", "enfants", "nb enfants", "nb enfant"}:
            return "child_count"
        if norm in {"tarif adulte", "tarif adultes", "prix adulte", "prix adultes"}:
            return "adult_tarif"
        if norm in {
            "tarif enfant",
            "tarifs enfants",
            "tarif enfants",
            "prix enfant",
            "prix enfants",
            "tarif bebe",
            "tarif bebes",
            "tarif bébé",
            "tarif bébés",
            "prix bebe",
            "prix bebes",
            "prix bébé",
            "prix bébés",
        }:
            return "child_tarif"
        if norm in {"montant adulte", "montant adultes", "total adultes", "montant adulte total"}:
            return "adult_amount"
        if norm in {"montant enfant", "montant enfants", "total enfants", "montant enfant total"}:
            return "child_amount"
        if norm in {"total", "montant total", "total prix"}:
            return "total"
        return "text"

    def _find_header_by_type(self, field_type):
        for header, (ftype, _) in self.field_widgets.items():
            if ftype == field_type:
                return header
        return None

    def _create_form(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        title_text = "MODIFIER RÉSERVATION BILLET AVION" if self.edit_data else "RÉSERVATION BILLET AVION"
        tk.Label(
            self.parent,
            text=title_text,
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(20, 10), fill="x")

        tk.Label(
            self.parent,
            text="1) Sélectionnez le client  2) Choisissez départ/arrivée  3) Vérifiez quantités et total puis Enregistrer.",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 10))

        form_frame = tk.LabelFrame(
            self.parent,
            text="Informations billet avion",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        form_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.field_vars = {}
        self.field_widgets = {}

        display_headers = self._build_form_layout_headers()

        for index, header in enumerate(display_headers):
            row = index // 2
            col_group = index % 2
            label_col = col_group * 2
            entry_col = label_col + 1

            tk.Label(
                form_frame,
                text=f"{self._display_label(header)} :",
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
                names = sorted(self.clients_by_name.keys())
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[""] + names,
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_client_name_changed)
            elif field_type == "departure_city":
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[""],
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_city_changed)
            elif field_type == "arrival_city":
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=[""],
                    font=ENTRY_FONT,
                    width=30,
                    state="readonly",
                )
                widget.bind("<<ComboboxSelected>>", self._on_city_changed)
            elif field_type in {"adult_count", "child_count", "adult_tarif", "child_tarif", "adult_amount", "child_amount", "total"}:
                widget = tk.Entry(
                    form_frame,
                    textvariable=field_var,
                    font=ENTRY_FONT,
                    width=32,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                    state="readonly",
                )
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

    def _to_int(self, value):
        try:
            if value in (None, ""):
                return 0
            return int(float(str(value).strip()))
        except Exception:
            return 0

    def _to_float(self, value):
        try:
            if value in (None, ""):
                return 0.0
            return float(str(value).replace(" ", "").replace(",", "."))
        except Exception:
            return 0.0

    def _build_tarif_filters(self):
        filters = {}
        for header, var in self.field_vars.items():
            field_type = self._get_field_type(header)
            if field_type in {
                "date",
                "client_id",
                "client_name",
                "adult_count",
                "child_count",
                "adult_tarif",
                "child_tarif",
                "adult_amount",
                "child_amount",
                "total",
            }:
                continue
            value = var.get().strip()
            if value:
                if field_type in {"departure_city", "arrival_city"} and header in self.field_widgets:
                    _, widget = self.field_widgets[header]
                    value = getattr(widget, "_display_to_raw", {}).get(value, value)
                filters[header] = value
        return filters

    def _sync_city_dropdowns(self):
        departure_header = self._find_header_by_type("departure_city")
        arrival_header = self._find_header_by_type("arrival_city")

        if departure_header and departure_header in self.field_widgets:
            current_departure = self.field_vars[departure_header].get().strip()
            _, dep_widget = self.field_widgets[departure_header]
            current_departure_raw = self._resolve_city_raw(dep_widget, current_departure)
            arrival_value = ""
            if arrival_header:
                arrival_value = self.field_vars[arrival_header].get().strip()
                if arrival_header in self.field_widgets:
                    _, arr_widget_for_dep = self.field_widgets[arrival_header]
                    arrival_value = self._resolve_city_raw(arr_widget_for_dep, arrival_value)
            departure_options = get_avion_departure_cities(
                {arrival_header: arrival_value} if arrival_header and arrival_value else None
            )

            if current_departure_raw and all(str(current_departure_raw).strip().lower() != str(opt).strip().lower() for opt in departure_options):
                departure_options = departure_options + [current_departure_raw]

            self._set_city_options(dep_widget, departure_options)
            if current_departure_raw and all(str(current_departure_raw).strip().lower() != str(opt).strip().lower() for opt in departure_options):
                self.field_vars[departure_header].set("")
            else:
                self._normalize_city_selection(departure_header)

        if arrival_header and arrival_header in self.field_widgets:
            current_arrival = self.field_vars[arrival_header].get().strip()
            _, arr_widget = self.field_widgets[arrival_header]
            current_arrival_raw = self._resolve_city_raw(arr_widget, current_arrival)
            departure_value = ""
            if departure_header:
                departure_value = self.field_vars[departure_header].get().strip()
                if departure_header in self.field_widgets:
                    _, dep_widget_for_arr = self.field_widgets[departure_header]
                    departure_value = self._resolve_city_raw(dep_widget_for_arr, departure_value)
            arrival_options = get_avion_arrival_cities(
                {departure_header: departure_value} if departure_header and departure_value else None
            )

            if current_arrival_raw and all(str(current_arrival_raw).strip().lower() != str(opt).strip().lower() for opt in arrival_options):
                arrival_options = arrival_options + [current_arrival_raw]

            self._set_city_options(arr_widget, arrival_options)
            if current_arrival_raw and all(str(current_arrival_raw).strip().lower() != str(opt).strip().lower() for opt in arrival_options):
                self.field_vars[arrival_header].set("")
            else:
                self._normalize_city_selection(arrival_header)

    def _on_city_changed(self, event=None):
        self._sync_city_dropdowns()
        self._update_tarifs_and_totals()

    def _apply_client(self, client):
        if not client:
            return

        id_header = self._find_header_by_type("client_id")
        name_header = self._find_header_by_type("client_name")
        adult_header = self._find_header_by_type("adult_count")
        child_header = self._find_header_by_type("child_count")

        if id_header:
            self.field_vars[id_header].set(str(client.get("ref_client") or "").strip())
        if name_header:
            self.field_vars[name_header].set(str(client.get("nom") or "").strip())
        if adult_header:
            adultes = self._to_int(client.get("nombre_adultes"))
            self.field_vars[adult_header].set(str(adultes))
        if child_header:
            enfants = self._to_int(client.get("nombre_enfants_2_12"))
            bebes = self._to_int(client.get("nombre_bebes_0_2"))
            self.field_vars[child_header].set(str(enfants + bebes))

        self._sync_city_dropdowns()
        self._update_tarifs_and_totals()

    def _on_client_id_changed(self, event=None):
        try:
            id_header = self._find_header_by_type("client_id")
            if not id_header:
                return
            client_id = self.field_vars[id_header].get().strip()
            client = self.client_map.get(client_id)
            self._apply_client(client)
        except Exception as e:
            logger.error(f"Error on client id change: {e}", exc_info=True)

    def _on_client_name_changed(self, event=None):
        try:
            name_header = self._find_header_by_type("client_name")
            if not name_header:
                return
            name = self.field_vars[name_header].get().strip()
            client = self.clients_by_name.get(name)
            self._apply_client(client)
        except Exception as e:
            logger.error(f"Error on client name change: {e}", exc_info=True)

    def _update_tarifs_and_totals(self):
        try:
            departure_header = self._find_header_by_type("departure_city")
            arrival_header = self._find_header_by_type("arrival_city")
            adult_header = self._find_header_by_type("adult_count")
            child_header = self._find_header_by_type("child_count")
            adult_tarif_header = self._find_header_by_type("adult_tarif")
            child_tarif_header = self._find_header_by_type("child_tarif")
            adult_amount_header = self._find_header_by_type("adult_amount")
            child_amount_header = self._find_header_by_type("child_amount")
            total_header = self._find_header_by_type("total")

            departure_selected = True
            if departure_header:
                departure_selected = bool(self.field_vars.get(departure_header, tk.StringVar()).get().strip())

            arrival_selected = True
            if arrival_header:
                arrival_selected = bool(self.field_vars.get(arrival_header, tk.StringVar()).get().strip())

            if not (departure_selected and arrival_selected):
                if adult_tarif_header:
                    self.field_vars[adult_tarif_header].set("")
                if child_tarif_header:
                    self.field_vars[child_tarif_header].set("")
                if adult_amount_header:
                    self.field_vars[adult_amount_header].set("")
                if child_amount_header:
                    self.field_vars[child_amount_header].set("")
                if total_header:
                    self.field_vars[total_header].set("")
                return

            adult_count = self._to_int(self.field_vars.get(adult_header, tk.StringVar()).get() if adult_header else 0)
            child_count = self._to_int(self.field_vars.get(child_header, tk.StringVar()).get() if child_header else 0)

            tarif_adulte, tarif_enfant = get_avion_tarifs(self._build_tarif_filters())
            montant_adultes = adult_count * self._to_float(tarif_adulte)
            montant_enfants = child_count * self._to_float(tarif_enfant)
            total = montant_adultes + montant_enfants

            if adult_tarif_header:
                self.field_vars[adult_tarif_header].set(str(self._to_float(tarif_adulte)) if self._to_float(tarif_adulte) else "")
            if child_tarif_header:
                self.field_vars[child_tarif_header].set(str(self._to_float(tarif_enfant)) if self._to_float(tarif_enfant) else "")
            if adult_amount_header:
                self.field_vars[adult_amount_header].set(str(montant_adultes) if montant_adultes else "")
            if child_amount_header:
                self.field_vars[child_amount_header].set(str(montant_enfants) if montant_enfants else "")
            if total_header:
                self.field_vars[total_header].set(str(total) if total else "")
        except Exception as e:
            logger.error(f"Error updating tarifs and totals: {e}", exc_info=True)

    def _load_edit_data(self):
        try:
            for header, var in self.field_vars.items():
                var.set(str(self.edit_data.get(header, "")).strip())
            self._sync_city_dropdowns()
            has_tarif_or_total = any(
                str(self.field_vars.get(header, tk.StringVar()).get() if header else "").strip()
                for header in [
                    self._find_header_by_type("adult_tarif"),
                    self._find_header_by_type("child_tarif"),
                    self._find_header_by_type("adult_amount"),
                    self._find_header_by_type("child_amount"),
                    self._find_header_by_type("total"),
                ]
            )
            if not has_tarif_or_total:
                self._update_tarifs_and_totals()
        except Exception as e:
            logger.error(f"Error loading edit data: {e}", exc_info=True)

    def _collect_form_data(self):
        return {header: var.get().strip() for header, var in self.field_vars.items()}

    def _save(self):
        form_data = self._collect_form_data()
        has_data = any(str(v).strip() for k, v in form_data.items() if self._normalize_header(k) != "date")
        if not has_data:
            messagebox.showwarning("Validation", "Veuillez renseigner au moins un champ avant l'enregistrement.")
            return

        client_id_header = self._find_header_by_type("client_id")
        client_name_header = self._find_header_by_type("client_name")
        client_id = form_data.get(client_id_header, "").strip() if client_id_header else ""
        client_name = form_data.get(client_name_header, "").strip() if client_name_header else ""
        if not client_id and not client_name:
            messagebox.showwarning(
                "Validation",
                "Veuillez sélectionner un client avant d'enregistrer la réservation.",
            )
            return

        if self.edit_data and self.row_number is not None:
            result = update_air_ticket_quotation_in_excel(self.row_number, form_data)
            if result == -2:
                messagebox.showerror("Fichier verrouillé", "Fermez data.xlsx puis réessayez.")
                return
            if result == -1:
                messagebox.showerror("Erreur", "Échec de la modification dans AVION.")
                return
            messagebox.showinfo("Succès", "Réservation modifiée avec succès.")
            if self.callback_on_done:
                self.callback_on_done()
            return

        row = save_air_ticket_quotation_to_excel(form_data)
        if row == -2:
            messagebox.showerror("Fichier verrouillé", "Fermez data.xlsx puis réessayez.")
            return
        if row == -1:
            messagebox.showerror("Erreur", "Échec de l'enregistrement dans AVION.")
            return

        messagebox.showinfo("Succès", f"Réservation enregistrée avec succès à la ligne {row}.")
        if self.callback_on_done:
            self.callback_on_done()
        else:
            self._clear()

    def _delete(self):
        if self.row_number is None:
            return

        confirm = messagebox.askyesno(
            "Confirmation",
            "Êtes-vous sûr de vouloir supprimer cette réservation ?\n\nCette action ne peut pas être annulée.",
        )
        if not confirm:
            return

        success = delete_air_ticket_from_excel(self.row_number)
        if not success:
            messagebox.showerror("Erreur", "Impossible de supprimer. Vérifiez que data.xlsx n'est pas ouvert.")
            return

        messagebox.showinfo("Succès", "Réservation supprimée avec succès.")
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
        self._sync_city_dropdowns()

    def _reload_from_sheet(self):
        self._load_base_data()
        self._load_headers()
        self._create_form()
