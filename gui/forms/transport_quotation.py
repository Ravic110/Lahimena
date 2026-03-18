"""
Transport quotation GUI component
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
    READONLY_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
)
from gui.forms.client_form import CalendarDialog
from utils.excel_handler import (
    get_km_mada_duration_for_repere,
    get_km_mada_km_for_repere,
    get_km_mada_reperes,
    get_transport_fuel_price,
    get_transport_headers,
    get_transport_prestataires,
    get_transport_vehicle_data,
    get_transport_vehicle_types,
    load_all_clients,
    save_transport_quotation_to_excel,
    update_transport_quotation_in_excel,
    delete_transport_from_excel,
)
from utils.logger import logger


class TransportQuotation:
    """Transport quotation form with source-driven dropdowns and automatic calculations."""

    REF_SPEED_KMH = 45.0
    MIN_ROUTE_FACTOR = 0.85
    MAX_ROUTE_FACTOR = 1.70

    DEFAULT_HEADERS = [
        "ID_client",
        "Nom client",
        "Numéro dossier",
        "Jour",
        "Date",
        "Lieu de départ",
        "Lieu d'arrivée",
        "Prestataire",
        "Type de voiture",
        "Nombre de place",
        "Location par jour",
        "Distances étapes",
        "Besoin en litre",
        "Budget carburant",
    ]

    def __init__(self, parent, edit_data=None, row_number=None, callback_on_save=None):
        self.parent = parent
        self.edit_data = edit_data
        self.row_number = row_number
        self.callback_on_save = callback_on_save

        self.headers = []
        self.field_vars = {}
        self.field_widgets = {}

        self.clients = []
        self.client_map = {}
        self.clients_by_name = {}
        self.clients_by_dossier = {}

        self.prestataires = []
        self.reperes = []
        self._km_data_warning_shown = False

        self._load_base_data()
        self._load_headers()
        self._create_form()
        if self.edit_data:
            self._load_edit_data()

    def _load_base_data(self):
        self.clients = load_all_clients()
        self.client_map = {}
        self.clients_by_name = {}
        self.clients_by_dossier = {}

        for client in self.clients:
            ref_client = str(client.get("ref_client") or "").strip()
            nom = str(client.get("nom") or "").strip()
            dossier = str(client.get("numero_dossier") or "").strip()

            if ref_client:
                self.client_map[ref_client] = client
            if nom and nom not in self.clients_by_name:
                self.clients_by_name[nom] = client
            if dossier and dossier not in self.clients_by_dossier:
                self.clients_by_dossier[dossier] = client

        self.prestataires = get_transport_prestataires()
        self.reperes = get_km_mada_reperes()

    def _has_repere(self, repere):
        lookup = str(repere or "").strip().lower()
        if not lookup:
            return False
        return any(str(value or "").strip().lower() == lookup for value in self.reperes)

    def _load_headers(self):
        headers = get_transport_headers()
        self.headers = headers if headers else self.DEFAULT_HEADERS.copy()

        required_fields = [
            ("client_id", "ID_client"),
            ("client_name", "Nom client"),
            ("dossier_number", "Numéro dossier"),
            ("day", "Jour"),
            ("date", "Date"),
            ("departure", "Lieu de départ"),
            ("arrival", "Lieu d'arrivée"),
            ("prestataire", "Prestataire"),
            ("vehicle_type", "Type de voiture"),
            ("seat_count", "Nombre de place"),
            ("location_per_day", "Location par jour"),
            ("distance", "Distances étapes"),
            ("fuel_need", "Besoin en litre"),
            ("fuel_budget", "Budget carburant"),
        ]

        for field_type, fallback_label in reversed(required_fields):
            has_field = any(self._get_field_type(header) == field_type for header in self.headers)
            if not has_field:
                self.headers.insert(0, fallback_label)

    def _normalize_header(self, header):
        normalized = (
            str(header)
            .strip()
            .lower()
            .replace("_", " ")
            .replace("-", " ")
            .replace("/", " ")
            .replace("'", " ")
            .replace("’", " ")
        )
        for char, replacement in [
            ("é", "e"),
            ("è", "e"),
            ("ê", "e"),
            ("à", "a"),
            ("â", "a"),
            ("ù", "u"),
            ("û", "u"),
            ("î", "i"),
            ("ï", "i"),
            ("ô", "o"),
            ("ç", "c"),
            ("º", "o"),
            ("°", "o"),
        ]:
            normalized = normalized.replace(char, replacement)
        normalized = re.sub(r"\s+", " ", normalized).strip()
        return normalized

    def _get_field_type(self, header):
        norm = self._normalize_header(header)

        if norm in {"id client", "id_client", "idclient", "ref client", "reference"}:
            return "client_id"
        if norm in {"nom", "nom client", "client nom", "nom du client"}:
            return "client_name"
        if "dossier" in norm:
            return "dossier_number"
        if norm in {"jour", "jours"}:
            return "day"
        if norm == "date":
            return "date"
        if "lieu" in norm and "depart" in norm:
            return "departure"
        if "lieu" in norm and ("arrive" in norm or "arrivee" in norm):
            return "arrival"
        if norm in {"prestataire", "fournisseur", "provider"}:
            return "prestataire"
        if "type" in norm and "voiture" in norm:
            return "vehicle_type"
        if "nombre" in norm and "place" in norm:
            return "seat_count"
        if "location" in norm and "jour" in norm:
            return "location_per_day"
        if "distance" in norm:
            return "distance"
        if "besoin" in norm and "litre" in norm:
            return "fuel_need"
        if "budget" in norm and "carburant" in norm:
            return "fuel_budget"
        if "energie" in norm or "énergie" in norm:
            return "energy"
        return "text"

    def _find_header_by_type(self, field_type):
        for header, (ftype, _widget) in self.field_widgets.items():
            if ftype == field_type:
                return header
        return None

    def _create_form(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        title_text = "MODIFIER COTATION TRANSPORT" if self.edit_data else "COTATION TRANSPORT"
        tk.Label(
            self.parent,
            text=title_text,
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(20, 10), fill="x")

        tk.Label(
            self.parent,
            text="1) Sélectionnez le client  2) Choisissez le trajet/transport  3) Vérifiez distance/budget puis Enregistrer.",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 10))

        form_frame = tk.LabelFrame(
            self.parent,
            text="Informations transport",
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
            widget = self._create_field_widget(form_frame, field_var, field_type)
            widget.grid(row=row, column=entry_col, sticky="w", padx=(0, 12), pady=6)

            self.field_vars[header] = field_var
            self.field_widgets[header] = (field_type, widget)

        button_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        button_frame.pack(fill="x", padx=20, pady=(0, 12))

        tk.Button(
            button_frame,
            text="💾 Enregistrer",
            command=self._save,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            button_frame,
            text="🔄 Recharger",
            command=self._reload_from_sheet,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            button_frame,
            text="🧹 Réinitialiser",
            command=self._clear,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left", padx=(0, 8))

        if self.edit_data:
            tk.Button(
                button_frame,
                text="🗑️ Supprimer",
                command=self._delete,
                bg=BUTTON_RED,
                fg="white",
                font=BUTTON_FONT,
                padx=12,
                pady=5,
            ).pack(side="left", padx=(0, 8))

            tk.Button(
                button_frame,
                text="↩ Annuler",
                command=self._cancel,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=12,
                pady=5,
            ).pack(side="left")

        self._bind_dynamic_events()

    def _create_field_widget(self, parent, field_var, field_type):
        if field_type == "client_id":
            values = [""] + sorted(self.client_map.keys())
            widget = ttk.Combobox(parent, textvariable=field_var, values=values, font=ENTRY_FONT, width=30, state="readonly")
        elif field_type == "client_name":
            values = [""] + sorted(self.clients_by_name.keys())
            widget = ttk.Combobox(parent, textvariable=field_var, values=values, font=ENTRY_FONT, width=30, state="readonly")
        elif field_type == "dossier_number":
            values = [""] + sorted(self.clients_by_dossier.keys())
            widget = ttk.Combobox(parent, textvariable=field_var, values=values, font=ENTRY_FONT, width=30, state="readonly")
        elif field_type in {"departure", "arrival"}:
            values = [""] + self.reperes
            widget = ttk.Combobox(parent, textvariable=field_var, values=values, font=ENTRY_FONT, width=30, state="readonly")
        elif field_type == "prestataire":
            values = [""] + self.prestataires
            widget = ttk.Combobox(parent, textvariable=field_var, values=values, font=ENTRY_FONT, width=30, state="readonly")
        elif field_type == "vehicle_type":
            widget = ttk.Combobox(parent, textvariable=field_var, values=[""], font=ENTRY_FONT, width=30, state="readonly")
        elif field_type == "date":
            frame = tk.Frame(parent, bg=MAIN_BG_COLOR)
            entry = tk.Entry(
                frame,
                textvariable=field_var,
                font=ENTRY_FONT,
                width=27,
                bg=READONLY_BG_COLOR,
                readonlybackground=READONLY_BG_COLOR,
                fg=TEXT_COLOR,
                state="readonly",
            )
            entry.pack(side="left")
            tk.Button(
                frame,
                text="📅",
                font=("Poppins", 12),
                width=3,
                bg=BUTTON_GREEN,
                fg="white",
                command=lambda: self._open_calendar(field_var),
            ).pack(side="left", padx=(5, 0))
            return frame
        else:
            state = "normal"
            bg = INPUT_BG_COLOR
            if field_type in {"seat_count", "location_per_day", "distance", "fuel_need", "fuel_budget", "energy"}:
                state = "readonly"
                bg = READONLY_BG_COLOR

            widget = tk.Entry(
                parent,
                textvariable=field_var,
                font=ENTRY_FONT,
                width=32,
                bg=bg,
                fg=TEXT_COLOR,
                state=state,
                readonlybackground=READONLY_BG_COLOR,
            )
        return widget

    def _bind_dynamic_events(self):
        self._bind_combobox("client_id", self._on_client_id_changed)
        self._bind_combobox("client_name", self._on_client_name_changed)
        self._bind_combobox("dossier_number", self._on_dossier_changed)
        self._bind_combobox("prestataire", self._on_prestataire_changed)
        self._bind_combobox("vehicle_type", self._on_vehicle_changed)
        self._bind_combobox("departure", self._on_route_changed)
        self._bind_combobox("arrival", self._on_route_changed)

        day_header = self._find_header_by_type("day")
        if day_header:
            self.field_vars[day_header].trace_add("write", self._on_day_changed)

    def _bind_combobox(self, field_type, callback):
        header = self._find_header_by_type(field_type)
        if not header:
            return
        _ftype, widget = self.field_widgets[header]
        if isinstance(widget, ttk.Combobox):
            widget.bind("<<ComboboxSelected>>", callback)

    def _open_calendar(self, field_var):
        cal_dialog = CalendarDialog(self.parent, "Choisir une date")
        self.parent.wait_window(cal_dialog)
        if cal_dialog.selected_date:
            field_var.set(cal_dialog.selected_date.strftime("%d/%m/%Y"))

    def _get_client_from_id(self, client_id):
        return self.client_map.get(str(client_id or "").strip())

    def _get_client_from_name(self, name):
        return self.clients_by_name.get(str(name or "").strip())

    def _get_client_from_dossier(self, dossier):
        return self.clients_by_dossier.get(str(dossier or "").strip())

    def _fill_client_fields(self, client):
        if not client:
            return

        header_id = self._find_header_by_type("client_id")
        header_name = self._find_header_by_type("client_name")
        header_dossier = self._find_header_by_type("dossier_number")

        if header_id:
            self.field_vars[header_id].set(str(client.get("ref_client") or "").strip())
        if header_name:
            self.field_vars[header_name].set(str(client.get("nom") or "").strip())
        if header_dossier:
            self.field_vars[header_dossier].set(str(client.get("numero_dossier") or "").strip())

    def _on_client_id_changed(self, *_args):
        header = self._find_header_by_type("client_id")
        if not header:
            return
        client = self._get_client_from_id(self.field_vars[header].get())
        self._fill_client_fields(client)

    def _on_client_name_changed(self, *_args):
        header = self._find_header_by_type("client_name")
        if not header:
            return
        client = self._get_client_from_name(self.field_vars[header].get())
        self._fill_client_fields(client)

    def _on_dossier_changed(self, *_args):
        header = self._find_header_by_type("dossier_number")
        if not header:
            return
        client = self._get_client_from_dossier(self.field_vars[header].get())
        self._fill_client_fields(client)

    def _on_day_changed(self, *_args):
        header = self._find_header_by_type("day")
        if not header:
            return
        value = self.field_vars[header].get().strip()
        if value and not value.isdigit():
            sanitized = "".join(ch for ch in value if ch.isdigit())
            self.field_vars[header].set(sanitized)

    def _on_prestataire_changed(self, *_args):
        prestataire_header = self._find_header_by_type("prestataire")
        vehicle_header = self._find_header_by_type("vehicle_type")
        if not prestataire_header or not vehicle_header:
            return

        prestataire = self.field_vars[prestataire_header].get().strip()
        _ftype, vehicle_widget = self.field_widgets[vehicle_header]
        vehicle_values = [""] + get_transport_vehicle_types(prestataire)
        if isinstance(vehicle_widget, ttk.Combobox):
            vehicle_widget["values"] = vehicle_values

        current_vehicle = self.field_vars[vehicle_header].get().strip()
        if current_vehicle and current_vehicle not in vehicle_values:
            self.field_vars[vehicle_header].set("")

        self._update_transport_auto_fields()

    def _on_vehicle_changed(self, *_args):
        self._update_transport_auto_fields()

    def _on_route_changed(self, *_args):
        self._update_distance_and_budget()

    def _set_field_value(self, field_type, value):
        header = self._find_header_by_type(field_type)
        if not header:
            return
        self.field_vars[header].set("" if value in (None, "") else str(value))

    def _to_float(self, value):
        if value in (None, ""):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        text = str(value).strip().replace(" ", "").replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return 0.0

    def _format_number(self, value):
        numeric = self._to_float(value)
        return f"{numeric:.2f}" if numeric else ""

    def _compute_route_consumption_factor(self, distance_km, duration_hours):
        distance = self._to_float(distance_km)
        duration = self._to_float(duration_hours)
        if distance <= 0 or duration <= 0:
            return 1.0

        expected_duration = distance / self.REF_SPEED_KMH if self.REF_SPEED_KMH > 0 else 0
        if expected_duration <= 0:
            return 1.0

        ratio = duration / expected_duration
        if ratio < self.MIN_ROUTE_FACTOR:
            return self.MIN_ROUTE_FACTOR
        if ratio > self.MAX_ROUTE_FACTOR:
            return self.MAX_ROUTE_FACTOR
        return ratio

    def _update_transport_auto_fields(self):
        prestataire_header = self._find_header_by_type("prestataire")
        vehicle_header = self._find_header_by_type("vehicle_type")
        if not prestataire_header or not vehicle_header:
            return

        prestataire = self.field_vars[prestataire_header].get().strip()
        vehicle = self.field_vars[vehicle_header].get().strip()
        if not prestataire or not vehicle:
            self._set_field_value("seat_count", "")
            self._set_field_value("location_per_day", "")
            self._set_field_value("energy", "")
            self._update_distance_and_budget()
            return

        data = get_transport_vehicle_data(prestataire, vehicle)
        seats = data.get("nombre_place", "")
        location = data.get("location_par_jour", "")
        energy = data.get("energie", "")

        self._set_field_value("seat_count", int(seats) if seats else "")
        self._set_field_value("location_per_day", self._format_number(location))
        self._set_field_value("energy", energy)

        self._update_distance_and_budget()

    def _update_distance_and_budget(self):
        departure_header = self._find_header_by_type("departure")
        arrival_header = self._find_header_by_type("arrival")
        prestataire_header = self._find_header_by_type("prestataire")
        vehicle_header = self._find_header_by_type("vehicle_type")

        departure = self.field_vars[departure_header].get().strip() if departure_header else ""
        arrival = self.field_vars[arrival_header].get().strip() if arrival_header else ""

        distance = 0.0
        duration = 0.0
        if departure and arrival:
            if (not self._has_repere(departure) or not self._has_repere(arrival)) and not self._km_data_warning_shown:
                self._km_data_warning_shown = True
                messagebox.showwarning(
                    "Référentiel KM_MADA",
                    "Certains repères KM_MADA sont introuvables. Les calculs distance/durée utilisent 0 pour ces étapes.",
                )
            km_depart = self._to_float(get_km_mada_km_for_repere(departure))
            km_arrivee = self._to_float(get_km_mada_km_for_repere(arrival))
            distance = abs(km_arrivee - km_depart)

            dur_depart = self._to_float(get_km_mada_duration_for_repere(departure))
            dur_arrivee = self._to_float(get_km_mada_duration_for_repere(arrival))
            duration = abs(dur_arrivee - dur_depart)

        self._set_field_value("distance", self._format_number(distance))

        prestataire = self.field_vars[prestataire_header].get().strip() if prestataire_header else ""
        vehicle = self.field_vars[vehicle_header].get().strip() if vehicle_header else ""

        consommation = 0.0
        energie = ""
        if prestataire and vehicle:
            vehicle_data = get_transport_vehicle_data(prestataire, vehicle)
            consommation = self._to_float(vehicle_data.get("consommation", 0))
            energie = str(vehicle_data.get("energie") or "").strip()

        facteur_route = self._compute_route_consumption_factor(distance, duration)
        besoin_base = (consommation * distance) / 100 if consommation and distance else 0.0
        besoin = besoin_base * facteur_route if besoin_base else 0.0
        self._set_field_value("fuel_need", self._format_number(besoin))

        fuel_price = self._to_float(get_transport_fuel_price(energie)) if energie else 0.0
        budget = besoin * fuel_price if besoin and fuel_price else 0.0
        self._set_field_value("fuel_budget", self._format_number(budget))

    def _load_edit_data(self):
        try:
            for header, var in self.field_vars.items():
                var.set(str(self.edit_data.get(header, "")).strip())

            self._on_prestataire_changed()
            self._update_distance_and_budget()
        except (KeyError, TypeError, ValueError, AttributeError) as e:
            logger.error(f"Error loading transport edit data: {e}", exc_info=True)

    def _collect_form_data(self):
        return {header: var.get().strip() for header, var in self.field_vars.items()}

    def _save(self):
        form_data = self._collect_form_data()
        has_data = any(str(v).strip() for v in form_data.values())
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
                "Veuillez sélectionner un client avant d'enregistrer la cotation.",
            )
            return

        if self.edit_data and self.row_number is not None:
            result = update_transport_quotation_in_excel(self.row_number, form_data)
            if result == -2:
                messagebox.showerror("Fichier verrouillé", "Fermez data.xlsx puis réessayez.")
                return
            if result == -1:
                messagebox.showerror("Erreur", "Échec de la modification dans TRANSPORT.")
                return
            messagebox.showinfo("Succès", "Transport modifié avec succès.")
            if self.callback_on_save:
                self.callback_on_save()
            return

        row = save_transport_quotation_to_excel(form_data)
        if row == -2:
            messagebox.showerror("Fichier verrouillé", "Fermez data.xlsx puis réessayez.")
            return
        if row == -1:
            messagebox.showerror("Erreur", "Échec de l'enregistrement dans TRANSPORT.")
            return

        messagebox.showinfo("Succès", f"Transport enregistré à la ligne {row}.")
        if self.callback_on_save:
            self.callback_on_save()
        else:
            self._clear()

    def _delete(self):
        if self.row_number is None:
            return

        confirm = messagebox.askyesno(
            "Confirmation",
            "Êtes-vous sûr de vouloir supprimer cet enregistrement transport ?\n\nCette action ne peut pas être annulée.",
        )
        if not confirm:
            return

        success = delete_transport_from_excel(self.row_number)
        if not success:
            messagebox.showerror("Erreur", "Impossible de supprimer. Vérifiez que data.xlsx n'est pas ouvert.")
            return

        messagebox.showinfo("Succès", "Enregistrement transport supprimé avec succès.")
        if self.callback_on_save:
            self.callback_on_save()

    def _cancel(self):
        if self.callback_on_save:
            self.callback_on_save()

    def _clear(self):
        for header, var in self.field_vars.items():
            if self._get_field_type(header) == "date":
                var.set(datetime.now().strftime("%d/%m/%Y"))
            else:
                var.set("")
        self._update_distance_and_budget()

    def _reload_from_sheet(self):
        self._load_base_data()
        self._load_headers()
        self._create_form()
