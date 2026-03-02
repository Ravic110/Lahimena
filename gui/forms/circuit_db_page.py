"""Circuit DB management page (data-hotel.xlsx / Circuits)."""

import math
import re
import tkinter as tk
import unicodedata
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
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.excel_handler import (
    delete_circuit_db_row,
    get_collective_expense_prestataires,
    get_km_mada_duration_for_repere,
    get_km_mada_km_for_repere,
    get_transport_prestataires,
    get_transport_vehicle_data,
    get_transport_vehicle_types,
    load_all_hotels,
    get_circuit_db_headers,
    load_circuit_db_rows,
    save_circuit_db_row,
    update_circuit_db_row,
)
from utils.logger import logger


class CircuitDBPage:
    """Circuits DB CRUD page."""

    def __init__(self, parent, on_back_to_db=None):
        self.parent = parent
        self.on_back_to_db = on_back_to_db
        self.headers = []
        self.rows = []
        self.filtered_rows = []
        self.vars = {}
        self.selected_row_number = None
        self.search_var = tk.StringVar()

        self.tree = None
        self.form_frame = None
        self.status_label = None
        self.btn_edit = None
        self.btn_delete = None
        self.itinerary_header = None
        self.cities_header = None
        self.default_hotels_header = None
        self.included_services_header = None
        self.linked_transports_header = None
        self.prestation_options = []
        self.included_services_selected = []
        self.included_services_choice_var = tk.StringVar()
        self.included_services_combo = None
        self.included_services_listbox = None

        self.hotels = []
        self.city_step_var = tk.StringVar()
        self.hotel_step_var = tk.StringVar()
        self.city_step_combo = None
        self.hotel_step_combo = None
        self.mapping_tree = None
        self.city_hotel_mapping = {}
        self.itinerary_cities_order = []
        self.hotel_display_map = {}
        self.mapping_iid_to_key = {}
        self.transport_options = []
        self.transport_display_map = {}
        self.transport_segment_pairs = []
        self.transport_segment_mapping = {}
        self.transport_iid_to_key = {}
        self.transport_depart_var = tk.StringVar()
        self.transport_arrivee_var = tk.StringVar()
        self.transport_choice_var = tk.StringVar()
        self.transport_depart_combo = None
        self.transport_arrivee_combo = None
        self.transport_choice_combo = None
        self.transport_tree = None
        self.segment_metrics_cache = {}

        self._load_hotel_reference()
        self._load_transport_reference()

        self._create_page()

    def _create_page(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        tk.Label(
            self.parent,
            text="GESTION BASE CIRCUITS (DB)",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(16, 6))

        tk.Label(
            self.parent,
            text="Source: data-hotel.xlsx / feuille Circuits",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 8))

        if self.on_back_to_db:
            tk.Button(
                self.parent,
                text="⬅ Retour vers Bases de données",
                command=self._go_back_to_db,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
            ).pack(anchor="w", padx=16, pady=(0, 8))

        search_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        search_frame.pack(fill="x", padx=16, pady=(0, 8))

        tk.Label(
            search_frame,
            text="Rechercher:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        self.search_var.trace("w", self._on_filter_change)
        tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=ENTRY_FONT,
            width=38,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        ).pack(side="left", padx=(8, 0))

        btn_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        btn_frame.pack(fill="x", padx=16, pady=(0, 8))

        tk.Button(
            btn_frame,
            text="🔄 Actualiser",
            command=self._load_data,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
        ).pack(side="left", padx=4)

        tk.Button(
            btn_frame,
            text="➕ Ajouter",
            command=self._new_row,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
        ).pack(side="left", padx=4)

        self.btn_edit = tk.Button(
            btn_frame,
            text="✏️ Modifier",
            command=self._edit_selected,
            bg=BUTTON_ORANGE,
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
        )
        self.btn_edit.pack(side="left", padx=4)

        self.btn_delete = tk.Button(
            btn_frame,
            text="🗑️ Supprimer",
            command=self._delete_selected,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
        )
        self.btn_delete.pack(side="left", padx=4)

        tree_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal")

        style = ttk.Style()
        style.configure(
            "Treeview",
            background=INPUT_BG_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR,
        )
        style.map("Treeview", background=[("selected", BUTTON_GREEN)])

        self.tree = ttk.Treeview(
            tree_frame,
            columns=(),
            show="headings",
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            style="Treeview",
        )
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)

        self.tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        self.tree.bind("<<TreeviewSelect>>", self._on_selection_change)
        self.tree.bind("<Double-1>", lambda _e: self._edit_selected())

        self.status_label = tk.Label(
            self.parent,
            text="",
            font=("Arial", 10),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        self.status_label.pack(anchor="w", padx=16, pady=(0, 6))

        self.form_frame = tk.LabelFrame(
            self.parent,
            text="Formulaire",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        self.form_frame.pack(fill="x", padx=16, pady=(0, 12))

        self._load_data()

    def _go_back_to_db(self):
        if self.on_back_to_db:
            self.on_back_to_db()

    def _load_hotel_reference(self):
        """Load hotels used for city -> hotel matching."""
        try:
            hotels = load_all_hotels()
        except (FileNotFoundError, PermissionError, OSError, ValueError) as exc:
            logger.warning(f"Unable to load hotel reference for circuits: {exc}")
            hotels = []

        unique_hotels = {}
        for hotel in hotels:
            nom = str(hotel.get("nom", "")).strip()
            lieu = str(hotel.get("lieu", "")).strip()
            categorie = str(hotel.get("categorie", "")).strip()
            key = (
                self._normalize_city(lieu),
                self._normalize_city(nom),
                self._normalize_city(categorie),
            )
            if nom and lieu and key not in unique_hotels:
                unique_hotels[key] = hotel
        self.hotels = list(unique_hotels.values())

    def _load_transport_reference(self):
        """Load transport choices (prestataire + type de voiture)."""
        options = []
        display_map = {}
        seen = set()
        try:
            prestataires = get_transport_prestataires()
        except (FileNotFoundError, PermissionError, OSError, ValueError) as exc:
            logger.warning(f"Unable to load transport prestataires for circuits: {exc}")
            prestataires = []

        for prestataire in prestataires:
            try:
                vehicle_types = get_transport_vehicle_types(prestataire) or []
            except (FileNotFoundError, PermissionError, OSError, ValueError) as exc:
                logger.warning(
                    f"Unable to load vehicle types for prestataire '{prestataire}': {exc}"
                )
                vehicle_types = []
            for vehicle in vehicle_types:
                data = get_transport_vehicle_data(prestataire, vehicle) or {}
                label = f"{prestataire} - {vehicle}"
                key = self._normalize_text(label)
                if not key or key in seen:
                    continue
                seen.add(key)
                options.append(label)
                display_map[label] = {
                    "prestataire": str(prestataire or "").strip(),
                    "type_voiture": str(vehicle or "").strip(),
                    "nombre_place": data.get("nombre_place", 0),
                    "energie": data.get("energie", ""),
                }

        self.transport_options = sorted(options, key=lambda v: v.lower())
        self.transport_display_map = display_map

    def _normalize_text(self, value):
        if not value:
            return ""
        text = unicodedata.normalize("NFKD", str(value).strip().lower())
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        text = re.sub(r"[^a-z0-9]+", " ", text)
        return re.sub(r"\s+", " ", text).strip()

    def _normalize_city(self, value):
        return self._normalize_text(value)

    def _normalize_header(self, value):
        return self._normalize_text(value)

    def _find_header(self, *aliases):
        normalized_headers = {
            self._normalize_header(header): header for header in self.headers
        }
        for alias in aliases:
            found = normalized_headers.get(self._normalize_header(alias))
            if found:
                return found
        return None

    def _load_prestation_options(self):
        """Load prestataire options for the circuit dropdown."""
        options = set()
        try:
            for prestataire in get_collective_expense_prestataires():
                value = str(prestataire).strip()
                if value:
                    options.add(value)
        except (FileNotFoundError, PermissionError, OSError, ValueError) as exc:
            logger.warning(f"Unable to load collective expense prestataires: {exc}")

        if self.included_services_header:
            for row in self.rows:
                value = str(row.get(self.included_services_header, "")).strip()
                for item in re.split(r"[|,;\n]+", value):
                    item = item.strip()
                    if item:
                        options.add(item)

        return sorted(options, key=lambda v: v.lower())

    def _refresh_included_services_widget(self):
        if not self.included_services_listbox:
            return

        self.included_services_listbox.delete(0, tk.END)
        for item in self.included_services_selected:
            self.included_services_listbox.insert(tk.END, item)

        if self.included_services_header and self.included_services_header in self.vars:
            self.vars[self.included_services_header].set(
                " | ".join(self.included_services_selected)
            )

    def _sync_included_services_widget_from_var(self):
        raw = ""
        if self.included_services_header and self.included_services_header in self.vars:
            raw = self.vars[self.included_services_header].get().strip()

        selected = []
        seen = set()
        for part in re.split(r"[|,;\n]+", raw):
            value = part.strip()
            key = self._normalize_text(value)
            if value and key and key not in seen:
                seen.add(key)
                selected.append(value)
        self.included_services_selected = selected
        self._refresh_included_services_widget()

    def _add_included_service(self):
        value = self.included_services_choice_var.get().strip()
        if not value:
            return
        key = self._normalize_text(value)
        existing_keys = {self._normalize_text(v) for v in self.included_services_selected}
        if key in existing_keys:
            self.included_services_choice_var.set("")
            return
        self.included_services_selected.append(value)
        self.included_services_choice_var.set("")
        self._refresh_included_services_widget()

    def _remove_selected_included_service(self):
        if not self.included_services_listbox:
            return
        selected_idx = list(self.included_services_listbox.curselection())
        if not selected_idx:
            return
        for idx in reversed(selected_idx):
            if 0 <= idx < len(self.included_services_selected):
                self.included_services_selected.pop(idx)
        self._refresh_included_services_widget()

    def _parse_city_list(self, raw_value):
        """Parse ordered city list from itinerary text."""
        if not raw_value:
            return []

        raw_segments = re.split(r"[;,>\n|]+", str(raw_value))
        cities = []
        seen = set()
        for segment in raw_segments:
            city = segment.strip()
            if not city:
                continue
            if " - " in city:
                city = city.split(" - ", 1)[1].strip()
            city = re.sub(r"\(\s*\d+\s*(?:j|jour|jours)\s*\)", "", city, flags=re.IGNORECASE)
            city = re.sub(r"\b\d+\s*(?:j|jour|jours)\b", "", city, flags=re.IGNORECASE)
            city = re.sub(r"\s+", " ", city).strip(" -")
            city_key = self._normalize_city(city)
            if city and city_key and city_key not in seen:
                seen.add(city_key)
                cities.append(city)
        return cities

    def _parse_hotels_mapping(self, raw_value):
        """Parse mapping string into ordered city->hotel dictionary."""
        parsed = {}
        if not raw_value:
            return parsed

        for chunk in re.split(r"[|\n;]+", str(raw_value)):
            part = chunk.strip()
            if not part:
                continue

            city = ""
            hotel = ""
            for separator in (":", "=", "->", "=>"):
                if separator in part:
                    left, right = part.split(separator, 1)
                    city = left.strip()
                    hotel = right.strip()
                    break

            if not city or not hotel:
                continue

            key = self._normalize_city(city)
            if key:
                parsed[key] = {"city": city, "hotel": hotel}
        return parsed

    def _parse_transports_mapping(self, raw_value):
        """Parse stored transport mapping string."""
        parsed = {}
        if not raw_value:
            return parsed

        for chunk in re.split(r"[|\n;]+", str(raw_value)):
            part = chunk.strip()
            if not part or ":" not in part:
                continue

            route, transport = part.split(":", 1)
            route = route.strip()
            transport = transport.strip()
            if "->" in route:
                depart, arrivee = [x.strip() for x in route.split("->", 1)]
            elif "-" in route:
                depart, arrivee = [x.strip() for x in route.split("-", 1)]
            else:
                continue
            if not depart or not arrivee or not transport:
                continue
            key = f"{self._normalize_city(depart)}->{self._normalize_city(arrivee)}"
            parsed[key] = {
                "depart": depart,
                "arrivee": arrivee,
                "transport": transport,
            }
        return parsed

    def _format_hotels_mapping(self):
        parts = []
        for city in self.itinerary_cities_order:
            key = self._normalize_city(city)
            entry = self.city_hotel_mapping.get(key)
            if entry and entry.get("hotel"):
                parts.append(f"{entry.get('city', city)}: {entry['hotel']}")
        for key, entry in self.city_hotel_mapping.items():
            if key in {self._normalize_city(c) for c in self.itinerary_cities_order}:
                continue
            if entry.get("city") and entry.get("hotel"):
                parts.append(f"{entry['city']}: {entry['hotel']}")
        return " | ".join(parts)

    def _format_transport_mapping(self):
        parts = []
        ordered_keys = [
            f"{self._normalize_city(dep)}->{self._normalize_city(arr)}"
            for dep, arr in self.transport_segment_pairs
        ]
        for key in ordered_keys:
            entry = self.transport_segment_mapping.get(key)
            if entry and entry.get("transport"):
                parts.append(
                    f"{entry.get('depart', '')} -> {entry.get('arrivee', '')}: {entry.get('transport', '')}"
                )
        for key, entry in self.transport_segment_mapping.items():
            if key in ordered_keys:
                continue
            if entry.get("depart") and entry.get("arrivee") and entry.get("transport"):
                parts.append(
                    f"{entry['depart']} -> {entry['arrivee']}: {entry['transport']}"
                )
        return " | ".join(parts)

    def _sync_itinerary_from_fields(self):
        raw_values = []
        if self.cities_header and self.cities_header in self.vars:
            raw_values.append(self.vars[self.cities_header].get().strip())
        if self.itinerary_header and self.itinerary_header in self.vars:
            raw_values.append(self.vars[self.itinerary_header].get().strip())

        cities = []
        seen = set()
        for raw_value in raw_values:
            for city in self._parse_city_list(raw_value):
                key = self._normalize_city(city)
                if key and key not in seen:
                    seen.add(key)
                    cities.append(city)
        self.itinerary_cities_order = cities
        segment_pairs = []
        for idx in range(len(cities) - 1):
            depart = cities[idx]
            arrivee = cities[idx + 1]
            if depart and arrivee:
                segment_pairs.append((depart, arrivee))
        self.transport_segment_pairs = segment_pairs

    def _hotels_for_city(self, city_name):
        city_key = self._normalize_city(city_name)
        if not city_key:
            return []

        options = []
        self.hotel_display_map = {}
        for hotel in self.hotels:
            if self._normalize_city(hotel.get("lieu")) != city_key:
                continue
            nom = str(hotel.get("nom", "")).strip()
            categorie = str(hotel.get("categorie", "")).strip()
            if not nom:
                continue
            label = f"{nom} ({categorie})" if categorie else nom
            options.append(label)
            self.hotel_display_map[label] = nom
        options.sort(key=lambda x: x.lower())
        return options

    def _refresh_city_hotel_editor(self):
        if not self.city_step_combo or not self.hotel_step_combo:
            return

        self._sync_itinerary_from_fields()
        self.city_step_combo["values"] = self.itinerary_cities_order

        current_city = self.city_step_var.get().strip()
        if current_city not in self.itinerary_cities_order:
            current_city = self.itinerary_cities_order[0] if self.itinerary_cities_order else ""
            self.city_step_var.set(current_city)

        hotel_values = self._hotels_for_city(current_city)
        self.hotel_step_combo["values"] = hotel_values

        key = self._normalize_city(current_city)
        selected_hotel = ""
        if key in self.city_hotel_mapping:
            selected_hotel = self.city_hotel_mapping[key].get("hotel", "")
        if selected_hotel:
            matched = next(
                (
                    label
                    for label in hotel_values
                    if self._normalize_text(label) == self._normalize_text(selected_hotel)
                    or self._normalize_text(self.hotel_display_map.get(label, ""))
                    == self._normalize_text(selected_hotel)
                ),
                "",
            )
            self.hotel_step_var.set(matched or selected_hotel)
        else:
            self.hotel_step_var.set("")

        self._refresh_mapping_tree()
        self._refresh_transport_editor()

    def _refresh_mapping_tree(self):
        if not self.mapping_tree:
            return

        self.mapping_iid_to_key = {}
        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)

        inserted = set()
        for city in self.itinerary_cities_order:
            key = self._normalize_city(city)
            entry = self.city_hotel_mapping.get(key, {})
            iid = key or f"city_{len(self.mapping_tree.get_children())}"
            if self.mapping_tree.exists(iid):
                suffix = 2
                while self.mapping_tree.exists(f"{iid}_{suffix}"):
                    suffix += 1
                iid = f"{iid}_{suffix}"
            self.mapping_tree.insert(
                "",
                "end",
                iid=iid,
                values=(entry.get("city", city), entry.get("hotel", "")),
            )
            self.mapping_iid_to_key[iid] = key
            inserted.add(key)

        for key, entry in self.city_hotel_mapping.items():
            if key in inserted:
                continue
            iid = key or f"city_{len(self.mapping_tree.get_children())}"
            if self.mapping_tree.exists(iid):
                suffix = 2
                while self.mapping_tree.exists(f"{iid}_{suffix}"):
                    suffix += 1
                iid = f"{iid}_{suffix}"
            self.mapping_tree.insert(
                "",
                "end",
                iid=iid,
                values=(entry.get("city", ""), entry.get("hotel", "")),
            )
            self.mapping_iid_to_key[iid] = key

    def _refresh_transport_editor(self):
        if not self.transport_depart_combo or not self.transport_arrivee_combo or not self.transport_choice_combo:
            return

        self._sync_itinerary_from_fields()
        depart_values = []
        seen = set()
        for dep, _arr in self.transport_segment_pairs:
            k = self._normalize_city(dep)
            if k not in seen:
                seen.add(k)
                depart_values.append(dep)

        self.transport_depart_combo["values"] = depart_values
        current_dep = self.transport_depart_var.get().strip()
        if current_dep not in depart_values:
            current_dep = depart_values[0] if depart_values else ""
            self.transport_depart_var.set(current_dep)

        arrivee_values = [
            arr for dep, arr in self.transport_segment_pairs if dep == current_dep
        ]
        self.transport_arrivee_combo["values"] = arrivee_values
        current_arr = self.transport_arrivee_var.get().strip()
        if current_arr not in arrivee_values:
            current_arr = arrivee_values[0] if arrivee_values else ""
            self.transport_arrivee_var.set(current_arr)

        self.transport_choice_combo["values"] = self.transport_options

        key = f"{self._normalize_city(current_dep)}->{self._normalize_city(current_arr)}"
        existing = self.transport_segment_mapping.get(key, {}).get("transport", "")
        self.transport_choice_var.set(existing)
        self._refresh_transport_tree()

    def _refresh_transport_tree(self):
        if not self.transport_tree:
            return

        self.transport_iid_to_key = {}
        for item in self.transport_tree.get_children():
            self.transport_tree.delete(item)

        ordered_keys = []
        for dep, arr in self.transport_segment_pairs:
            key = f"{self._normalize_city(dep)}->{self._normalize_city(arr)}"
            ordered_keys.append(key)
            entry = self.transport_segment_mapping.get(key, {})
            route = f"{entry.get('depart', dep)} -> {entry.get('arrivee', arr)}"
            transport = entry.get("transport", "")
            km, duree = self._lookup_segment_metrics(dep, arr)
            metrics = self._format_segment_metrics(km, duree)
            iid = key or f"segment_{len(self.transport_tree.get_children())}"
            if self.transport_tree.exists(iid):
                suffix = 2
                while self.transport_tree.exists(f"{iid}_{suffix}"):
                    suffix += 1
                iid = f"{iid}_{suffix}"
            self.transport_tree.insert(
                "",
                "end",
                iid=iid,
                values=(route, transport, metrics),
            )
            self.transport_iid_to_key[iid] = key

        for key, entry in self.transport_segment_mapping.items():
            if key in ordered_keys:
                continue
            route = f"{entry.get('depart', '')} -> {entry.get('arrivee', '')}"
            transport = entry.get("transport", "")
            km, duree = self._lookup_segment_metrics(
                entry.get("depart", ""), entry.get("arrivee", "")
            )
            metrics = self._format_segment_metrics(km, duree)
            iid = key or f"segment_{len(self.transport_tree.get_children())}"
            if self.transport_tree.exists(iid):
                suffix = 2
                while self.transport_tree.exists(f"{iid}_{suffix}"):
                    suffix += 1
                iid = f"{iid}_{suffix}"
            self.transport_tree.insert(
                "",
                "end",
                iid=iid,
                values=(route, transport, metrics),
            )
            self.transport_iid_to_key[iid] = key

    def _lookup_segment_metrics(self, depart, arrivee):
        cache_key = f"{self._normalize_city(depart)}->{self._normalize_city(arrivee)}"
        if cache_key in self.segment_metrics_cache:
            return self.segment_metrics_cache[cache_key]

        repere_candidates = [
            f"{depart} - {arrivee}",
            f"{depart}-{arrivee}",
            f"{arrivee} - {depart}",
            f"{arrivee}-{depart}",
        ]
        for repere in repere_candidates:
            km = get_km_mada_km_for_repere(repere)
            duree = get_km_mada_duration_for_repere(repere)
            if km or duree:
                self.segment_metrics_cache[cache_key] = (km, duree)
                return km, duree
        self.segment_metrics_cache[cache_key] = (0, 0)
        return 0, 0

    def _format_segment_metrics(self, km, duree):
        def _to_number(value):
            try:
                number = float(value)
            except (TypeError, ValueError):
                return 0.0
            if not math.isfinite(number):
                return 0.0
            return number

        km_val = _to_number(km)
        duree_val = _to_number(duree)
        parts = []
        if km_val > 0:
            parts.append(f"{int(round(km_val))} km")
        if duree_val > 0:
            parts.append(f"{duree_val:.1f} h")
        return " | ".join(parts)

    def _on_itinerary_fields_changed(self, *_args):
        self._refresh_city_hotel_editor()

    def _on_city_step_selected(self, _event=None):
        self._refresh_city_hotel_editor()

    def _on_transport_depart_selected(self, _event=None):
        self._refresh_transport_editor()

    def _on_transport_arrivee_selected(self, _event=None):
        self._refresh_transport_editor()

    def _assign_hotel_to_city(self):
        city = self.city_step_var.get().strip()
        hotel_label = self.hotel_step_var.get().strip()
        if not city:
            messagebox.showwarning("Validation", "Veuillez sélectionner une ville.")
            return
        if not hotel_label:
            messagebox.showwarning(
                "Validation", "Veuillez sélectionner un hôtel pour cette ville."
            )
            return

        hotel_name = self.hotel_display_map.get(hotel_label, hotel_label)
        key = self._normalize_city(city)
        self.city_hotel_mapping[key] = {"city": city, "hotel": hotel_name}
        if self.default_hotels_header and self.default_hotels_header in self.vars:
            self.vars[self.default_hotels_header].set(self._format_hotels_mapping())
        self._refresh_city_hotel_editor()

    def _remove_selected_city_hotel_mapping(self):
        if not self.mapping_tree:
            return
        selected = self.mapping_tree.selection()
        if not selected:
            return
        for item_id in selected:
            self.city_hotel_mapping.pop(self.mapping_iid_to_key.get(item_id, item_id), None)
        if self.default_hotels_header and self.default_hotels_header in self.vars:
            self.vars[self.default_hotels_header].set(self._format_hotels_mapping())
        self._refresh_city_hotel_editor()

    def _clear_city_hotel_mapping(self):
        self.city_hotel_mapping = {}
        if self.default_hotels_header and self.default_hotels_header in self.vars:
            self.vars[self.default_hotels_header].set("")
        self._refresh_city_hotel_editor()

    def _assign_transport_to_segment(self):
        depart = self.transport_depart_var.get().strip()
        arrivee = self.transport_arrivee_var.get().strip()
        transport = self.transport_choice_var.get().strip()
        if not depart or not arrivee:
            messagebox.showwarning(
                "Validation", "Veuillez sélectionner un départ et une arrivée."
            )
            return
        if not transport:
            messagebox.showwarning("Validation", "Veuillez sélectionner un transport.")
            return

        key = f"{self._normalize_city(depart)}->{self._normalize_city(arrivee)}"
        self.transport_segment_mapping[key] = {
            "depart": depart,
            "arrivee": arrivee,
            "transport": transport,
        }
        if self.linked_transports_header and self.linked_transports_header in self.vars:
            self.vars[self.linked_transports_header].set(self._format_transport_mapping())
        self._refresh_transport_editor()

    def _remove_selected_transport_segment(self):
        if not self.transport_tree:
            return
        selected = self.transport_tree.selection()
        if not selected:
            return
        for item_id in selected:
            self.transport_segment_mapping.pop(
                self.transport_iid_to_key.get(item_id, item_id), None
            )
        if self.linked_transports_header and self.linked_transports_header in self.vars:
            self.vars[self.linked_transports_header].set(self._format_transport_mapping())
        self._refresh_transport_editor()

    def _clear_transport_mapping(self):
        self.transport_segment_mapping = {}
        if self.linked_transports_header and self.linked_transports_header in self.vars:
            self.vars[self.linked_transports_header].set("")
        self.transport_depart_var.set("")
        self.transport_arrivee_var.set("")
        self.transport_choice_var.set("")
        self._refresh_transport_editor()

    def _init_city_hotel_mapping_from_form(self):
        self.city_hotel_mapping = {}
        if self.default_hotels_header and self.default_hotels_header in self.vars:
            self.city_hotel_mapping = self._parse_hotels_mapping(
                self.vars[self.default_hotels_header].get().strip()
            )
        self.transport_segment_mapping = {}
        if self.linked_transports_header and self.linked_transports_header in self.vars:
            self.transport_segment_mapping = self._parse_transports_mapping(
                self.vars[self.linked_transports_header].get().strip()
            )
        self._refresh_city_hotel_editor()

    def _configure_form(self):
        for widget in self.form_frame.winfo_children():
            widget.destroy()

        self.vars = {}
        self.itinerary_header = None
        self.cities_header = None
        self.default_hotels_header = None
        self.included_services_header = None
        self.linked_transports_header = None
        self.city_step_combo = None
        self.hotel_step_combo = None
        self.mapping_tree = None
        self.city_hotel_mapping = {}
        self.itinerary_cities_order = []
        self.mapping_iid_to_key = {}
        self.included_services_selected = []
        self.included_services_choice_var.set("")
        self.included_services_combo = None
        self.included_services_listbox = None
        self.transport_depart_var.set("")
        self.transport_arrivee_var.set("")
        self.transport_choice_var.set("")
        self.transport_depart_combo = None
        self.transport_arrivee_combo = None
        self.transport_choice_combo = None
        self.transport_tree = None
        self.segment_metrics_cache = {}
        self.transport_segment_pairs = []
        self.transport_segment_mapping = {}
        self.transport_iid_to_key = {}
        self.city_step_var.set("")
        self.hotel_step_var.set("")

        self.cities_header = self._find_header(
            "Villes parcourues",
            "Villes du circuit",
            "Villes",
        )
        self.itinerary_header = self._find_header(
            "itinéraire",
            "itineraire",
            "Itinéraire",
        )
        self.default_hotels_header = self._find_header(
            "Hôtels défaut par ville",
            "Hotels defaut par ville",
            "Hôtels par défaut par ville",
            "Hotels par defaut par ville",
            "Hôtels par ville",
            "Hotels par ville",
        )
        self.included_services_header = self._find_header(
            "Prestations incluses",
            "Prestations incluses circuit",
            "Prestations circuit",
        )
        self.linked_transports_header = self._find_header(
            "Transports associés",
            "Transports associes",
            "Transports associés circuit",
            "Transports associes circuit",
        )
        self.prestation_options = self._load_prestation_options()

        for idx, header in enumerate(self.headers):
            row = idx // 2
            col_group = idx % 2
            label_col = col_group * 2
            entry_col = label_col + 1

            tk.Label(
                self.form_frame,
                text=f"{header} :",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=row, column=label_col, sticky="w", padx=(0, 8), pady=4)

            var = tk.StringVar()
            if header == self.included_services_header:
                service_frame = tk.Frame(self.form_frame, bg=MAIN_BG_COLOR)
                service_frame.grid(row=row, column=entry_col, sticky="we", padx=(0, 10), pady=4)
                service_frame.grid_columnconfigure(0, weight=1)

                self.included_services_combo = ttk.Combobox(
                    service_frame,
                    textvariable=self.included_services_choice_var,
                    values=[""] + self.prestation_options,
                    font=ENTRY_FONT,
                    state="readonly",
                )
                self.included_services_combo.grid(row=0, column=0, sticky="we")
                tk.Button(
                    service_frame,
                    text="Ajouter",
                    command=self._add_included_service,
                    bg=BUTTON_GREEN,
                    fg="white",
                    font=("Arial", 9, "bold"),
                    padx=8,
                    pady=2,
                ).grid(row=0, column=1, padx=(8, 0))
                self.included_services_listbox = tk.Listbox(
                    service_frame,
                    height=4,
                    selectmode=tk.EXTENDED,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                )
                self.included_services_listbox.grid(
                    row=1, column=0, columnspan=2, sticky="we", pady=(6, 4)
                )
                tk.Button(
                    service_frame,
                    text="Retirer sélection",
                    command=self._remove_selected_included_service,
                    bg=BUTTON_ORANGE,
                    fg="white",
                    font=("Arial", 9),
                    padx=8,
                    pady=2,
                ).grid(row=2, column=1, sticky="e")
            else:
                state = (
                    "readonly"
                    if header in (self.default_hotels_header, self.linked_transports_header)
                    else "normal"
                )
                tk.Entry(
                    self.form_frame,
                    textvariable=var,
                    font=ENTRY_FONT,
                    width=35,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                    state=state,
                ).grid(row=row, column=entry_col, sticky="we", padx=(0, 10), pady=4)
            self.vars[header] = var

        if self.cities_header and self.cities_header in self.vars:
            self.vars[self.cities_header].trace_add("write", self._on_itinerary_fields_changed)
        if self.itinerary_header and self.itinerary_header in self.vars:
            self.vars[self.itinerary_header].trace_add("write", self._on_itinerary_fields_changed)

        max_row = (len(self.headers) + 1) // 2

        if self.default_hotels_header:
            workflow_row = max_row + 1
            workflow_frame = tk.LabelFrame(
                self.form_frame,
                text="Villes du circuit -> Hôtel par ville",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
                padx=10,
                pady=8,
            )
            workflow_frame.grid(
                row=workflow_row, column=0, columnspan=4, sticky="we", pady=(8, 0)
            )
            workflow_frame.grid_columnconfigure(1, weight=1)
            workflow_frame.grid_columnconfigure(3, weight=1)

            tk.Label(
                workflow_frame,
                text="Sélectionnez les villes dans l'ordre, puis assignez un hôtel à chaque ville.",
                font=("Arial", 9),
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

            tk.Label(
                workflow_frame,
                text="Ville:",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=1, column=0, sticky="w", padx=(0, 8))
            self.city_step_combo = ttk.Combobox(
                workflow_frame,
                textvariable=self.city_step_var,
                state="readonly",
                width=30,
            )
            self.city_step_combo.grid(row=1, column=1, sticky="we", padx=(0, 10), pady=4)
            self.city_step_combo.bind("<<ComboboxSelected>>", self._on_city_step_selected)

            tk.Label(
                workflow_frame,
                text="Hôtel:",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=1, column=2, sticky="w", padx=(0, 8))
            self.hotel_step_combo = ttk.Combobox(
                workflow_frame,
                textvariable=self.hotel_step_var,
                state="readonly",
                width=30,
            )
            self.hotel_step_combo.grid(row=1, column=3, sticky="we", padx=(0, 0), pady=4)

            workflow_buttons = tk.Frame(workflow_frame, bg=MAIN_BG_COLOR)
            workflow_buttons.grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 8))

            tk.Button(
                workflow_buttons,
                text="Affecter hôtel",
                command=self._assign_hotel_to_city,
                bg=BUTTON_GREEN,
                fg="white",
                font=BUTTON_FONT,
                padx=10,
                pady=4,
            ).pack(side="left", padx=(0, 8))
            tk.Button(
                workflow_buttons,
                text="Retirer sélection",
                command=self._remove_selected_city_hotel_mapping,
                bg=BUTTON_ORANGE,
                fg="white",
                font=BUTTON_FONT,
                padx=10,
                pady=4,
            ).pack(side="left", padx=(0, 8))
            tk.Button(
                workflow_buttons,
                text="Vider",
                command=self._clear_city_hotel_mapping,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=10,
                pady=4,
            ).pack(side="left")

            self.mapping_tree = ttk.Treeview(
                workflow_frame,
                columns=("city", "hotel"),
                show="headings",
                height=5,
            )
            self.mapping_tree.heading("city", text="Ville")
            self.mapping_tree.heading("hotel", text="Hôtel")
            self.mapping_tree.column("city", width=220, anchor="w")
            self.mapping_tree.column("hotel", width=280, anchor="w")
            self.mapping_tree.grid(row=3, column=0, columnspan=4, sticky="we")
            max_row = workflow_row + 1

        if self.linked_transports_header:
            transport_row = max_row + 1
            transport_frame = tk.LabelFrame(
                self.form_frame,
                text="Transports du circuit (segment départ -> arrivée)",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
                padx=10,
                pady=8,
            )
            transport_frame.grid(
                row=transport_row, column=0, columnspan=4, sticky="we", pady=(8, 0)
            )
            transport_frame.grid_columnconfigure(1, weight=1)
            transport_frame.grid_columnconfigure(3, weight=1)

            tk.Label(
                transport_frame,
                text="Choisissez un transport pour chaque segment du circuit.",
                font=("Arial", 9),
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

            tk.Label(
                transport_frame,
                text="Ville départ:",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=1, column=0, sticky="w", padx=(0, 8))
            self.transport_depart_combo = ttk.Combobox(
                transport_frame,
                textvariable=self.transport_depart_var,
                state="readonly",
                width=24,
            )
            self.transport_depart_combo.grid(
                row=1, column=1, sticky="we", padx=(0, 10), pady=4
            )
            self.transport_depart_combo.bind(
                "<<ComboboxSelected>>", self._on_transport_depart_selected
            )

            tk.Label(
                transport_frame,
                text="Ville arrivée:",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=1, column=2, sticky="w", padx=(0, 8))
            self.transport_arrivee_combo = ttk.Combobox(
                transport_frame,
                textvariable=self.transport_arrivee_var,
                state="readonly",
                width=24,
            )
            self.transport_arrivee_combo.grid(
                row=1, column=3, sticky="we", padx=(0, 0), pady=4
            )
            self.transport_arrivee_combo.bind(
                "<<ComboboxSelected>>", self._on_transport_arrivee_selected
            )

            tk.Label(
                transport_frame,
                text="Transport:",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=2, column=0, sticky="w", padx=(0, 8), pady=(4, 0))
            self.transport_choice_combo = ttk.Combobox(
                transport_frame,
                textvariable=self.transport_choice_var,
                values=self.transport_options,
                state="readonly",
                width=55,
            )
            self.transport_choice_combo.grid(
                row=2, column=1, columnspan=3, sticky="we", pady=(4, 0)
            )

            transport_buttons = tk.Frame(transport_frame, bg=MAIN_BG_COLOR)
            transport_buttons.grid(
                row=3, column=0, columnspan=4, sticky="w", pady=(8, 8)
            )
            tk.Button(
                transport_buttons,
                text="Affecter transport",
                command=self._assign_transport_to_segment,
                bg=BUTTON_GREEN,
                fg="white",
                font=BUTTON_FONT,
                padx=10,
                pady=4,
            ).pack(side="left", padx=(0, 8))
            tk.Button(
                transport_buttons,
                text="Retirer sélection",
                command=self._remove_selected_transport_segment,
                bg=BUTTON_ORANGE,
                fg="white",
                font=BUTTON_FONT,
                padx=10,
                pady=4,
            ).pack(side="left", padx=(0, 8))
            tk.Button(
                transport_buttons,
                text="Vider",
                command=self._clear_transport_mapping,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=10,
                pady=4,
            ).pack(side="left")

            self.transport_tree = ttk.Treeview(
                transport_frame,
                columns=("segment", "transport", "metrics"),
                show="headings",
                height=5,
            )
            self.transport_tree.heading("segment", text="Segment")
            self.transport_tree.heading("transport", text="Transport")
            self.transport_tree.heading("metrics", text="KM / Durée")
            self.transport_tree.column("segment", width=240, anchor="w")
            self.transport_tree.column("transport", width=330, anchor="w")
            self.transport_tree.column("metrics", width=140, anchor="w")
            self.transport_tree.grid(row=4, column=0, columnspan=4, sticky="we")
            max_row = transport_row + 1

        actions_row = max_row + 1

        btns = tk.Frame(self.form_frame, bg=MAIN_BG_COLOR)
        btns.grid(row=actions_row, column=0, columnspan=4, sticky="w", pady=(8, 0))

        tk.Button(
            btns,
            text="💾 Enregistrer",
            command=self._save_form,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btns,
            text="❌ Annuler",
            command=self._clear_form,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left", padx=(0, 8))

        for col in range(4):
            self.form_frame.grid_columnconfigure(col, weight=1 if col % 2 == 1 else 0)
        self._sync_included_services_widget_from_var()
        self._refresh_city_hotel_editor()

    def _configure_tree_columns(self):
        self.tree.delete(*self.tree.get_children())
        columns = ["row_number"] + self.headers
        self.tree.configure(columns=columns)

        self.tree.heading("row_number", text="N°")
        self.tree.column("row_number", width=70, minwidth=70)

        for header in self.headers:
            self.tree.heading(header, text=header)
            self.tree.column(header, width=170, minwidth=120)

    def _load_data(self):
        self.headers = get_circuit_db_headers() or []
        self.rows = load_circuit_db_rows() or []
        self._configure_tree_columns()
        self._configure_form()
        self._apply_filters()
        self._update_treeview()
        self._clear_form()

    def _apply_filters(self):
        query = self.search_var.get().strip().lower()
        if not query:
            self.filtered_rows = list(self.rows)
            return

        filtered = []
        for row in self.rows:
            hay = " ".join(str(row.get(h, "")) for h in self.headers).lower()
            if query in hay:
                filtered.append(row)
        self.filtered_rows = filtered

    def _update_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for row in self.filtered_rows:
            values = [row.get("row_number", "")]
            values.extend(row.get(h, "") for h in self.headers)
            self.tree.insert("", "end", iid=str(row.get("row_number")), values=values)

        self.status_label.config(
            text=f"Affichage de {len(self.filtered_rows)} ligne(s) sur {len(self.rows)}"
        )

    def _on_filter_change(self, *_args):
        self._apply_filters()
        self._update_treeview()

    def _on_selection_change(self, _event=None):
        selected = bool(self.tree.selection())
        self.btn_edit.config(state="normal" if selected else "disabled")
        self.btn_delete.config(state="normal" if selected else "disabled")

    def _new_row(self):
        self._clear_form()

    def _edit_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à modifier.")
            return

        try:
            row_number = int(selection[0])
        except (TypeError, ValueError):
            return

        row_data = next((r for r in self.rows if r.get("row_number") == row_number), None)
        if not row_data:
            return

        self.selected_row_number = row_number
        for header in self.headers:
            self.vars[header].set(str(row_data.get(header, "")))
        self._sync_included_services_widget_from_var()
        self._init_city_hotel_mapping_from_form()

    def _delete_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à supprimer.")
            return

        confirm = messagebox.askyesno(
            "Confirmation",
            "Supprimer cette ligne de la base Circuits ?",
        )
        if not confirm:
            return

        try:
            row_number = int(selection[0])
        except (TypeError, ValueError):
            return

        success = delete_circuit_db_row(row_number)
        if success:
            messagebox.showinfo("Succès", "Ligne supprimée avec succès.")
            self._load_data()
        else:
            messagebox.showerror(
                "Erreur",
                "Suppression impossible. Vérifiez que data-hotel.xlsx n'est pas ouvert.",
            )

    def _collect_form_data(self):
        self._refresh_included_services_widget()
        if self.default_hotels_header and self.default_hotels_header in self.vars:
            self.vars[self.default_hotels_header].set(self._format_hotels_mapping())
        if self.linked_transports_header and self.linked_transports_header in self.vars:
            self.vars[self.linked_transports_header].set(self._format_transport_mapping())
        return {header: self.vars[header].get().strip() for header in self.headers}

    def _save_form(self):
        if not self.headers:
            messagebox.showerror("Erreur", "Aucun en-tête détecté pour la feuille Circuits.")
            return

        self._sync_itinerary_from_fields()
        if self.itinerary_cities_order and self.default_hotels_header:
            missing_required = []
            for city in self.itinerary_cities_order:
                key = self._normalize_city(city)
                entry = self.city_hotel_mapping.get(key, {})
                if not entry.get("hotel"):
                    available_hotels = self._hotels_for_city(city)
                    if available_hotels:
                        missing_required.append(city)
            if missing_required:
                messagebox.showwarning(
                    "Validation",
                    "Veuillez choisir un hôtel pour les villes suivantes (hôtels disponibles en base):\n"
                    + "Villes sans hôtel: "
                    + ", ".join(missing_required),
                )
                return

        data = self._collect_form_data()
        if not any(data.values()):
            messagebox.showwarning("Validation", "Veuillez renseigner au moins un champ.")
            return

        if self.selected_row_number is not None:
            result = update_circuit_db_row(self.selected_row_number, data)
            if result == -2:
                messagebox.showerror("Fichier verrouillé", "Fermez data-hotel.xlsx puis réessayez.")
                return
            if result == -1:
                messagebox.showerror("Erreur", "Échec de la mise à jour.")
                return
            messagebox.showinfo("Succès", "Ligne mise à jour avec succès.")
            self._load_data()
            return

        row = save_circuit_db_row(data)
        if row == -2:
            messagebox.showerror("Fichier verrouillé", "Fermez data-hotel.xlsx puis réessayez.")
            return
        if row == -1:
            messagebox.showerror("Erreur", "Échec de l'enregistrement.")
            return

        messagebox.showinfo("Succès", f"Ligne enregistrée à la ligne {row}.")
        self._load_data()

    def _clear_form(self):
        self.selected_row_number = None
        for var in self.vars.values():
            var.set("")
        self.city_hotel_mapping = {}
        self.itinerary_cities_order = []
        self.included_services_selected = []
        self.included_services_choice_var.set("")
        self._refresh_included_services_widget()
        self.transport_segment_pairs = []
        self.transport_segment_mapping = {}
        self.transport_depart_var.set("")
        self.transport_arrivee_var.set("")
        self.transport_choice_var.set("")
        if self.linked_transports_header and self.linked_transports_header in self.vars:
            self.vars[self.linked_transports_header].set("")
        self._refresh_transport_editor()
        self.city_step_var.set("")
        self.hotel_step_var.set("")
        self._refresh_city_hotel_editor()
        self.tree.selection_remove(self.tree.selection())
        self._on_selection_change()
