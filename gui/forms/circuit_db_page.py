"""Circuit DB management page (data-hotel.xlsx / Circuits)."""

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
    load_all_hotels,
    get_circuit_db_headers,
    load_circuit_db_rows,
    save_circuit_db_row,
    update_circuit_db_row,
)


class CircuitDBPage:
    """Circuits DB CRUD page."""

    def __init__(self, parent):
        self.parent = parent
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

        self.hotels = []
        self.city_step_var = tk.StringVar()
        self.hotel_step_var = tk.StringVar()
        self.city_step_combo = None
        self.hotel_step_combo = None
        self.mapping_tree = None
        self.city_hotel_mapping = {}
        self.itinerary_cities_order = []
        self.hotel_display_map = {}

        self._load_hotel_reference()

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

    def _load_hotel_reference(self):
        """Load hotels used for city -> hotel matching."""
        try:
            hotels = load_all_hotels()
        except Exception:
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

    def _refresh_mapping_tree(self):
        if not self.mapping_tree:
            return

        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)

        inserted = set()
        for city in self.itinerary_cities_order:
            key = self._normalize_city(city)
            entry = self.city_hotel_mapping.get(key, {})
            self.mapping_tree.insert(
                "",
                "end",
                iid=key,
                values=(entry.get("city", city), entry.get("hotel", "")),
            )
            inserted.add(key)

        for key, entry in self.city_hotel_mapping.items():
            if key in inserted:
                continue
            self.mapping_tree.insert(
                "",
                "end",
                iid=key,
                values=(entry.get("city", ""), entry.get("hotel", "")),
            )

    def _on_itinerary_fields_changed(self, *_args):
        self._refresh_city_hotel_editor()

    def _on_city_step_selected(self, _event=None):
        self._refresh_city_hotel_editor()

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
            self.city_hotel_mapping.pop(item_id, None)
        if self.default_hotels_header and self.default_hotels_header in self.vars:
            self.vars[self.default_hotels_header].set(self._format_hotels_mapping())
        self._refresh_city_hotel_editor()

    def _clear_city_hotel_mapping(self):
        self.city_hotel_mapping = {}
        if self.default_hotels_header and self.default_hotels_header in self.vars:
            self.vars[self.default_hotels_header].set("")
        self._refresh_city_hotel_editor()

    def _init_city_hotel_mapping_from_form(self):
        self.city_hotel_mapping = {}
        if self.default_hotels_header and self.default_hotels_header in self.vars:
            self.city_hotel_mapping = self._parse_hotels_mapping(
                self.vars[self.default_hotels_header].get().strip()
            )
        self._refresh_city_hotel_editor()

    def _configure_form(self):
        for widget in self.form_frame.winfo_children():
            widget.destroy()

        self.vars = {}
        self.itinerary_header = None
        self.cities_header = None
        self.default_hotels_header = None
        self.city_step_combo = None
        self.hotel_step_combo = None
        self.mapping_tree = None
        self.city_hotel_mapping = {}
        self.itinerary_cities_order = []
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
            state = "readonly" if header == self.default_hotels_header else "normal"
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
        except Exception:
            return

        row_data = next((r for r in self.rows if r.get("row_number") == row_number), None)
        if not row_data:
            return

        self.selected_row_number = row_number
        for header in self.headers:
            self.vars[header].set(str(row_data.get(header, "")))
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
        except Exception:
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
        if self.default_hotels_header and self.default_hotels_header in self.vars:
            self.vars[self.default_hotels_header].set(self._format_hotels_mapping())
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
        self.city_step_var.set("")
        self.hotel_step_var.set("")
        self._refresh_city_hotel_editor()
        self.tree.selection_remove(self.tree.selection())
        self._on_selection_change()
