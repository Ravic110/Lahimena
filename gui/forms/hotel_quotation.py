"""
Hotel quotation GUI component
"""

import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from config import *
from utils.excel_handler import load_all_hotels, load_all_clients, save_hotel_quotation_to_excel
from utils.logger import logger
from utils.pdf_generator import generate_hotel_quotation_pdf, REPORTLAB_AVAILABLE
import os
import datetime
import subprocess
from utils.validators import get_exchange_rates, convert_currency


ROOM_GROUP_LABELS = {
    "standard": "Standard",
    "bungalows": "Bungalows",
    "deluxe": "De luxe",
    "suite": "Suite"
}
ROOM_GROUP_KEYS = {label: key for key, label in ROOM_GROUP_LABELS.items()}

ROOM_TYPE_LABELS = {
    "single": "Single",
    "double": "Double",
    "twin": "Twin",
    "familiale": "Familiale",
    "triple": "Triple",
    "chauffeur": "Chauffeur",
    "dortoir": "Dortoir",
    "supp": "Suppl.",
    "studios": "Studio",
    "vip": "VIP"
}
ROOM_TYPE_KEYS = {label: key for key, label in ROOM_TYPE_LABELS.items()}


class HotelQuotation:
    """
    Hotel quotation component for creating hotel price quotes
    """

    def __init__(self, parent):
        """
        Initialize hotel quotation

        Args:
            parent: Parent widget
        """
        self.parent = parent
        self.hotels = self._load_and_filter_hotels()
        self.clients = self._load_clients()
        self.selected_hotel = None

        self._create_quotation_form()

        # Bind client type change to update hotel list
        if hasattr(self, 'client_type_var'):
            self.client_type_var.trace('w', self._on_client_type_changed)

    def _load_and_filter_hotels(self, client_type=None):
        """Load hotels and filter duplicates"""
        hotels = load_all_hotels(client_type)
        
        # Remove duplicates based on nom, lieu, and categorie
        unique_hotels = {}
        for hotel in hotels:
            categorie = (hotel.get('categorie') or '').strip()
            key = f"{hotel['nom']}_{hotel['lieu']}_{categorie}"
            if key not in unique_hotels:
                unique_hotels[key] = hotel
        
        return list(unique_hotels.values())

    def _hotel_display(self, hotel):
        """Build display label for hotel choice"""
        nom = (hotel.get('nom') or '').strip()
        lieu = (hotel.get('lieu') or '').strip()
        categorie = (hotel.get('categorie') or '').strip()
        if categorie:
            return f"{nom} - {lieu} ({categorie})"
        return f"{nom} - {lieu}"

    def _load_clients(self):
        """Load all clients from Excel"""
        return load_all_clients()

    def _create_quotation_form(self):
        """Create the hotel quotation interface"""
        # Clear parent
        for widget in self.parent.winfo_children():
            widget.destroy()

        # Title
        title = tk.Label(
            self.parent,
            text="COTATION HÔTEL",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        )
        title.pack(pady=(20, 10))

        # Main frame
        main_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Client information section (moved to top)
        client_frame = tk.LabelFrame(
            main_frame,
            text="Informations client",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10
        )
        client_frame.pack(fill="x", pady=(0, 10))

        # Row 0: Client selection
        tk.Label(
            client_frame,
            text="Sélectionner client:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=0, column=0, sticky="w", pady=5)

        self.client_var = tk.StringVar()
        client_names = [""]
        for client in self.clients:
            nom = (client.get('nom') or '').strip()
            prenom = (client.get('prenom') or '').strip()
            full_name = f"{nom} {prenom}".strip()
            client_names.append(full_name if full_name else (client.get('ref_client') or ''))
        self.client_combo = ttk.Combobox(
            client_frame,
            textvariable=self.client_var,
            values=client_names,
            font=ENTRY_FONT,
            width=30,
            state="readonly"
        )
        self.client_combo.grid(row=0, column=1, columnspan=3, padx=(10, 0), pady=5, sticky="w")
        self.client_combo.bind("<<ComboboxSelected>>", self._on_client_selected)

        # Row 1: Name and surname
        tk.Label(
            client_frame,
            text="Nom et Prénom:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=1, column=0, sticky="w", pady=5)

        self.client_name_var = tk.StringVar()
        # Make the name field wider and remove the separate 'Prénom' field
        self.client_name_entry = tk.Entry(
            client_frame,
            textvariable=self.client_name_var,
            font=ENTRY_FONT,
            width=40,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.client_name_entry.grid(row=1, column=1, columnspan=3, padx=(10, 0), pady=5, sticky="w")

        # Row 2: Email and phone
        tk.Label(
            client_frame,
            text="Email:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=2, column=0, sticky="w", pady=5)

        self.client_email_var = tk.StringVar()
        self.client_email_entry = tk.Entry(
            client_frame,
            textvariable=self.client_email_var,
            font=ENTRY_FONT,
            width=20,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.client_email_entry.grid(row=2, column=1, padx=(10, 20), pady=5)

        tk.Label(
            client_frame,
            text="Téléphone:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=2, column=2, sticky="w", pady=5)

        self.client_phone_var = tk.StringVar()
        self.client_phone_entry = tk.Entry(
            client_frame,
            textvariable=self.client_phone_var,
            font=ENTRY_FONT,
            width=20,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.client_phone_entry.grid(row=2, column=3, padx=(10, 0), pady=5)

        # Hotel selection section
        hotel_frame = tk.LabelFrame(
            main_frame,
            text="Sélection de l'hôtel",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10
        )
        hotel_frame.pack(fill="x", pady=(0, 10))

        # City selection + Hotel selection (city first)
        tk.Label(
            hotel_frame,
            text="Ville:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=0, column=0, sticky="w", pady=5)

        # Build unique city list from hotels
        cities = sorted(list({hotel.get('lieu', '') for hotel in self.hotels if hotel.get('lieu')}))
        self.city_var = tk.StringVar()
        self.city_combo = ttk.Combobox(
            hotel_frame,
            textvariable=self.city_var,
            values=[""] + cities,
            font=ENTRY_FONT,
            width=20,
            state="readonly"
        )
        self.city_combo.grid(row=0, column=1, padx=(10, 10), pady=5, sticky="w")
        self.city_combo.bind("<<ComboboxSelected>>", self._on_city_selected)

        tk.Label(
            hotel_frame,
            text="Hôtel:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=0, column=2, sticky="w", pady=5)

        self.hotel_var = tk.StringVar()
        # Initially empty; will be populated when a city is selected
        self.hotel_combo = ttk.Combobox(
            hotel_frame,
            textvariable=self.hotel_var,
            values=[],
            font=ENTRY_FONT,
            width=40,
            state="readonly"
        )
        self.hotel_combo.grid(row=0, column=3, padx=(10, 0), pady=5)
        self.hotel_combo.bind("<<ComboboxSelected>>", self._on_hotel_selected)

        # Parameters section
        params_frame = tk.LabelFrame(
            main_frame,
            text="Paramètres du séjour",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10
        )
        params_frame.pack(fill="x", pady=(0, 10))

        # Row 0: Number of nights and room group
        tk.Label(
            params_frame,
            text="Nombre de nuits:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=0, column=0, sticky="w", pady=5)

        self.nights_var = tk.StringVar(value="1")
        self.nights_entry = tk.Entry(
            params_frame,
            textvariable=self.nights_var,
            font=ENTRY_FONT,
            width=10,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.nights_entry.grid(row=0, column=1, padx=(10, 20), pady=5)

        tk.Label(
            params_frame,
            text="Gamme:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=0, column=2, sticky="w", pady=5)

        self.room_group_var = tk.StringVar(value="")
        self.room_group_combo = ttk.Combobox(
            params_frame,
            textvariable=self.room_group_var,
            values=[],
            font=ENTRY_FONT,
            width=15,
            state="readonly"
        )
        self.room_group_combo.grid(row=0, column=3, padx=(10, 0), pady=5)
        self.room_group_combo.bind("<<ComboboxSelected>>", self._on_room_group_selected)

        # Row 1: Room type and adults
        tk.Label(
            params_frame,
            text="Type de chambre:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=1, column=0, sticky="w", pady=5)

        self.room_type_var = tk.StringVar(value="")
        self.room_type_combo = ttk.Combobox(
            params_frame,
            textvariable=self.room_type_var,
            values=[],
            font=ENTRY_FONT,
            width=15,
            state="readonly"
        )
        self.room_type_combo.grid(row=1, column=1, padx=(10, 20), pady=5)

        tk.Label(
            params_frame,
            text="Adultes:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=1, column=2, sticky="w", pady=5)

        self.adults_var = tk.StringVar(value="2")
        self.adults_entry = tk.Entry(
            params_frame,
            textvariable=self.adults_var,
            font=ENTRY_FONT,
            width=10,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.adults_entry.grid(row=1, column=3, padx=(10, 0), pady=5)

        # Row 2: Number of children and client type
        tk.Label(
            params_frame,
            text="Enfants:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=2, column=0, sticky="w", pady=5)

        self.children_var = tk.StringVar(value="0")
        self.children_entry = tk.Entry(
            params_frame,
            textvariable=self.children_var,
            font=ENTRY_FONT,
            width=10,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.children_entry.grid(row=2, column=1, padx=(10, 20), pady=5)

        tk.Label(
            params_frame,
            text="Type de client:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=2, column=2, sticky="w", pady=5)

        self.client_type_var = tk.StringVar(value="PBC")
        self.client_type_combo = ttk.Combobox(
            params_frame,
            textvariable=self.client_type_var,
            values=["TO", "PBC"],
            font=ENTRY_FONT,
            width=10,
            state="readonly"
        )
        self.client_type_combo.grid(row=2, column=3, padx=(10, 0), pady=5)

        # Row 3: Period and meal plan
        tk.Label(
            params_frame,
            text="Période:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=3, column=0, sticky="w", pady=5)

        self.period_var = tk.StringVar(value="Moyenne saison")
        self.period_combo = ttk.Combobox(
            params_frame,
            textvariable=self.period_var,
            values=PERIODES,
            font=ENTRY_FONT,
            width=15,
            state="readonly"
        )
        self.period_combo.grid(row=3, column=1, padx=(10, 20), pady=5)

        tk.Label(
            params_frame,
            text="Restauration:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=3, column=2, sticky="w", pady=5)

        self.meal_var = tk.StringVar(value="Petit déjeuner")
        self.meal_combo = ttk.Combobox(
            params_frame,
            textvariable=self.meal_var,
            values=RESTAURATIONS,
            font=ENTRY_FONT,
            width=20,
            state="readonly"
        )
        self.meal_combo.grid(row=3, column=3, padx=(10, 0), pady=5)

        # Row 4: Currency
        tk.Label(
            params_frame,
            text="Devise:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=4, column=0, sticky="w", pady=5)

        self.currency_var = tk.StringVar(value="Ariary")
        self.currency_combo = ttk.Combobox(
            params_frame,
            textvariable=self.currency_var,
            values=["Ariary", "Euro", "Dollar US"],
            font=ENTRY_FONT,
            width=10,
            state="readonly"
        )
        self.currency_combo.grid(row=4, column=1, padx=(10, 0), pady=5)

        # Exchange rates display
        rates_frame = tk.Frame(params_frame, bg=MAIN_BG_COLOR)
        rates_frame.grid(row=4, column=2, columnspan=2, padx=(20, 0), pady=5, sticky="w")
        
        self.rates_label = tk.Label(
            rates_frame,
            text="Taux de change :\nChargement...",
            font=("Arial", 10, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            justify="left"
        )
        self.rates_label.pack(side="left")
        
        # Refresh rates button
        self.refresh_rates_button = tk.Button(
            rates_frame,
            text="🔄",
            command=self._update_exchange_rates,
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
            font=("Arial", 10),
            padx=5,
            pady=2
        )
        self.refresh_rates_button.pack(side="left", padx=(10, 0))

        # Load initial exchange rates
        self._update_exchange_rates()

        # Bind currency change to update rates display
        self.currency_var.trace('w', self._on_currency_changed)

        # Calculate button
        calc_frame = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
        calc_frame.pack(fill="x", pady=(10, 10))

        self.calc_button = tk.Button(
            calc_frame,
            text="🧮 Calculer le prix",
            command=self._calculate_price,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=20,
            pady=10
        )
        self.calc_button.pack()

        # Results section
        self.results_frame = tk.LabelFrame(
            main_frame,
            text="Résultats de la cotation",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10
        )
        self.results_frame.pack(fill="both", expand=True, pady=(0, 10))

        # Results will be displayed here
        self.results_text = tk.Text(
            self.results_frame,
            font=("Courier", 10),
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
            height=15,
            wrap=tk.WORD
        )
        self.results_text.pack(fill="both", expand=True)

        # Action buttons
        action_frame = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
        action_frame.pack(fill="x", pady=(0, 10))

        tk.Button(
            action_frame,
            text="📄 Générer devis",
            command=self._generate_quote,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8
        ).pack(side="left", padx=5)

        tk.Button(
            action_frame,
            text="🔄 Nouvelle cotation",
            command=self._reset_form,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8
        ).pack(side="left", padx=5)

        # Quotation history section
        history_frame = tk.LabelFrame(
            main_frame,
            text="📋 Liste des cotations générées",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10
        )
        history_frame.pack(fill="both", expand=True, pady=(0, 10))

        # Treeview for quotations
        columns = ("filename", "date", "size")
        self.quotations_tree = ttk.Treeview(
            history_frame,
            columns=columns,
            height=8,
            show="tree headings"
        )
        self.quotations_tree.heading("#0", text="Nom du fichier")
        self.quotations_tree.heading("date", text="Date")
        self.quotations_tree.heading("size", text="Taille")
        
        self.quotations_tree.column("#0", width=250)
        self.quotations_tree.column("date", width=120)
        self.quotations_tree.column("size", width=80)
        
        self.quotations_tree.pack(fill="both", expand=True, side="left")
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(history_frame, orient="vertical", command=self.quotations_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.quotations_tree.configure(yscroll=scrollbar.set)
        
        # Buttons for quotation management
        history_buttons = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
        history_buttons.pack(fill="x", pady=(0, 10))
        
        tk.Button(
            history_buttons,
            text="🔄 Rafraîchir",
            command=self._refresh_quotations_list,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8
        ).pack(side="left", padx=5)
        
        tk.Button(
            history_buttons,
            text="📂 Ouvrir devis",
            command=self._open_selected_quotation,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8
        ).pack(side="left", padx=5)
        
        tk.Button(
            history_buttons,
            text="🗑️ Supprimer devis",
            command=self._delete_selected_quotation,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8
        ).pack(side="left", padx=5)
        
        # Load initial quotation list
        self._refresh_quotations_list()

    def _on_hotel_selected(self, event=None):
        """Handle hotel selection"""
        selection = self.hotel_var.get()
        if selection:
            # Find the selected hotel
            for hotel in self.hotels:
                hotel_display = self._hotel_display(hotel)
                if hotel_display == selection:
                    self.selected_hotel = hotel
                    self._update_room_group_options()
                    break

    def _on_city_selected(self, event=None):
        """Filter hotels when a city is selected"""
        city = self.city_var.get()
        # Filter hotels by city and update hotel combobox
        if city:
            filtered = [self._hotel_display(h) for h in self.hotels if h.get('lieu') == city]
        else:
            filtered = [self._hotel_display(h) for h in self.hotels]
        # Update combobox values and clear previous selection
        self.hotel_combo['values'] = filtered
        self.hotel_var.set("")
        self.selected_hotel = None
        self._clear_room_selections()

    def _clear_room_selections(self):
        """Reset room group/type selections"""
        if hasattr(self, 'room_group_combo'):
            self.room_group_combo['values'] = []
        if hasattr(self, 'room_type_combo'):
            self.room_type_combo['values'] = []
        if hasattr(self, 'room_group_var'):
            self.room_group_var.set("")
        if hasattr(self, 'room_type_var'):
            self.room_type_var.set("")

    def _get_room_group_options(self):
        """Get available room groups for the selected hotel"""
        if not self.selected_hotel:
            return []
        rates = self.selected_hotel.get('room_rates', {})
        options = []
        for key, label in ROOM_GROUP_LABELS.items():
            group_rates = rates.get(key, {})
            if any(v for v in group_rates.values() if v):
                options.append(label)
        return options

    def _get_room_type_options(self, group_key):
        """Get available room types for a room group"""
        if not self.selected_hotel or not group_key:
            return []
        rates = self.selected_hotel.get('room_rates', {}).get(group_key, {})
        options = []
        for key, label in ROOM_TYPE_LABELS.items():
            if rates.get(key):
                options.append(label)
        if not options and rates:
            for key in rates.keys():
                label = ROOM_TYPE_LABELS.get(key)
                if label:
                    options.append(label)
        return options

    def _update_room_group_options(self):
        """Update room group and room type options based on selected hotel"""
        group_options = self._get_room_group_options()
        self.room_group_combo['values'] = group_options
        if group_options:
            self.room_group_var.set(group_options[0])
        else:
            self.room_group_var.set("")
        self._on_room_group_selected()

    def _on_room_group_selected(self, event=None):
        """Handle room group selection"""
        group_label = self.room_group_var.get()
        group_key = ROOM_GROUP_KEYS.get(group_label, "")
        room_options = self._get_room_type_options(group_key)
        self.room_type_combo['values'] = room_options
        if room_options:
            self.room_type_var.set(room_options[0])
        else:
            self.room_type_var.set("")

    def _get_room_display(self):
        """Build a display label for the selected room group/type"""
        group_label = self.room_group_var.get().strip() if hasattr(self, 'room_group_var') else ""
        room_label = self.room_type_var.get().strip() if hasattr(self, 'room_type_var') else ""
        if group_label and room_label:
            return f"{group_label} / {room_label}"
        return room_label or group_label or ""

    def _on_client_selected(self, event=None):
        """Handle client selection and auto-fill fields"""
        selection = self.client_var.get()
        if selection:
            # Find the selected client
            for client in self.clients:
                nom = (client.get('nom') or '').strip()
                prenom = (client.get('prenom') or '').strip()
                client_display = f"{nom} {prenom}".strip() or (client.get('ref_client') or '')
                if client_display == selection:
                    # Auto-fill the fields
                    full_name = f"{nom} {prenom}".strip()
                    self.client_name_var.set(full_name)
                    self.client_email_var.set(client['email'])
                    self.client_phone_var.set(client['telephone'])
                    self.client_var.set(client.get('ref_client', ''))
                    # Auto-fill other quotation parameters from client data when available
                    # Period
                    period = client.get('periode')
                    if period:
                        try:
                            self.period_var.set(period)
                        except Exception:
                            pass

                    # Meal plan / restauration
                    resta = client.get('restauration')
                    if resta and resta in RESTAURATIONS:
                        try:
                            self.meal_var.set(resta)
                        except Exception:
                            pass

                    # Room type: match against available options or keywords
                    ch = client.get('chambre')
                    if ch:
                        try:
                            ch_str = str(ch)
                            room_options = list(self.room_type_combo['values']) if hasattr(self, 'room_type_combo') else []
                            if not room_options:
                                room_options = list(ROOM_TYPE_LABELS.values())
                            # direct match
                            if ch_str in room_options:
                                self.room_type_var.set(ch_str)
                            else:
                                # try case-insensitive keyword match
                                lower = ch_str.lower()
                                for t in room_options:
                                    if t.lower() in lower:
                                        self.room_type_var.set(t)
                                        break
                        except Exception:
                            pass

                    # Children count
                    enfants = client.get('enfant')
                    if enfants:
                        try:
                            # try to convert to int
                            self.children_var.set(str(int(enfants)))
                        except (ValueError, TypeError):
                            # if not numeric, check for common non-values
                            enfants_str = str(enfants).lower()
                            if enfants_str in ('non', 'no', 'n', '', '0'):
                                self.children_var.set("0")
                            else:
                                # default to 0 for unknown values
                                self.children_var.set("0")
                    else:
                        # Keep default value (0) if enfant field is empty
                        self.children_var.set("0")
                    
                    # Ensure numeric fields are always initialized
                    if not self.nights_var.get():
                        self.nights_var.set("1")
                    if not self.adults_var.get():
                        self.adults_var.set("2")
                    if not self.children_var.get():
                        self.children_var.set("0")
                    break
        else:
            # Clear fields if no client selected
            self.client_name_var.set("")
            self.client_email_var.set("")
            self.client_phone_var.set("")
            # Keep default values for numeric fields
            self.nights_var.set("1")
            self.adults_var.set("2")
            self.children_var.set("0")

    def _on_client_type_changed(self, *args):
        """Handle client type change to update hotel list"""
        if not hasattr(self, 'client_type_var'):
            return
        client_type = self.client_type_var.get()
        self.hotels = self._load_and_filter_hotels(client_type)
        
        # Update hotel combobox
        hotel_names = [self._hotel_display(hotel) for hotel in self.hotels]
        self.hotel_combo['values'] = hotel_names
        
        # Reset hotel selection
        self.hotel_var.set("")
        self.selected_hotel = None
        self._clear_room_selections()

    def _update_exchange_rates(self):
        """Update the exchange rates display"""
        try:
            from utils.validators import get_exchange_rates
            rates = get_exchange_rates()
            
            selected_currency = self.currency_var.get()
            
            eur_text = f"1 EUR = {rates['EUR']:.4f} MGA"
            usd_text = f"1 USD = {rates['USD']:.4f} MGA"
            
            # Highlight selected currency
            if selected_currency == "Euro":
                eur_text = f"▶ {eur_text} ◀"
            elif selected_currency == "Dollar US":
                usd_text = f"▶ {usd_text} ◀"
            
            rates_text = f"Taux de change :\n{eur_text}\n{usd_text}"
            self.rates_label.config(text=rates_text, fg=TEXT_COLOR)
        except Exception as e:
            self.rates_label.config(text=f"Taux de change :\nErreur: {str(e)}", fg="red")

    def _on_currency_changed(self, *args):
        """Handle currency change to update rates display"""
        # Update rates display with highlighting
        self._update_exchange_rates()

    def _calculate_price(self):
        """Calculate the total price based on parameters"""
        if not self.selected_hotel:
            messagebox.showwarning("Hôtel non sélectionné", "Veuillez d'abord sélectionner un hôtel.")
            return

        try:
            # Get and validate parameters with safe fallbacks
            nights_str = self.nights_var.get().strip() if self.nights_var.get() else "1"
            adults_str = self.adults_var.get().strip() if self.adults_var.get() else "2"
            children_str = self.children_var.get().strip() if self.children_var.get() else "0"
            
            # Restore defaults if empty
            if not nights_str:
                nights_str = "1"
                self.nights_var.set("1")
            if not adults_str:
                adults_str = "2"
                self.adults_var.set("2")
            if not children_str:
                children_str = "0"
                self.children_var.set("0")
            
            # Validate and convert to integers
            try:
                nights = int(nights_str)
                adults = int(adults_str)
                children = int(children_str)
            except ValueError:
                messagebox.showerror(
                    "❌ Erreur de saisie",
                    "Veuillez saisir des valeurs numériques valides pour:\n"
                    f"- Nuits: '{nights_str}'\n"
                    f"- Adultes: '{adults_str}'\n"
                    f"- Enfants: '{children_str}'\n\n"
                    "Exemple: 3, 2, 1"
                )
                logger.warning(f"Invalid numeric input - Nuits: {nights_str}, Adultes: {adults_str}, Enfants: {children_str}")
                return
            
            if nights <= 0 or adults <= 0:
                messagebox.showerror("Erreur", "Le nombre de nuits et d'adultes doit être supérieur à 0.")
                return
                
            room_group_label = self.room_group_var.get().strip() if self.room_group_var.get() else ""
            room_type_label = self.room_type_var.get().strip() if self.room_type_var.get() else ""
            room_group_key = ROOM_GROUP_KEYS.get(room_group_label, "standard")
            room_type_key = ROOM_TYPE_KEYS.get(room_type_label, "")

            # Get prices from hotel data
            room_price = 0
            if room_type_key:
                group_rates = self.selected_hotel.get('room_rates', {}).get(room_group_key, {})
                room_price = group_rates.get(room_type_key, 0)

            if room_price == 0:
                # Fallback to legacy fields if room_rates are not populated
                if room_type_label == "Single" and self.selected_hotel.get('chambre_single'):
                    room_price = self.selected_hotel['chambre_single']
                elif room_type_label == "Double" and self.selected_hotel.get('chambre_double'):
                    room_price = self.selected_hotel['chambre_double']
                elif room_type_label == "Triple" and self.selected_hotel.get('chambre_double'):
                    room_price = self.selected_hotel['chambre_double']
                elif room_type_label == "Familliale" and self.selected_hotel.get('chambre_familiale'):
                    room_price = self.selected_hotel['chambre_familiale']

            if room_price == 0:
                messagebox.showwarning("Prix non disponible", f"Le prix pour {room_type_label} n'est pas disponible pour cet hôtel.")
                return

            # Calculate base price
            base_price = room_price * nights

            # Calculate meal supplements
            meals = self.selected_hotel.get('meals', {})
            petit_dej = meals.get('petit_dejeuner', self.selected_hotel.get('petit_dejeuner', 0))
            dejeuner = meals.get('dejeuner', self.selected_hotel.get('dejeuner', 0))
            diner = meals.get('diner', self.selected_hotel.get('diner', 0))
            meal_price = 0
            meal_plan = self.meal_var.get()
            if meal_plan == "Petit déjeuner" and petit_dej:
                meal_price = petit_dej * nights * (adults + children)
            elif meal_plan == "Demi-pension" and dejeuner and diner:
                meal_price = (dejeuner + diner) * nights * (adults + children)
            elif meal_plan == "Pension complète" and petit_dej and dejeuner and diner:
                meal_price = (petit_dej + dejeuner + diner) * nights * (adults + children)

            # Calculate total
            total_price = base_price + meal_price

            # Apply client type adjustment
            client_type = self.client_type_var.get()
            if client_type == "PBC":
                # Public pricing with markup
                total_price *= 1.2

            # Apply currency conversion
            currency = self.currency_var.get()
            if currency != "Ariary":
                rates = get_exchange_rates()
                base_price = convert_currency(base_price, "Ariary", currency, rates)
                meal_price = convert_currency(meal_price, "Ariary", currency, rates)
                total_price = convert_currency(total_price, "Ariary", currency, rates)

            # Display results
            client_name = self.client_name_var.get()
            client_email = self.client_email_var.get()
            client_phone = self.client_phone_var.get()
            display_room = room_type_label
            if room_group_label:
                display_room = f"{room_group_label} / {room_type_label}" if room_type_label else room_group_label
            self._display_results(base_price, meal_price, total_price, nights, adults, children, display_room, meal_plan, client_type, currency, client_name, client_email, client_phone)

        except ValueError as e:
            error_details = str(e)
            messagebox.showerror(
                "❌ Erreur",
                f"Une erreur est survenue:\n\n{error_details}\n\nVeuillez vérifier vos valeurs."
            )
            logger.error(f"ValueError in _calculate_price: {error_details}", exc_info=True)
        except Exception as e:
            error_details = str(e)
            messagebox.showerror(
                "❌ Erreur inattendue",
                f"Une erreur inattendue s'est produite:\n\n{error_details}\n\nVérifiez les logs."
            )
            logger.error(f"Unexpected error in _calculate_price: {error_details}", exc_info=True)

    def _display_results(self, base_price, meal_price, total_price, nights, adults, children, room_type, meal_plan, client_type, currency, client_name, client_email, client_phone):
        """Display calculation results"""
        # Get currency symbol
        currency_symbols = {
            "Ariary": "Ar",
            "Euro": "€",
            "Dollar US": "$"
        }
        symbol = currency_symbols.get(currency, "Ar")

        # Format prices (format specifier strings should NOT include the leading ':')
        if currency == "Ariary":
            price_format = ",.0f"
        else:
            price_format = ",.2f"

        self.results_text.delete(1.0, tk.END)

        # Prepare a prominent client name banner (remove first name display)
        display_name = (client_name or "").upper()
        name_banner = display_name.center(50)

        result_text = f"""
    COTATION HÔTEL - {self.selected_hotel['nom']}
    {'='*50}

    HÔTEL: {self.selected_hotel['nom']}
    LIEU: {self.selected_hotel['lieu']}
    CATÉGORIE: {self.selected_hotel['categorie']}

    PARAMÈTRES:
    - Nombre de nuits: {nights}
    - Type de chambre: {room_type}
    - Adultes: {adults}
    - Enfants: {children}
    - Restauration: {meal_plan}
    - Période: {self.period_var.get()}
    - Type de client: {client_type}
    - Devise: {currency}

    INFORMATIONS CLIENT:
    {name_banner}
    - Nom client: {display_name}
    - Email: {client_email}
    - Téléphone: {client_phone}

    CALCUL:
    - Prix chambre/nuit: {base_price/nights:{price_format}} {symbol}
    - Prix total chambre: {base_price:{price_format}} {symbol}
    - Supplément restauration: {meal_price:{price_format}} {symbol}

    TOTAL: {total_price:{price_format}} {symbol}

    Prix par personne: {total_price/(adults+children):{price_format}} {symbol}
    Prix par nuit: {total_price/nights:{price_format}} {symbol}
    """

        self.results_text.insert(tk.END, result_text.strip())

    def _generate_quote(self):
        """Generate a formal quote"""
        if not self.selected_hotel:
            messagebox.showwarning("Hôtel non sélectionné", "Veuillez d'abord sélectionner un hôtel et calculer le prix.")
            return

        if not self.results_text.get(1.0, tk.END).strip():
            messagebox.showwarning("Calcul manquant", "Veuillez d'abord calculer le prix.")
            return

        try:
            # Check if ReportLab is available for PDF generation
            if not REPORTLAB_AVAILABLE:
                messagebox.showwarning(
                    "⚠️ Génération PDF non disponible",
                    "ReportLab n'est pas installé. Veuillez installer le package:\n\npip install reportlab\n\nFallback: Génération de fichier texte..."
                )
                logger.warning("ReportLab not available, falling back to text generation")

            # Extract quote parameters
            currency = self.currency_var.get()
            currency_map = {"Ariary": "MGA", "Euro": "EUR", "Dollar US": "USD"}
            currency_code = currency_map.get(currency, "MGA")

            # Get client info if available
            client_name = self.client_var.get() if hasattr(self, 'client_var') else "Client"
            client_email = ""
            
            # Try to get client email if client selected
            if client_name and client_name != "Sélectionner un client":
                try:
                    clients = load_all_clients()
                    for client in clients:
                        if client['client_name'] == client_name:
                            client_email = client.get('email', '')
                            break
                except Exception as e:
                    logger.warning(f"Could not retrieve client email: {e}")

            # Extract pricing details from results
            room_type = self._get_room_display() or "Standard"
            try:
                nights = int(self.nights_var.get()) if hasattr(self, 'nights_var') else 1
            except (ValueError, TypeError):
                nights = 1
            try:
                adults = int(self.adults_var.get()) if hasattr(self, 'adults_var') else 1
            except (ValueError, TypeError):
                adults = 1
            
            # Extract price from results or calculate
            try:
                results_content = self.results_text.get(1.0, tk.END).strip()
                # Extract total price from results text
                price_per_night = float(self.total_price_per_night) if hasattr(self, 'total_price_per_night') else 0
                total_price = price_per_night * nights
            except:
                price_per_night = 0
                total_price = 0

            # Generate PDF quotation
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            quote_number = f"DEVIS_HOTEL_{timestamp}"
            devis_folder = DEVIS_FOLDER
            
            # Ensure devis folder exists
            if not os.path.exists(devis_folder):
                os.makedirs(devis_folder)

            filename = generate_hotel_quotation_pdf(
                hotel_name=self.selected_hotel['nom'],
                client_name=client_name,
                client_email=client_email,
                quote_number=quote_number,
                quote_date=datetime.datetime.now().strftime("%d/%m/%Y"),
                nights=nights,
                adults=adults,
                room_type=room_type,
                price_per_night=price_per_night,
                total_price=total_price,
                currency=currency_code,
                hotel_location=self.selected_hotel.get('lieu', ''),
                hotel_category=self.selected_hotel.get('categorie', ''),
                hotel_contact=self.selected_hotel.get('contact', ''),
                hotel_email=self.selected_hotel.get('email', ''),
                output_dir=devis_folder
            )

            # Save quotation to COTATION_H sheet
            try:
                quotation_data = {
                    'client_id': self.client_var.get().split(' - ')[0] if ' - ' in self.client_var.get() else '',
                    'client_name': client_name,
                    'hotel_name': self.selected_hotel['nom'],
                    'city': self.selected_hotel.get('lieu', ''),
                    'total_price': total_price,
                    'currency': currency,
                    'nights': nights,
                    'adults': adults,
                    'children': int(self.children_var.get()) if hasattr(self, 'children_var') else 0,
                    'room_type': room_type,
                    'meal_plan': self.meal_var.get() if hasattr(self, 'meal_var') else '',
                    'period': self.period_var.get() if hasattr(self, 'period_var') else '',
                    'quote_date': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                save_hotel_quotation_to_excel(quotation_data)
                logger.info(f"Quotation saved to COTATION_H sheet: {client_name} - {self.selected_hotel['nom']}")
            except Exception as e:
                logger.warning(f"Could not save quotation to Excel: {e}")

            # Show success message with file location
            messagebox.showinfo(
                "✅ Devis généré avec succès",
                f"Le devis PDF a été sauvegardé :\n{filename}\n\nLe fichier va s'ouvrir automatiquement."
            )
            logger.info(f"PDF quotation generated successfully: {filename}")

            # Refresh quotation history list
            self._refresh_quotations_list()

            # Try to open the file
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(filename)
                elif os.name == 'posix':  # macOS/Linux
                    subprocess.run(['xdg-open', filename])
            except Exception as e:
                logger.warning(f"Could not open quotation file automatically: {e}")
                messagebox.showwarning(
                    "⚠️ Ouverture automatique impossible",
                    f"Le fichier a été créé mais n'a pas pu s'ouvrir automatiquement.\n\nVous pouvez l'ouvrir manuellement : {filename}"
                )

        except Exception as e:
            error_msg = f"Une erreur s'est produite lors de la génération du devis :\n{str(e)}\n\nDétails dans les logs de l'application."
            messagebox.showerror(
                "❌ Erreur lors de la génération",
                error_msg
            )
            logger.error(f"Error generating quotation: {e}", exc_info=True)

    def _reset_form(self):
        """Reset the form for a new quotation"""
        self.hotel_var.set("")
        self.selected_hotel = None
        self.nights_var.set("1")
        self.adults_var.set("2")
        self.children_var.set("0")
        self._clear_room_selections()
        self.client_type_var.set("PBC")
        self.currency_var.set("Ariary")
        self.period_var.set("Moyenne saison")
        self.meal_var.set("Petit déjeuner")
        self.client_name_var.set("")
        self.client_email_var.set("")
        self.client_phone_var.set("")
        self.results_text.delete(1.0, tk.END)

    def _refresh_quotations_list(self):
        """Load and display all quotation files from devis folder"""
        # Clear existing items
        for item in self.quotations_tree.get_children():
            self.quotations_tree.delete(item)
        
        try:
            devis_folder = DEVIS_FOLDER
            if not os.path.exists(devis_folder):
                logger.warning(f"Devis folder not found: {devis_folder}")
                return
            
            # Get all PDF files
            pdf_files = [f for f in os.listdir(devis_folder) if f.endswith('.pdf')]
            pdf_files.sort(reverse=True)  # Most recent first
            
            # Add files to treeview
            for filename in pdf_files:
                filepath = os.path.join(devis_folder, filename)
                try:
                    # Get file info
                    file_stat = os.stat(filepath)
                    file_size = f"{file_stat.st_size / 1024:.1f} KB"
                    file_date = datetime.datetime.fromtimestamp(file_stat.st_mtime).strftime("%d/%m/%Y %H:%M")
                    
                    # Add to treeview
                    self.quotations_tree.insert("", "end", text=filename, values=(filename, file_date, file_size))
                except Exception as e:
                    logger.error(f"Error processing quotation file {filename}: {e}")
        
        except Exception as e:
            logger.error(f"Error refreshing quotations list: {e}")
            messagebox.showerror("Erreur", f"Erreur lors du chargement de la liste des devis:\n{str(e)}")

    def _open_selected_quotation(self):
        """Open the selected quotation file"""
        selection = self.quotations_tree.selection()
        if not selection:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un devis à ouvrir.")
            return
        
        try:
            selected_item = selection[0]
            filename = self.quotations_tree.item(selected_item)['text']
            filepath = os.path.join(DEVIS_FOLDER, filename)
            
            if not os.path.exists(filepath):
                messagebox.showerror("Erreur", f"Le fichier n'existe pas:\n{filepath}")
                return
            
            # Open the file
            if os.name == 'nt':  # Windows
                os.startfile(filepath)
            elif os.name == 'posix':  # macOS/Linux
                subprocess.run(['xdg-open', filepath])
            
            logger.info(f"Opened quotation: {filename}")
        
        except Exception as e:
            logger.error(f"Error opening quotation: {e}")
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le devis:\n{str(e)}")

    def _delete_selected_quotation(self):
        """Delete the selected quotation file"""
        selection = self.quotations_tree.selection()
        if not selection:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un devis à supprimer.")
            return
        
        try:
            selected_item = selection[0]
            filename = self.quotations_tree.item(selected_item)['text']
            filepath = os.path.join(DEVIS_FOLDER, filename)
            
            # Confirm deletion
            if messagebox.askyesno("Confirmer suppression", f"Êtes-vous sûr de vouloir supprimer:\n{filename}?"):
                if os.path.exists(filepath):
                    os.remove(filepath)
                    logger.info(f"Deleted quotation: {filename}")
                    messagebox.showinfo("Succès", f"Le devis a été supprimé:\n{filename}")
                    # Refresh the list
                    self._refresh_quotations_list()
                else:
                    messagebox.showerror("Erreur", f"Le fichier n'existe pas:\n{filepath}")
        
        except Exception as e:
            logger.error(f"Error deleting quotation: {e}")
            messagebox.showerror("Erreur", f"Impossible de supprimer le devis:\n{str(e)}")
