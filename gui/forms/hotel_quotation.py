"""
Hotel quotation GUI component
"""

import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from config import *
from utils.excel_handler import load_all_hotels, load_all_clients
from utils.logger import logger
from utils.pdf_generator import generate_hotel_quotation_pdf, REPORTLAB_AVAILABLE
import os
import datetime
import subprocess
from utils.validators import get_exchange_rates, convert_currency


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
        
        # Remove duplicates based on nom and lieu
        unique_hotels = {}
        for hotel in hotels:
            key = f"{hotel['nom']}_{hotel['lieu']}"
            if key not in unique_hotels:
                unique_hotels[key] = hotel
        
        return list(unique_hotels.values())

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
        client_names = [""] + [f"{client['ref_client']} - {client['nom']}" for client in self.clients]
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
            text="Nom:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=1, column=0, sticky="w", pady=5)

        self.client_name_var = tk.StringVar()
        self.client_name_entry = tk.Entry(
            client_frame,
            textvariable=self.client_name_var,
            font=ENTRY_FONT,
            width=20,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.client_name_entry.grid(row=1, column=1, padx=(10, 20), pady=5)

        tk.Label(
            client_frame,
            text="Prénom:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=1, column=2, sticky="w", pady=5)

        self.client_surname_var = tk.StringVar()
        self.client_surname_entry = tk.Entry(
            client_frame,
            textvariable=self.client_surname_var,
            font=ENTRY_FONT,
            width=20,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.client_surname_entry.grid(row=1, column=3, padx=(10, 0), pady=5)

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

        # Hotel selection
        tk.Label(
            hotel_frame,
            text="Hôtel:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=0, column=0, sticky="w", pady=5)

        self.hotel_var = tk.StringVar()
        hotel_names = [f"{hotel['nom']} - {hotel['lieu']}" for hotel in self.hotels]
        self.hotel_combo = ttk.Combobox(
            hotel_frame,
            textvariable=self.hotel_var,
            values=hotel_names,
            font=ENTRY_FONT,
            width=40,
            state="readonly"
        )
        self.hotel_combo.grid(row=0, column=1, padx=(10, 0), pady=5)
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

        # Row 1: Number of nights and room type
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
            text="Type de chambre:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=0, column=2, sticky="w", pady=5)

        self.room_type_var = tk.StringVar(value="Double")
        self.room_type_combo = ttk.Combobox(
            params_frame,
            textvariable=self.room_type_var,
            values=TYPE_CHAMBRES,
            font=ENTRY_FONT,
            width=15,
            state="readonly"
        )
        self.room_type_combo.grid(row=0, column=3, padx=(10, 0), pady=5)

        # Row 2: Number of adults and children
        tk.Label(
            params_frame,
            text="Adultes:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=1, column=0, sticky="w", pady=5)

        self.adults_var = tk.StringVar(value="2")
        self.adults_entry = tk.Entry(
            params_frame,
            textvariable=self.adults_var,
            font=ENTRY_FONT,
            width=10,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.adults_entry.grid(row=1, column=1, padx=(10, 20), pady=5)

        tk.Label(
            params_frame,
            text="Enfants:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=1, column=2, sticky="w", pady=5)

        self.children_var = tk.StringVar(value="0")
        self.children_entry = tk.Entry(
            params_frame,
            textvariable=self.children_var,
            font=ENTRY_FONT,
            width=10,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.children_entry.grid(row=1, column=3, padx=(10, 0), pady=5)

        # Row 2: Client type
        tk.Label(
            params_frame,
            text="Type de client:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=2, column=0, sticky="w", pady=5)

        self.client_type_var = tk.StringVar(value="PBC")
        self.client_type_combo = ttk.Combobox(
            params_frame,
            textvariable=self.client_type_var,
            values=["TO", "PBC"],
            font=ENTRY_FONT,
            width=10,
            state="readonly"
        )
        self.client_type_combo.grid(row=2, column=1, padx=(10, 0), pady=5)

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

    def _on_hotel_selected(self, event=None):
        """Handle hotel selection"""
        selection = self.hotel_var.get()
        if selection:
            # Find the selected hotel
            for hotel in self.hotels:
                hotel_display = f"{hotel['nom']} - {hotel['lieu']}"
                if hotel_display == selection:
                    self.selected_hotel = hotel
                    break

    def _on_client_selected(self, event=None):
        """Handle client selection and auto-fill fields"""
        selection = self.client_var.get()
        if selection:
            # Find the selected client
            for client in self.clients:
                client_display = f"{client['ref_client']} - {client['nom']}"
                if client_display == selection:
                    # Auto-fill the fields
                    # Assuming the 'nom' field contains the full name or last name
                    # We'll put it in the name field and leave surname empty for now
                    self.client_name_var.set(client['nom'])
                    self.client_surname_var.set("")  # Can be adjusted based on data structure
                    self.client_email_var.set(client['email'])
                    self.client_phone_var.set(client['telephone'])
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

                    # Room type: match against TYPE_CHAMBRES or keywords
                    ch = client.get('chambre')
                    if ch:
                        try:
                            ch_str = str(ch)
                            # direct match
                            if ch_str in TYPE_CHAMBRES:
                                self.room_type_var.set(ch_str)
                            else:
                                # try case-insensitive keyword match
                                lower = ch_str.lower()
                                for t in TYPE_CHAMBRES:
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
            self.client_surname_var.set("")
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
        hotel_names = [f"{hotel['nom']} - {hotel['lieu']}" for hotel in self.hotels]
        self.hotel_combo['values'] = hotel_names
        
        # Reset hotel selection
        self.hotel_var.set("")
        self.selected_hotel = None

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
                
            room_type = self.room_type_var.get()

            # Get prices from hotel data
            room_price = 0
            if room_type == "Single" and self.selected_hotel.get('chambre_single'):
                room_price = self.selected_hotel['chambre_single']
            elif room_type == "Double" and self.selected_hotel.get('chambre_double'):
                room_price = self.selected_hotel['chambre_double']
            elif room_type == "Triple" and self.selected_hotel.get('chambre_double'):
                room_price = self.selected_hotel['chambre_double']  # Approximation
            elif room_type == "Familliale" and self.selected_hotel.get('chambre_familiale'):
                room_price = self.selected_hotel['chambre_familiale']

            if room_price == 0:
                messagebox.showwarning("Prix non disponible", f"Le prix pour {room_type} n'est pas disponible pour cet hôtel.")
                return

            # Calculate base price
            base_price = room_price * nights

            # Calculate meal supplements
            meal_price = 0
            meal_plan = self.meal_var.get()
            if meal_plan == "Petit déjeuner" and self.selected_hotel.get('petit_dejeuner'):
                meal_price = self.selected_hotel['petit_dejeuner'] * nights * (adults + children)
            elif meal_plan == "Demi-pension" and self.selected_hotel.get('dejeuner') and self.selected_hotel.get('diner'):
                meal_price = (self.selected_hotel['dejeuner'] + self.selected_hotel['diner']) * nights * (adults + children)
            elif meal_plan == "Pension complète" and self.selected_hotel.get('petit_dejeuner') and self.selected_hotel.get('dejeuner') and self.selected_hotel.get('diner'):
                meal_price = (self.selected_hotel['petit_dejeuner'] + self.selected_hotel['dejeuner'] + self.selected_hotel['diner']) * nights * (adults + children)

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
            client_surname = self.client_surname_var.get()
            client_email = self.client_email_var.get()
            client_phone = self.client_phone_var.get()
            self._display_results(base_price, meal_price, total_price, nights, adults, children, room_type, meal_plan, client_type, currency, client_name, client_surname, client_email, client_phone)

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

    def _display_results(self, base_price, meal_price, total_price, nights, adults, children, room_type, meal_plan, client_type, currency, client_name, client_surname, client_email, client_phone):
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
- Nom: {client_name}
- Prénom: {client_surname}
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
            room_type = self.room_type_var.get() if hasattr(self, 'room_type_var') else "Standard"
            nights = int(self.nights_spinbox.get()) if hasattr(self, 'nights_spinbox') else 1
            adults = int(self.adults_spinbox.get()) if hasattr(self, 'adults_spinbox') else 1
            
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

            # Show success message with file location
            messagebox.showinfo(
                "✅ Devis généré avec succès",
                f"Le devis PDF a été sauvegardé :\n{filename}\n\nLe fichier va s'ouvrir automatiquement."
            )
            logger.info(f"PDF quotation generated successfully: {filename}")

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
        self.room_type_var.set("Double")
        self.client_type_var.set("PBC")
        self.currency_var.set("Ariary")
        self.period_var.set("Moyenne saison")
        self.meal_var.set("Petit déjeuner")
        self.client_name_var.set("")
        self.client_surname_var.set("")
        self.client_email_var.set("")
        self.client_phone_var.set("")
        self.results_text.delete(1.0, tk.END)