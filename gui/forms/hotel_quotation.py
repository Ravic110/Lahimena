"""
Hotel quotation GUI component
"""

import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from config import *
from utils.excel_handler import load_all_hotels
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

        # Client information section
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

        # Row 1: Name and surname
        tk.Label(
            client_frame,
            text="Nom:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=0, column=0, sticky="w", pady=5)

        self.client_name_var = tk.StringVar()
        self.client_name_entry = tk.Entry(
            client_frame,
            textvariable=self.client_name_var,
            font=ENTRY_FONT,
            width=20,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.client_name_entry.grid(row=0, column=1, padx=(10, 20), pady=5)

        tk.Label(
            client_frame,
            text="Prénom:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=0, column=2, sticky="w", pady=5)

        self.client_surname_var = tk.StringVar()
        self.client_surname_entry = tk.Entry(
            client_frame,
            textvariable=self.client_surname_var,
            font=ENTRY_FONT,
            width=20,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.client_surname_entry.grid(row=0, column=3, padx=(10, 0), pady=5)

        # Row 2: Email and phone
        tk.Label(
            client_frame,
            text="Email:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=1, column=0, sticky="w", pady=5)

        self.client_email_var = tk.StringVar()
        self.client_email_entry = tk.Entry(
            client_frame,
            textvariable=self.client_email_var,
            font=ENTRY_FONT,
            width=20,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.client_email_entry.grid(row=1, column=1, padx=(10, 20), pady=5)

        tk.Label(
            client_frame,
            text="Téléphone:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).grid(row=1, column=2, sticky="w", pady=5)

        self.client_phone_var = tk.StringVar()
        self.client_phone_entry = tk.Entry(
            client_frame,
            textvariable=self.client_phone_var,
            font=ENTRY_FONT,
            width=20,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.client_phone_entry.grid(row=1, column=3, padx=(10, 0), pady=5)

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

    def _calculate_price(self):
        """Calculate the total price based on parameters"""
        if not self.selected_hotel:
            messagebox.showwarning("Hôtel non sélectionné", "Veuillez d'abord sélectionner un hôtel.")
            return

        try:
            # Get parameters
            nights = int(self.nights_var.get())
            adults = int(self.adults_var.get())
            children = int(self.children_var.get())
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

        except ValueError:
            messagebox.showerror("Erreur", "Veuillez saisir des valeurs numériques valides.")

    def _display_results(self, base_price, meal_price, total_price, nights, adults, children, room_type, meal_plan, client_type, currency, client_name, client_surname, client_email, client_phone):
        """Display calculation results"""
        # Get currency symbol
        currency_symbols = {
            "Ariary": "Ar",
            "Euro": "€",
            "Dollar US": "$"
        }
        symbol = currency_symbols.get(currency, "Ar")

        # Format prices
        if currency == "Ariary":
            price_format = ":,.0f"
        else:
            price_format = ":,.2f"

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
            # Create quotes directory if it doesn't exist
            quotes_dir = "devis"
            if not os.path.exists(quotes_dir):
                os.makedirs(quotes_dir)

            # Generate quote number and filename
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            quote_number = f"DEVIS_HOTEL_{timestamp}"
            filename = f"{quotes_dir}/{quote_number}.txt"

            # Get quote data
            results_content = self.results_text.get(1.0, tk.END).strip()

            # Get currency info
            currency = self.currency_var.get()
            currency_symbols = {
                "Ariary": "Ar",
                "Euro": "€",
                "Dollar US": "$"
            }
            symbol = currency_symbols.get(currency, "Ar")

            # Create detailed quote content
            quote_content = f"""
================================================================================
                            LAHIMENA TOURS
                       DEVIS D'HÉBERGEMENT HÔTEL
================================================================================

Numéro de devis: {quote_number}
Date de génération: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}

{results_content}

================================================================================
INFORMATIONS COMPLÉMENTAIRES:
- Hôtel: {self.selected_hotel['nom']}
- Localisation: {self.selected_hotel['lieu']}
- Catégorie: {self.selected_hotel['categorie']}
- Contact hôtel: {self.selected_hotel.get('contact', 'Non disponible')}
- Email hôtel: {self.selected_hotel.get('email', 'Non disponible')}

CONDITIONS:
- Prix en {currency} ({symbol})
- TVA non incluse
- Validité du devis: 30 jours
- Conditions d'annulation selon politique hôtelière

================================================================================
Généré par Lahimena Tours - Système de gestion de devis
================================================================================
"""

            # Save to file
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(quote_content.strip())

            # Show success message with file location
            messagebox.showinfo(
                "Devis généré avec succès",
                f"Le devis a été sauvegardé dans le fichier :\n{filename}\n\nLe fichier va s'ouvrir automatiquement."
            )

            # Try to open the file
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(filename)
                elif os.name == 'posix':  # macOS/Linux
                    subprocess.run(['xdg-open', filename])
            except Exception as e:
                messagebox.showwarning(
                    "Ouverture automatique impossible",
                    f"Le fichier a été créé mais n'a pas pu s'ouvrir automatiquement.\n\nVous pouvez l'ouvrir manuellement : {filename}"
                )

        except Exception as e:
            messagebox.showerror(
                "Erreur lors de la génération",
                f"Une erreur s'est produite lors de la génération du devis :\n{str(e)}"
            )

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