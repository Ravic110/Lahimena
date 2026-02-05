"""
Hotel form GUI component
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from config import *
from models.hotel_data import HotelData
from utils.validators import validate_email
from utils.excel_handler import save_hotel_to_excel, update_hotel_in_excel
from utils.logger import logger


class HotelForm:
    """
    Hotel form component
    """

    def __init__(self, parent, hotel_to_edit=None, on_save_callback=None):
        """
        Initialize hotel form

        Args:
            parent: Parent widget
            hotel_to_edit: Optional hotel dict to edit
            on_save_callback: Optional callback after saving
        """
        self.parent = parent
        self.hotel_to_edit = hotel_to_edit
        self.on_save_callback = on_save_callback

        self._create_form()

    def _create_form(self):
        """Create the hotel form"""
        # Clear parent
        for widget in self.parent.winfo_children():
            widget.destroy()

        # Title
        title_text = "MODIFIER HÔTEL" if self.hotel_to_edit else "AJOUTER HÔTEL"
        title = tk.Label(
            self.parent,
            text=title_text,
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        )
        title.pack(pady=(20, 10))

        # Create scrollable frame for the form
        form_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        form_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Create canvas and scrollbar for scrolling
        canvas = tk.Canvas(form_frame, bg=MAIN_BG_COLOR, highlightthickness=0)
        scrollbar = ttk.Scrollbar(form_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=MAIN_BG_COLOR)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enable mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Form content
        self._create_form_content(scrollable_frame)

        # Populate fields if editing
        if self.hotel_to_edit:
            self._populate_fields()

    def _create_form_content(self, parent):
        """Create form content"""
        # Basic Information Section
        basic_frame = tk.LabelFrame(parent, text="Informations de base", bg=MAIN_BG_COLOR, fg=TEXT_COLOR, font=LABEL_FONT)
        basic_frame.pack(fill="x", padx=10, pady=10)

        # Row 1: ID and Name
        row1 = tk.Frame(basic_frame, bg=MAIN_BG_COLOR)
        row1.pack(fill="x", padx=10, pady=5)

        tk.Label(row1, text="ID Hôtel *", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=0, column=0, sticky="w", pady=2)
        self.entry_id = tk.Entry(row1, font=ENTRY_FONT, width=15, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
        self.entry_id.grid(row=1, column=0, padx=(0, 10), pady=2)

        tk.Label(row1, text="Nom de l'hôtel *", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=0, column=1, sticky="w", pady=2)
        self.entry_nom = tk.Entry(row1, font=ENTRY_FONT, width=30, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
        self.entry_nom.grid(row=1, column=1, padx=(0, 10), pady=2)

        # Row 2: Location and Type
        row2 = tk.Frame(basic_frame, bg=MAIN_BG_COLOR)
        row2.pack(fill="x", padx=10, pady=5)

        tk.Label(row2, text="Lieu *", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=0, column=0, sticky="w", pady=2)
        self.entry_lieu = tk.Entry(row2, font=ENTRY_FONT, width=20, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
        self.entry_lieu.grid(row=1, column=0, padx=(0, 10), pady=2)

        tk.Label(row2, text="Type d'hébergement *", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=0, column=1, sticky="w", pady=2)
        self.combo_type = ttk.Combobox(row2, values=TYPE_HEBERGEMENTS, state="readonly", width=25)
        self.combo_type.grid(row=1, column=1, padx=(0, 10), pady=2)

        tk.Label(row2, text="Catégorie", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=0, column=2, sticky="w", pady=2)
        self.entry_categorie = tk.Entry(row2, font=ENTRY_FONT, width=10, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
        self.entry_categorie.grid(row=1, column=2, pady=2)

        # Contact Information Section
        contact_frame = tk.LabelFrame(parent, text="Informations de contact", bg=MAIN_BG_COLOR, fg=TEXT_COLOR, font=LABEL_FONT)
        contact_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(contact_frame, text="Contact", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.entry_contact = tk.Entry(contact_frame, font=ENTRY_FONT, width=25, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
        self.entry_contact.grid(row=1, column=0, padx=10, pady=5)

        tk.Label(contact_frame, text="Email", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=0, column=1, sticky="w", padx=10, pady=5)
        self.entry_email = tk.Entry(contact_frame, font=ENTRY_FONT, width=30, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
        self.entry_email.grid(row=1, column=1, padx=10, pady=5)

        # Pricing Section
        pricing_frame = tk.LabelFrame(parent, text="Tarification (en Ariary)", bg=MAIN_BG_COLOR, fg=TEXT_COLOR, font=LABEL_FONT)
        pricing_frame.pack(fill="x", padx=10, pady=10)

        # Room prices
        room_prices = [
            ("Chambre Single", "entry_single"),
            ("Chambre Double/Twin", "entry_double"),
            ("Chambre Familiale", "entry_familiale"),
            ("Lit Supplémentaire", "entry_lit_supp"),
            ("Day Use", "entry_day_use")
        ]

        for i, (label, attr) in enumerate(room_prices):
            row = i // 2
            col = (i % 2) * 2

            tk.Label(pricing_frame, text=label, font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=row, column=col, sticky="w", padx=10, pady=2)
            entry = tk.Entry(pricing_frame, font=ENTRY_FONT, width=15, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
            entry.grid(row=row, column=col+1, padx=10, pady=2)
            setattr(self, attr, entry)

        # Additional fees
        fees = [
            ("Vignette", "entry_vignette"),
            ("Taxe de séjour", "entry_taxe_sejour")
        ]

        for i, (label, attr) in enumerate(fees):
            row = 3 + i // 2
            col = (i % 2) * 2

            tk.Label(pricing_frame, text=label, font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=row, column=col, sticky="w", padx=10, pady=2)
            entry = tk.Entry(pricing_frame, font=ENTRY_FONT, width=15, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
            entry.grid(row=row, column=col+1, padx=10, pady=2)
            setattr(self, attr, entry)

        # Meal prices
        meal_frame = tk.LabelFrame(parent, text="Prix des repas (en Ariary)", bg=MAIN_BG_COLOR, fg=TEXT_COLOR, font=LABEL_FONT)
        meal_frame.pack(fill="x", padx=10, pady=10)

        meals = [
            ("Petit déjeuner", "entry_petit_dej"),
            ("Déjeuner", "entry_dejeuner"),
            ("Dîner", "entry_diner")
        ]

        for i, (label, attr) in enumerate(meals):
            tk.Label(meal_frame, text=label, font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR).grid(row=0, column=i*2, sticky="w", padx=10, pady=2)
            entry = tk.Entry(meal_frame, font=ENTRY_FONT, width=12, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
            entry.grid(row=1, column=i*2, padx=10, pady=2)
            setattr(self, attr, entry)

        # Description
        desc_frame = tk.LabelFrame(parent, text="Description", bg=MAIN_BG_COLOR, fg=TEXT_COLOR, font=LABEL_FONT)
        desc_frame.pack(fill="x", padx=10, pady=10)

        self.text_description = tk.Text(desc_frame, height=4, width=60, font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR)
        self.text_description.pack(padx=10, pady=10)

        # Buttons
        btn_frame = tk.Frame(parent, bg=MAIN_BG_COLOR)
        btn_frame.pack(pady=20)

        if self.hotel_to_edit:
            tk.Button(
                btn_frame,
                text="MODIFIER",
                command=self._validate,
                bg="#f39c12",
                fg="white",
                font=BUTTON_FONT,
                width=12
            ).pack(side="left", padx=5)
            tk.Button(
                btn_frame,
                text="ANNULER",
                command=self._cancel,
                bg="#95a5a6",
                fg="white",
                font=BUTTON_FONT,
                width=12
            ).pack(side="left", padx=5)
        else:
            tk.Button(
                btn_frame,
                text="AJOUTER",
                command=self._validate,
                bg="#27ae60",
                fg="white",
                font=BUTTON_FONT,
                width=12
            ).pack(side="left", padx=5)

    def _populate_fields(self):
        """Populate form fields with hotel data"""
        if not self.hotel_to_edit:
            return

        self.entry_id.insert(0, self.hotel_to_edit.get('id', ''))
        self.entry_nom.insert(0, self.hotel_to_edit.get('nom', ''))
        self.entry_lieu.insert(0, self.hotel_to_edit.get('lieu', ''))
        self.combo_type.set(self.hotel_to_edit.get('type_hebergement', ''))
        self.entry_categorie.insert(0, self.hotel_to_edit.get('categorie', ''))

        self.entry_single.insert(0, str(self.hotel_to_edit.get('chambre_single', 0)))
        self.entry_double.insert(0, str(self.hotel_to_edit.get('chambre_double', 0)))
        self.entry_familiale.insert(0, str(self.hotel_to_edit.get('chambre_familiale', 0)))
        self.entry_lit_supp.insert(0, str(self.hotel_to_edit.get('lit_supp', 0)))
        self.entry_day_use.insert(0, str(self.hotel_to_edit.get('day_use', 0)))
        self.entry_vignette.insert(0, str(self.hotel_to_edit.get('vignette', 0)))
        self.entry_taxe_sejour.insert(0, str(self.hotel_to_edit.get('taxe_sejour', 0)))

        self.entry_petit_dej.insert(0, str(self.hotel_to_edit.get('petit_dejeuner', 0)))
        self.entry_dejeuner.insert(0, str(self.hotel_to_edit.get('dejeuner', 0)))
        self.entry_diner.insert(0, str(self.hotel_to_edit.get('diner', 0)))

        self.entry_contact.insert(0, self.hotel_to_edit.get('contact', ''))
        self.entry_email.insert(0, self.hotel_to_edit.get('email', ''))

        description = self.hotel_to_edit.get('description', '')
        self.text_description.insert(1.0, description)

    def _validate(self):
        """Validate and save hotel data"""
        # Get form data
        form_data = {
            'id': self.entry_id.get().strip(),
            'nom': self.entry_nom.get().strip(),
            'lieu': self.entry_lieu.get().strip(),
            'type_hebergement': self.combo_type.get(),
            'categorie': self.entry_categorie.get().strip(),
            'chambre_single': self._parse_price(self.entry_single.get()),
            'chambre_double': self._parse_price(self.entry_double.get()),
            'chambre_familiale': self._parse_price(self.entry_familiale.get()),
            'lit_supp': self._parse_price(self.entry_lit_supp.get()),
            'day_use': self._parse_price(self.entry_day_use.get()),
            'vignette': self._parse_price(self.entry_vignette.get()),
            'taxe_sejour': self._parse_price(self.entry_taxe_sejour.get()),
            'petit_dejeuner': self._parse_price(self.entry_petit_dej.get()),
            'dejeuner': self._parse_price(self.entry_dejeuner.get()),
            'diner': self._parse_price(self.entry_diner.get()),
            'contact': self.entry_contact.get().strip(),
            'email': self.entry_email.get().strip(),
            'description': self.text_description.get(1.0, tk.END).strip()
        }

        # Create hotel data object
        hotel = HotelData.from_dict(form_data)

        # Validate
        errors = hotel.validate()

        # Additional validations
        if hotel.email and not validate_email(hotel.email):
            errors.append("Email invalide")

        if errors:
            messagebox.showerror("❌ Erreur", "\n".join(errors))
            return

        # Save or update
        try:
            if self.hotel_to_edit:
                # Update existing hotel
                success = update_hotel_in_excel(self.hotel_to_edit['row_number'], hotel.to_dict())
                if success:
                    messagebox.showinfo("✅ SUCCÈS", f"Hôtel {hotel.nom} modifié avec succès !")
                    logger.info(f"Hotel updated: {hotel.nom}")
                    if self.on_save_callback:
                        self.on_save_callback()
                else:
                    error_msg = "Erreur lors de la modification de l'hôtel. Voir les logs."
                    messagebox.showerror("❌ Erreur", error_msg)
                    logger.error(f"Failed to update hotel: {hotel.nom}")
            else:
                # Save new hotel
                row = save_hotel_to_excel(hotel.to_dict())
                if row > 0:
                    messagebox.showinfo("✅ SUCCÈS", f"Hôtel {hotel.nom} ajouté ligne Excel {row} !")
                    logger.info(f"New hotel saved: {hotel.nom} at row {row}")
                    self._reset_form()
                else:
                    error_msg = "Erreur lors de l'ajout de l'hôtel. Voir les logs."
                    messagebox.showerror("❌ Erreur", error_msg)
                    logger.error(f"Failed to save new hotel: {hotel.nom}")
        except Exception as e:
            error_msg = f"Erreur: {str(e)}\n\nDétails dans les logs de l'application."
            messagebox.showerror("❌ Erreur", error_msg)
            logger.error(f"Exception during hotel save/update: {e}", exc_info=True)

    def _parse_price(self, value):
        """Parse price value, return 0 if empty or invalid"""
        if not value.strip():
            return 0
        try:
            return float(value)
        except ValueError:
            return 0

    def _cancel(self):
        """Cancel editing and return to list"""
        if self.on_save_callback:
            self.on_save_callback()

    def _reset_form(self):
        """Reset all form fields"""
        # Clear all entries
        entries_to_clear = [
            self.entry_id, self.entry_nom, self.entry_lieu, self.entry_categorie,
            self.entry_single, self.entry_double, self.entry_familiale, self.entry_lit_supp,
            self.entry_day_use, self.entry_vignette, self.entry_taxe_sejour,
            self.entry_petit_dej, self.entry_dejeuner, self.entry_diner,
            self.entry_contact, self.entry_email
        ]

        for entry in entries_to_clear:
            entry.delete(0, tk.END)

        # Clear combobox
        self.combo_type.set('')

        # Clear text area
        self.text_description.delete(1.0, tk.END)