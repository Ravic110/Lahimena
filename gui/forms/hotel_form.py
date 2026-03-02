"""
Hotel form GUI component
"""

import tkinter as tk
from datetime import datetime
import re
import unicodedata
from tkinter import messagebox, ttk

from config import (
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GRAY,
    BUTTON_GREEN,
    BUTTON_ORANGE,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
    TYPE_HEBERGEMENTS,
)
from models.hotel_data import HotelData
from utils.excel_handler import save_hotel_to_excel, update_hotel_in_excel
from utils.logger import logger
from utils.validators import validate_email


class HotelForm:
    """
    Hotel form component
    """
    CATEGORIE_OPTIONS = ("TCO", "PCB", "DU")

    def __init__(
        self, parent, hotel_to_edit=None, on_save_callback=None, on_back_to_db=None
    ):
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
        self.on_back_to_db = on_back_to_db

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
            bg=MAIN_BG_COLOR,
        )
        title.pack(pady=(20, 10))

        # Create scrollable frame for the form
        form_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        form_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Form content
        self._create_form_content(form_frame)

        # Populate fields if editing
        if self.hotel_to_edit:
            self._populate_fields()

    def _create_form_content(self, parent):
        """Create form content"""
        # Basic Information Section
        basic_frame = tk.LabelFrame(
            parent,
            text="Informations de base",
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
            font=LABEL_FONT,
        )
        basic_frame.pack(fill="x", padx=10, pady=10)

        # Row 1: ID and Name
        row1 = tk.Frame(basic_frame, bg=MAIN_BG_COLOR)
        row1.pack(fill="x", padx=10, pady=5)

        tk.Label(
            row1, text="ID Hôtel (auto)", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR
        ).grid(row=0, column=0, sticky="w", pady=2)
        self.entry_id = tk.Entry(
            row1, font=ENTRY_FONT, width=20, bg=INPUT_BG_COLOR, fg=TEXT_COLOR, state="readonly"
        )
        self.entry_id.grid(row=1, column=0, padx=(0, 10), pady=2)

        tk.Label(
            row1,
            text="Nom de l'hôtel *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=0, column=1, sticky="w", pady=2)
        self.entry_nom = tk.Entry(
            row1, font=ENTRY_FONT, width=30, bg=INPUT_BG_COLOR, fg=TEXT_COLOR
        )
        self.entry_nom.grid(row=1, column=1, padx=(0, 10), pady=2)
        self.entry_nom.bind("<KeyRelease>", lambda _e: self._set_generated_id())

        # Row 2: Location and Type
        row2 = tk.Frame(basic_frame, bg=MAIN_BG_COLOR)
        row2.pack(fill="x", padx=10, pady=5)

        tk.Label(
            row2, text="Lieu *", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR
        ).grid(row=0, column=0, sticky="w", pady=2)
        self.entry_lieu = tk.Entry(
            row2, font=ENTRY_FONT, width=20, bg=INPUT_BG_COLOR, fg=TEXT_COLOR
        )
        self.entry_lieu.grid(row=1, column=0, padx=(0, 10), pady=2)
        self.entry_lieu.bind("<KeyRelease>", lambda _e: self._set_generated_id())

        tk.Label(
            row2,
            text="Type d'hébergement *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=0, column=1, sticky="w", pady=2)
        self.combo_type = ttk.Combobox(
            row2, values=TYPE_HEBERGEMENTS, state="readonly", width=25
        )
        self.combo_type.grid(row=1, column=1, padx=(0, 10), pady=2)

        tk.Label(
            row2, text="Catégorie", font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR
        ).grid(row=0, column=2, sticky="w", pady=2)
        self.combo_categorie = ttk.Combobox(
            row2, values=self.CATEGORIE_OPTIONS, state="readonly", width=10
        )
        self.combo_categorie.grid(row=1, column=2, pady=2)
        self.combo_categorie.set("TCO")
        self.combo_categorie.bind("<<ComboboxSelected>>", lambda _e: self._set_generated_id())

        # Contact Information Section
        contact_frame = tk.LabelFrame(
            parent,
            text="Informations de contact",
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
            font=LABEL_FONT,
        )
        contact_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(
            contact_frame,
            text="Contact",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.entry_contact = tk.Entry(
            contact_frame, font=ENTRY_FONT, width=25, bg=INPUT_BG_COLOR, fg=TEXT_COLOR
        )
        self.entry_contact.grid(row=1, column=0, padx=10, pady=5)

        tk.Label(
            contact_frame,
            text="Email",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).grid(row=0, column=1, sticky="w", padx=10, pady=5)
        self.entry_email = tk.Entry(
            contact_frame, font=ENTRY_FONT, width=30, bg=INPUT_BG_COLOR, fg=TEXT_COLOR
        )
        self.entry_email.grid(row=1, column=1, padx=10, pady=5)

        # Pricing Section
        pricing_frame = tk.LabelFrame(
            parent,
            text="Tarification (en Ariary)",
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
            font=LABEL_FONT,
        )
        pricing_frame.pack(fill="x", padx=10, pady=10)

        # Room prices (STANDARD)
        room_prices = [
            ("Chambre Single", "entry_single"),
            ("Chambre Double", "entry_double"),
            ("Chambre Twin", "entry_twin"),
            ("Chambre Familiale", "entry_familiale"),
            ("Chambre Triple", "entry_triple"),
            ("Chambre Chauffeur", "entry_chauffeur"),
            ("Dortoir", "entry_dortoir"),
            ("Lit Supplémentaire", "entry_lit_supp"),
            ("Day Use", "entry_day_use"),
        ]

        for i, (label, attr) in enumerate(room_prices):
            row = i // 3
            col = (i % 3) * 2

            tk.Label(
                pricing_frame,
                text=label,
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=row, column=col, sticky="w", padx=10, pady=2)
            entry = tk.Entry(
                pricing_frame,
                font=ENTRY_FONT,
                width=15,
                bg=INPUT_BG_COLOR,
                fg=TEXT_COLOR,
            )
            entry.grid(row=row, column=col + 1, padx=10, pady=2)
            setattr(self, attr, entry)

        # Additional fees
        fees = [("Vignette", "entry_vignette"), ("Taxe de séjour", "entry_taxe_sejour")]

        for i, (label, attr) in enumerate(fees):
            row = 4 + i // 2
            col = (i % 2) * 2

            tk.Label(
                pricing_frame,
                text=label,
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=row, column=col, sticky="w", padx=10, pady=2)
            entry = tk.Entry(
                pricing_frame,
                font=ENTRY_FONT,
                width=15,
                bg=INPUT_BG_COLOR,
                fg=TEXT_COLOR,
            )
            entry.grid(row=row, column=col + 1, padx=10, pady=2)
            setattr(self, attr, entry)

        # Villa prices (current DB has a dedicated Villa group)
        villa_frame = tk.LabelFrame(
            parent,
            text="Tarification Villa (en Ariary)",
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
            font=LABEL_FONT,
        )
        villa_frame.pack(fill="x", padx=10, pady=10)

        villa_prices = [
            ("Villa Single", "entry_villa_single"),
            ("Villa Double", "entry_villa_double"),
            ("Villa Twin", "entry_villa_twin"),
            ("Villa Familiale", "entry_villa_familiale"),
            ("Villa Triple", "entry_villa_triple"),
            ("Villa Studios", "entry_villa_studios"),
            ("Villa VIP", "entry_villa_vip"),
            ("Villa Supp", "entry_villa_supp"),
        ]
        for i, (label, attr) in enumerate(villa_prices):
            row = (i // 4) * 2
            col = (i % 4) * 2
            tk.Label(
                villa_frame,
                text=label,
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=row, column=col, sticky="w", padx=10, pady=2)
            entry = tk.Entry(
                villa_frame,
                font=ENTRY_FONT,
                width=12,
                bg=INPUT_BG_COLOR,
                fg=TEXT_COLOR,
            )
            entry.grid(row=row + 1, column=col, padx=10, pady=2)
            setattr(self, attr, entry)

        bungalow_frame = tk.LabelFrame(
            parent,
            text="Tarification Bungalows (en Ariary)",
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
            font=LABEL_FONT,
        )
        bungalow_frame.pack(fill="x", padx=10, pady=10)
        bungalow_prices = [
            ("Bungalow Single", "entry_bungalow_single"),
            ("Bungalow Double", "entry_bungalow_double"),
            ("Bungalow Twin", "entry_bungalow_twin"),
            ("Bungalow Familiale", "entry_bungalow_familiale"),
            ("Bungalow Triple", "entry_bungalow_triple"),
            ("Bungalow Supp", "entry_bungalow_supp"),
        ]
        for i, (label, attr) in enumerate(bungalow_prices):
            row = (i // 3) * 2
            col = (i % 3) * 2
            tk.Label(
                bungalow_frame, text=label, font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR
            ).grid(row=row, column=col, sticky="w", padx=10, pady=2)
            entry = tk.Entry(
                bungalow_frame, font=ENTRY_FONT, width=12, bg=INPUT_BG_COLOR, fg=TEXT_COLOR
            )
            entry.grid(row=row + 1, column=col, padx=10, pady=2)
            setattr(self, attr, entry)

        deluxe_frame = tk.LabelFrame(
            parent,
            text="Tarification De Luxe (en Ariary)",
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
            font=LABEL_FONT,
        )
        deluxe_frame.pack(fill="x", padx=10, pady=10)
        deluxe_prices = [
            ("De Luxe Single", "entry_deluxe_single"),
            ("De Luxe Double", "entry_deluxe_double"),
            ("De Luxe Twin", "entry_deluxe_twin"),
            ("De Luxe Familiale", "entry_deluxe_familiale"),
            ("De Luxe Triple", "entry_deluxe_triple"),
            ("De Luxe Supp", "entry_deluxe_supp"),
        ]
        for i, (label, attr) in enumerate(deluxe_prices):
            row = (i // 3) * 2
            col = (i % 3) * 2
            tk.Label(
                deluxe_frame, text=label, font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR
            ).grid(row=row, column=col, sticky="w", padx=10, pady=2)
            entry = tk.Entry(
                deluxe_frame, font=ENTRY_FONT, width=12, bg=INPUT_BG_COLOR, fg=TEXT_COLOR
            )
            entry.grid(row=row + 1, column=col, padx=10, pady=2)
            setattr(self, attr, entry)

        suite_frame = tk.LabelFrame(
            parent,
            text="Tarification Suite (en Ariary)",
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
            font=LABEL_FONT,
        )
        suite_frame.pack(fill="x", padx=10, pady=10)
        suite_prices = [
            ("Suite Single", "entry_suite_single"),
            ("Suite Double", "entry_suite_double"),
            ("Suite Twin", "entry_suite_twin"),
            ("Suite Familiale", "entry_suite_familiale"),
            ("Suite Triple", "entry_suite_triple"),
            ("Suite Studios", "entry_suite_studios"),
            ("Suite VIP", "entry_suite_vip"),
            ("Suite Supp", "entry_suite_supp"),
        ]
        for i, (label, attr) in enumerate(suite_prices):
            row = (i // 4) * 2
            col = (i % 4) * 2
            tk.Label(
                suite_frame, text=label, font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR
            ).grid(row=row, column=col, sticky="w", padx=10, pady=2)
            entry = tk.Entry(
                suite_frame, font=ENTRY_FONT, width=12, bg=INPUT_BG_COLOR, fg=TEXT_COLOR
            )
            entry.grid(row=row + 1, column=col, padx=10, pady=2)
            setattr(self, attr, entry)

        # Meal prices
        meal_frame = tk.LabelFrame(
            parent,
            text="Prix des repas (en Ariary)",
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
            font=LABEL_FONT,
        )
        meal_frame.pack(fill="x", padx=10, pady=10)

        meals = [
            ("Petit déjeuner", "entry_petit_dej"),
            ("Déjeuner", "entry_dejeuner"),
            ("Dîner", "entry_diner"),
        ]

        for i, (label, attr) in enumerate(meals):
            tk.Label(
                meal_frame, text=label, font=LABEL_FONT, fg=TEXT_COLOR, bg=MAIN_BG_COLOR
            ).grid(row=0, column=i * 2, sticky="w", padx=10, pady=2)
            entry = tk.Entry(
                meal_frame, font=ENTRY_FONT, width=12, bg=INPUT_BG_COLOR, fg=TEXT_COLOR
            )
            entry.grid(row=1, column=i * 2, padx=10, pady=2)
            setattr(self, attr, entry)

        # Description
        desc_frame = tk.LabelFrame(
            parent, text="Description", bg=MAIN_BG_COLOR, fg=TEXT_COLOR, font=LABEL_FONT
        )
        desc_frame.pack(fill="x", padx=10, pady=10)

        self.text_description = tk.Text(
            desc_frame,
            height=4,
            width=60,
            font=ENTRY_FONT,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        )
        self.text_description.pack(padx=10, pady=10)

        inclus_frame = tk.Frame(parent, bg=MAIN_BG_COLOR)
        inclus_frame.pack(fill="x", padx=10, pady=(0, 10))
        tk.Label(
            inclus_frame,
            text="Inclus / remarques courtes",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(anchor="w")
        self.entry_inclus = tk.Entry(
            inclus_frame, font=ENTRY_FONT, width=60, bg=INPUT_BG_COLOR, fg=TEXT_COLOR
        )
        self.entry_inclus.pack(anchor="w", pady=(2, 0))

        # Buttons
        btn_frame = tk.Frame(parent, bg=MAIN_BG_COLOR)
        btn_frame.pack(pady=20)

        if self.hotel_to_edit:
            tk.Button(
                btn_frame,
                text="MODIFIER",
                command=self._validate,
                bg=BUTTON_ORANGE,
                fg="white",
                font=BUTTON_FONT,
                width=12,
            ).pack(side="left", padx=5)
            tk.Button(
                btn_frame,
                text="ANNULER",
                command=self._cancel,
                bg=BUTTON_GRAY,
                fg="white",
                font=BUTTON_FONT,
                width=12,
            ).pack(side="left", padx=5)
        else:
            tk.Button(
                btn_frame,
                text="AJOUTER",
                command=self._validate,
                bg=BUTTON_GREEN,
                fg="white",
                font=BUTTON_FONT,
                width=12,
            ).pack(side="left", padx=5)

        if self.on_back_to_db:
            tk.Button(
                btn_frame,
                text="⬅ RETOUR BDD",
                command=self._go_back_to_db,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                width=14,
            ).pack(side="left", padx=5)

    def _go_back_to_db(self):
        if self.on_back_to_db:
            self.on_back_to_db()

    def _populate_fields(self):
        """Populate form fields with hotel data"""
        if not self.hotel_to_edit:
            return

        self._set_entry_value(self.entry_id, self.hotel_to_edit.get("id", ""))
        self.entry_nom.insert(0, self.hotel_to_edit.get("nom", ""))
        self.entry_lieu.insert(0, self.hotel_to_edit.get("lieu", ""))
        self.combo_type.set(self.hotel_to_edit.get("type_hebergement", ""))
        categorie_value = self.hotel_to_edit.get("categorie", "") or self.hotel_to_edit.get(
            "type_client", ""
        )
        if categorie_value in self.CATEGORIE_OPTIONS:
            self.combo_categorie.set(categorie_value)
        else:
            self.combo_categorie.set("TCO")
        self._set_generated_id()

        self.entry_single.insert(0, str(self.hotel_to_edit.get("chambre_single", 0)))
        self.entry_double.insert(0, str(self.hotel_to_edit.get("chambre_double", 0)))
        self.entry_twin.insert(0, str(self.hotel_to_edit.get("chambre_twin", 0)))
        self.entry_familiale.insert(
            0, str(self.hotel_to_edit.get("chambre_familiale", 0))
        )
        self.entry_triple.insert(0, str(self.hotel_to_edit.get("chambre_triple", 0)))
        self.entry_chauffeur.insert(0, str(self.hotel_to_edit.get("chambre_chauffeur", 0)))
        self.entry_dortoir.insert(0, str(self.hotel_to_edit.get("dortoir", 0)))
        self.entry_lit_supp.insert(0, str(self.hotel_to_edit.get("lit_supp", 0)))
        self.entry_day_use.insert(0, str(self.hotel_to_edit.get("day_use", 0)))
        self.entry_villa_single.insert(0, str(self.hotel_to_edit.get("villa_single", 0)))
        self.entry_villa_double.insert(0, str(self.hotel_to_edit.get("villa_double", 0)))
        self.entry_villa_twin.insert(0, str(self.hotel_to_edit.get("villa_twin", 0)))
        self.entry_villa_familiale.insert(0, str(self.hotel_to_edit.get("villa_familiale", 0)))
        self.entry_villa_triple.insert(0, str(self.hotel_to_edit.get("villa_triple", 0)))
        self.entry_villa_studios.insert(0, str(self.hotel_to_edit.get("villa_studios", 0)))
        self.entry_villa_vip.insert(0, str(self.hotel_to_edit.get("villa_vip", 0)))
        self.entry_villa_supp.insert(0, str(self.hotel_to_edit.get("villa_supp", 0)))
        self.entry_bungalow_single.insert(0, str(self.hotel_to_edit.get("bungalow_single", 0)))
        self.entry_bungalow_double.insert(0, str(self.hotel_to_edit.get("bungalow_double", 0)))
        self.entry_bungalow_twin.insert(0, str(self.hotel_to_edit.get("bungalow_twin", 0)))
        self.entry_bungalow_familiale.insert(0, str(self.hotel_to_edit.get("bungalow_familiale", 0)))
        self.entry_bungalow_triple.insert(0, str(self.hotel_to_edit.get("bungalow_triple", 0)))
        self.entry_bungalow_supp.insert(0, str(self.hotel_to_edit.get("bungalow_supp", 0)))
        self.entry_deluxe_single.insert(0, str(self.hotel_to_edit.get("deluxe_single", 0)))
        self.entry_deluxe_double.insert(0, str(self.hotel_to_edit.get("deluxe_double", 0)))
        self.entry_deluxe_twin.insert(0, str(self.hotel_to_edit.get("deluxe_twin", 0)))
        self.entry_deluxe_familiale.insert(0, str(self.hotel_to_edit.get("deluxe_familiale", 0)))
        self.entry_deluxe_triple.insert(0, str(self.hotel_to_edit.get("deluxe_triple", 0)))
        self.entry_deluxe_supp.insert(0, str(self.hotel_to_edit.get("deluxe_supp", 0)))
        self.entry_suite_single.insert(0, str(self.hotel_to_edit.get("suite_single", 0)))
        self.entry_suite_double.insert(0, str(self.hotel_to_edit.get("suite_double", 0)))
        self.entry_suite_twin.insert(0, str(self.hotel_to_edit.get("suite_twin", 0)))
        self.entry_suite_familiale.insert(0, str(self.hotel_to_edit.get("suite_familiale", 0)))
        self.entry_suite_triple.insert(0, str(self.hotel_to_edit.get("suite_triple", 0)))
        self.entry_suite_studios.insert(0, str(self.hotel_to_edit.get("suite_studios", 0)))
        self.entry_suite_vip.insert(0, str(self.hotel_to_edit.get("suite_vip", 0)))
        self.entry_suite_supp.insert(0, str(self.hotel_to_edit.get("suite_supp", 0)))
        self.entry_vignette.insert(0, str(self.hotel_to_edit.get("vignette", 0)))
        self.entry_taxe_sejour.insert(0, str(self.hotel_to_edit.get("taxe_sejour", 0)))

        self.entry_petit_dej.insert(0, str(self.hotel_to_edit.get("petit_dejeuner", 0)))
        self.entry_dejeuner.insert(0, str(self.hotel_to_edit.get("dejeuner", 0)))
        self.entry_diner.insert(0, str(self.hotel_to_edit.get("diner", 0)))

        self.entry_contact.insert(0, self.hotel_to_edit.get("contact", ""))
        self.entry_email.insert(0, self.hotel_to_edit.get("email", ""))
        self.entry_inclus.insert(0, self.hotel_to_edit.get("inclus", ""))

        description = self.hotel_to_edit.get("description", "")
        self.text_description.insert(1.0, description)

    def _validate(self):
        """Validate and save hotel data"""
        self._set_generated_id()
        # Get form data
        form_data = {
            "id": self.entry_id.get().strip(),
            "nom": self.entry_nom.get().strip(),
            "lieu": self.entry_lieu.get().strip(),
            "type_hebergement": self.combo_type.get(),
            "categorie": self.combo_categorie.get().strip() or "TCO",
            "type_client": self.combo_categorie.get().strip() or "TCO",
            "unite": "MGA",
            "chambre_single": self._parse_price(self.entry_single.get()),
            "chambre_double": self._parse_price(self.entry_double.get()),
            "chambre_twin": self._parse_price(self.entry_twin.get()),
            "chambre_familiale": self._parse_price(self.entry_familiale.get()),
            "chambre_triple": self._parse_price(self.entry_triple.get()),
            "chambre_chauffeur": self._parse_price(self.entry_chauffeur.get()),
            "dortoir": self._parse_price(self.entry_dortoir.get()),
            "lit_supp": self._parse_price(self.entry_lit_supp.get()),
            "day_use": self._parse_price(self.entry_day_use.get()),
            "villa_single": self._parse_price(self.entry_villa_single.get()),
            "villa_double": self._parse_price(self.entry_villa_double.get()),
            "villa_twin": self._parse_price(self.entry_villa_twin.get()),
            "villa_familiale": self._parse_price(self.entry_villa_familiale.get()),
            "villa_triple": self._parse_price(self.entry_villa_triple.get()),
            "villa_studios": self._parse_price(self.entry_villa_studios.get()),
            "villa_vip": self._parse_price(self.entry_villa_vip.get()),
            "villa_supp": self._parse_price(self.entry_villa_supp.get()),
            "bungalow_single": self._parse_price(self.entry_bungalow_single.get()),
            "bungalow_double": self._parse_price(self.entry_bungalow_double.get()),
            "bungalow_twin": self._parse_price(self.entry_bungalow_twin.get()),
            "bungalow_familiale": self._parse_price(self.entry_bungalow_familiale.get()),
            "bungalow_triple": self._parse_price(self.entry_bungalow_triple.get()),
            "bungalow_supp": self._parse_price(self.entry_bungalow_supp.get()),
            "deluxe_single": self._parse_price(self.entry_deluxe_single.get()),
            "deluxe_double": self._parse_price(self.entry_deluxe_double.get()),
            "deluxe_twin": self._parse_price(self.entry_deluxe_twin.get()),
            "deluxe_familiale": self._parse_price(self.entry_deluxe_familiale.get()),
            "deluxe_triple": self._parse_price(self.entry_deluxe_triple.get()),
            "deluxe_supp": self._parse_price(self.entry_deluxe_supp.get()),
            "suite_single": self._parse_price(self.entry_suite_single.get()),
            "suite_double": self._parse_price(self.entry_suite_double.get()),
            "suite_twin": self._parse_price(self.entry_suite_twin.get()),
            "suite_familiale": self._parse_price(self.entry_suite_familiale.get()),
            "suite_triple": self._parse_price(self.entry_suite_triple.get()),
            "suite_studios": self._parse_price(self.entry_suite_studios.get()),
            "suite_vip": self._parse_price(self.entry_suite_vip.get()),
            "suite_supp": self._parse_price(self.entry_suite_supp.get()),
            "vignette": self._parse_price(self.entry_vignette.get()),
            "taxe_sejour": self._parse_price(self.entry_taxe_sejour.get()),
            "petit_dejeuner": self._parse_price(self.entry_petit_dej.get()),
            "dejeuner": self._parse_price(self.entry_dejeuner.get()),
            "diner": self._parse_price(self.entry_diner.get()),
            "inclus": self.entry_inclus.get().strip(),
            "contact": self.entry_contact.get().strip(),
            "email": self.entry_email.get().strip(),
            "description": self.text_description.get(1.0, tk.END).strip(),
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
                success = update_hotel_in_excel(
                    self.hotel_to_edit["row_number"], hotel.to_dict()
                )
                if success:
                    messagebox.showinfo(
                        "✅ SUCCÈS", f"Hôtel {hotel.nom} modifié avec succès !"
                    )
                    logger.info(f"Hotel updated: {hotel.nom}")
                    if self.on_save_callback:
                        self.on_save_callback()
                else:
                    error_msg = (
                        "Erreur lors de la modification de l'hôtel. Voir les logs."
                    )
                    messagebox.showerror("❌ Erreur", error_msg)
                    logger.error(f"Failed to update hotel: {hotel.nom}")
            else:
                # Save new hotel
                row = save_hotel_to_excel(hotel.to_dict())
                if row > 0:
                    messagebox.showinfo(
                        "✅ SUCCÈS", f"Hôtel {hotel.nom} ajouté ligne Excel {row} !"
                    )
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

    def _set_entry_value(self, entry, value):
        state = str(entry.cget("state"))
        if state == "readonly":
            entry.configure(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, value)
            entry.configure(state="readonly")
            return
        entry.delete(0, tk.END)
        entry.insert(0, value)

    def _normalize_id_part(self, value):
        text = unicodedata.normalize("NFKD", str(value or "").strip().lower())
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        text = re.sub(r"[^a-z0-9]+", "_", text)
        return text.strip("_")

    def _build_generated_hotel_id(self):
        lieu = self._normalize_id_part(self.entry_lieu.get())
        nom = self._normalize_id_part(self.entry_nom.get())
        categorie = self._normalize_id_part(self.combo_categorie.get())
        parts = [part for part in (lieu, nom, categorie) if part]
        return "_".join(parts)

    def _set_generated_id(self):
        generated_id = self._build_generated_hotel_id()
        self._set_entry_value(self.entry_id, generated_id)

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
            self.entry_id,
            self.entry_nom,
            self.entry_lieu,
            self.entry_single,
            self.entry_double,
            self.entry_twin,
            self.entry_familiale,
            self.entry_triple,
            self.entry_chauffeur,
            self.entry_dortoir,
            self.entry_lit_supp,
            self.entry_day_use,
            self.entry_villa_single,
            self.entry_villa_double,
            self.entry_villa_twin,
            self.entry_villa_familiale,
            self.entry_villa_triple,
            self.entry_villa_studios,
            self.entry_villa_vip,
            self.entry_villa_supp,
            self.entry_bungalow_single,
            self.entry_bungalow_double,
            self.entry_bungalow_twin,
            self.entry_bungalow_familiale,
            self.entry_bungalow_triple,
            self.entry_bungalow_supp,
            self.entry_deluxe_single,
            self.entry_deluxe_double,
            self.entry_deluxe_twin,
            self.entry_deluxe_familiale,
            self.entry_deluxe_triple,
            self.entry_deluxe_supp,
            self.entry_suite_single,
            self.entry_suite_double,
            self.entry_suite_twin,
            self.entry_suite_familiale,
            self.entry_suite_triple,
            self.entry_suite_studios,
            self.entry_suite_vip,
            self.entry_suite_supp,
            self.entry_vignette,
            self.entry_taxe_sejour,
            self.entry_petit_dej,
            self.entry_dejeuner,
            self.entry_diner,
            self.entry_inclus,
            self.entry_contact,
            self.entry_email,
        ]

        for entry in entries_to_clear:
            entry.delete(0, tk.END)

        # Clear combobox
        self.combo_type.set("")
        self.combo_categorie.set("TCO")
        self._set_generated_id()

        # Clear text area
        self.text_description.delete(1.0, tk.END)
