"""
Client form GUI component - Version améliorée avec nouveaux champs
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from config import *
from models.client_data import ClientData
from utils.validators import validate_email, validate_phone_number
from utils.excel_handler import save_client_to_excel
from utils.logger import logger
import calendar


class CalendarDialog(tk.Toplevel):
    """
    Custom calendar dialog for date selection
    """
    
    def __init__(self, parent, title="Choisir une date"):
        super().__init__(parent)
        self.title(title)
        self.geometry("350x400")
        self.configure(bg=MAIN_BG_COLOR)
        self.resizable(False, False)
        
        # Center window
        self.transient(parent)
        self.grab_set()
        
        self.selected_date = None
        self.current_month = datetime.now().month
        self.current_year = datetime.now().year
        
        self._create_widgets()
        
    def _create_widgets(self):
        """Create calendar widgets"""
        # Header frame
        header_frame = tk.Frame(self, bg=MAIN_BG_COLOR)
        header_frame.pack(fill="x", padx=10, pady=10)
        
        # Previous month button
        tk.Button(
            header_frame,
            text="◀",
            command=self._prev_month,
            bg="#3498db",
            fg="white",
            font=("Arial", 12, "bold"),
            width=3
        ).pack(side="left")
        
        # Month/Year label
        self.month_year_label = tk.Label(
            header_frame,
            text="",
            font=("Arial", 14, "bold"),
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.month_year_label.pack(side="left", expand=True)
        
        # Next month button
        tk.Button(
            header_frame,
            text="▶",
            command=self._next_month,
            bg="#3498db",
            fg="white",
            font=("Arial", 12, "bold"),
            width=3
        ).pack(side="right")
        
        # Calendar frame
        self.calendar_frame = tk.Frame(self, bg=MAIN_BG_COLOR)
        self.calendar_frame.pack(padx=10, pady=10)
        
        self._show_calendar()
        
        # Button frame
        btn_frame = tk.Frame(self, bg=MAIN_BG_COLOR)
        btn_frame.pack(pady=10)
        
        tk.Button(
            btn_frame,
            text="Aujourd'hui",
            command=self._select_today,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10),
            width=12
        ).pack(side="left", padx=5)
        
        tk.Button(
            btn_frame,
            text="Annuler",
            command=self.destroy,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10),
            width=12
        ).pack(side="left", padx=5)
    
    def _show_calendar(self):
        """Display calendar for current month/year"""
        # Clear previous calendar
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        # Update month/year label
        months_fr = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                     "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        self.month_year_label.config(
            text=f"{months_fr[self.current_month - 1]} {self.current_year}"
        )
        
        # Day headers
        days = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]
        for i, day in enumerate(days):
            tk.Label(
                self.calendar_frame,
                text=day,
                font=("Arial", 10, "bold"),
                bg="#3498db",
                fg="white",
                width=5
            ).grid(row=0, column=i, padx=1, pady=1, sticky="nsew")
        
        # Get calendar data
        cal = calendar.monthcalendar(self.current_year, self.current_month)
        
        # Today's date
        today = datetime.now()
        
        # Create day buttons
        for week_num, week in enumerate(cal, start=1):
            for day_num, day in enumerate(week):
                if day == 0:
                    # Empty cell
                    tk.Label(
                        self.calendar_frame,
                        text="",
                        bg=MAIN_BG_COLOR,
                        width=5
                    ).grid(row=week_num, column=day_num, padx=1, pady=1)
                else:
                    # Check if it's today
                    is_today = (day == today.day and 
                               self.current_month == today.month and 
                               self.current_year == today.year)
                    
                    bg_color = "#27ae60" if is_today else "#ecf0f1"
                    fg_color = "white" if is_today else "black"
                    
                    btn = tk.Button(
                        self.calendar_frame,
                        text=str(day),
                        bg=bg_color,
                        fg=fg_color,
                        font=("Arial", 10),
                        width=5,
                        command=lambda d=day: self._select_date(d)
                    )
                    btn.grid(row=week_num, column=day_num, padx=1, pady=1, sticky="nsew")
    
    def _prev_month(self):
        """Go to previous month"""
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self._show_calendar()
    
    def _next_month(self):
        """Go to next month"""
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self._show_calendar()
    
    def _select_date(self, day):
        """Select a specific date"""
        self.selected_date = datetime(self.current_year, self.current_month, day)
        self.destroy()
    
    def _select_today(self):
        """Select today's date"""
        today = datetime.now()
        self.selected_date = today
        self.destroy()


class ClientForm:
    """
    Client request form component with extended fields
    """

    def __init__(self, parent, client_to_edit=None, on_save_callback=None):
        """
        Initialize client form

        Args:
            parent: Parent widget
            client_to_edit: Optional client dict to edit
            on_save_callback: Optional callback after saving
        """
        self.parent = parent
        self.client_to_edit = client_to_edit
        self.on_save_callback = on_save_callback
        self.client_data = {}
        self.canvas = None
        self.form_frame = None
        self.main_frame = None

        self._create_form()

    def _create_form(self):
        """Create the client form with scrollable area for many fields"""
        # Clear parent
        for widget in self.parent.winfo_children():
            widget.destroy()

        # Title - NOT part of scrollable area
        title_text = "MODIFIER CLIENT" if self.client_to_edit else "FORMULAIRE DEMANDE CLIENT"
        title = tk.Label(
            self.parent,
            text=title_text,
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        )
        title.pack(pady=(10, 5), fill="x")

        # Use the parent scrollable frame to occupy full height
        self.main_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        # ===== SECTION: INFOS CLIENTS =====
        section_label = tk.Label(
            self.main_frame,
            text="📋 INFOS CLIENTS",
            font=("Arial", 12, "bold"),
            fg="#2c3e50",
            bg=MAIN_BG_COLOR
        )
        section_label.pack(anchor="w", pady=(15, 10))

        # Date du jour
        tk.Label(
            self.main_frame,
            text="Date du jour",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_date_jour = tk.Entry(
            self.main_frame,
            font=ENTRY_FONT,
            width=40,
            bg="#e8f5e9",
            fg=TEXT_COLOR,
            state="readonly"
        )
        self.entry_date_jour.pack(fill="x", pady=(0, 15))
        current_date = datetime.now().strftime("%d/%m/%Y")
        self.entry_date_jour.config(state="normal")
        self.entry_date_jour.insert(0, current_date)
        self.entry_date_jour.config(state="readonly")

        # Référence client
        tk.Label(
            self.main_frame,
            text="Référence client *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_ref_client = tk.Entry(
            self.main_frame,
            font=ENTRY_FONT,
            width=40,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_ref_client.pack(fill="x", pady=(0, 15))

        # Type de client
        tk.Label(
            self.main_frame,
            text="Type de client *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_type_client = ttk.Combobox(
            self.main_frame,
            values=["Individuel", "Groupe"],
            state="readonly",
            width=37
        )
        self.combo_type_client.pack(fill="x", pady=(0, 15))

        # Nom et Prénom (séparés)
        tk.Label(
            self.main_frame,
            text="Prénom *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_prenom = tk.Entry(
            self.main_frame,
            font=ENTRY_FONT,
            width=40,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_prenom.pack(fill="x", pady=(0, 15))

        tk.Label(
            self.main_frame,
            text="Nom *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_nom = tk.Entry(
            self.main_frame,
            font=ENTRY_FONT,
            width=40,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_nom.pack(fill="x", pady=(0, 15))

        # Dates d'arrivée et départ
        dates_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        dates_frame.pack(fill="x", pady=(0, 15))
        
        left_frame = tk.Frame(dates_frame, bg=MAIN_BG_COLOR)
        left_frame.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        tk.Label(
            left_frame,
            text="Date d'arrivée *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        arrival_frame = tk.Frame(left_frame, bg=MAIN_BG_COLOR)
        arrival_frame.pack(fill="x")
        self.entry_date_arrivee = tk.Entry(
            arrival_frame,
            font=ENTRY_FONT,
            width=15,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
            state="readonly"
        )
        self.entry_date_arrivee.pack(side="left")
        tk.Button(
            arrival_frame,
            text="📅",
            font=("Arial", 12),
            width=3,
            bg="#27ae60",
            fg="white",
            command=lambda: self._open_calendar(self.entry_date_arrivee)
        ).pack(side="left", padx=(5, 0))

        right_frame = tk.Frame(dates_frame, bg=MAIN_BG_COLOR)
        right_frame.pack(side="right", fill="x", expand=True)
        
        tk.Label(
            right_frame,
            text="Date de départ *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        departure_frame = tk.Frame(right_frame, bg=MAIN_BG_COLOR)
        departure_frame.pack(fill="x")
        self.entry_date_depart = tk.Entry(
            departure_frame,
            font=ENTRY_FONT,
            width=15,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
            state="readonly"
        )
        self.entry_date_depart.pack(side="left")
        tk.Button(
            departure_frame,
            text="📅",
            font=("Arial", 12),
            width=3,
            bg="#27ae60",
            fg="white",
            command=lambda: self._open_calendar(self.entry_date_depart)
        ).pack(side="left", padx=(5, 0))

        # Durée du séjour (auto-calculée)
        tk.Label(
            self.main_frame,
            text="Durée du séjour",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_duree_sejour = tk.Entry(
            self.main_frame,
            font=ENTRY_FONT,
            width=40,
            bg="#e8f5e9",
            fg=TEXT_COLOR,
            state="readonly"
        )
        self.entry_duree_sejour.pack(fill="x", pady=(0, 15))

        # Children checkbox (placed before composition section)
        self.var_enfant = tk.BooleanVar()
        self.check_enfant_widget = tk.Checkbutton(
            self.main_frame,
            text="Voyage avec enfant(s)",
            variable=self.var_enfant,
            command=self._toggle_enfant,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            selectcolor="#4CAF50"
        )
        self.check_enfant_widget.pack(anchor="w", pady=(0, 15))

        # Nombre de participants et composition
        tk.Label(
            self.main_frame,
            text="Nombre total de participants *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_nombre_participants = tk.Entry(
            self.main_frame,
            font=ENTRY_FONT,
            width=40,
            bg="#e8f5e9",
            fg=TEXT_COLOR
        )
        self.entry_nombre_participants.pack(fill="x", pady=(0, 15))
        self.entry_nombre_participants.config(state="readonly")

        # Composition par âge
        composition_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        composition_frame.pack(fill="x", pady=(0, 15))

        col1 = tk.Frame(composition_frame, bg=MAIN_BG_COLOR)
        col1.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        tk.Label(
            col1,
            text="Adultes (+ de 12 ans) *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_adultes = tk.Entry(
            col1,
            font=ENTRY_FONT,
            width=15,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_adultes.pack(fill="x")
        self.entry_adultes.bind("<KeyRelease>", lambda e: self._update_participants_total())

        col2 = tk.Frame(composition_frame, bg=MAIN_BG_COLOR)
        col2.pack(side="left", fill="x", expand=True)

        tk.Label(
            col2,
            text="Enfants (2-12 ans)",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_enfants_2_12 = tk.Entry(
            col2,
            font=ENTRY_FONT,
            width=15,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_enfants_2_12.pack(fill="x")
        self.entry_enfants_2_12.bind("<KeyRelease>", lambda e: self._update_participants_total())

        col3 = tk.Frame(composition_frame, bg=MAIN_BG_COLOR)
        col3.pack(side="right", fill="x", expand=True, padx=(10, 0))

        tk.Label(
            col3,
            text="Bébés (0-2 ans)",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_bebes_0_2 = tk.Entry(
            col3,
            font=ENTRY_FONT,
            width=15,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_bebes_0_2.pack(fill="x")
        self.entry_bebes_0_2.bind("<KeyRelease>", lambda e: self._update_participants_total())

        # ===== SECTION: CONTACTS =====
        section_label = tk.Label(
            self.main_frame,
            text="📞 CONTACTS",
            font=("Arial", 12, "bold"),
            fg="#2c3e50",
            bg=MAIN_BG_COLOR
        )
        section_label.pack(anchor="w", pady=(20, 10))

        # Email
        tk.Label(
            self.main_frame,
            text="Adresse email (Obligatoire) *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_email = tk.Entry(
            self.main_frame,
            font=ENTRY_FONT,
            width=40,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_email.pack(fill="x", pady=(0, 15))

        # Téléphone principal
        tk.Label(
            self.main_frame,
            text="Numéro de téléphone *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        phone_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        phone_frame.pack(fill="x", pady=(0, 15))
        self.entry_code_pays = ttk.Combobox(
            phone_frame,
            values=PHONE_CODES,
            width=8,
            state="readonly"
        )
        self.entry_code_pays.set(DEFAULT_PHONE_CODE)
        self.entry_code_pays.pack(side="left")
        self.entry_telephone = tk.Entry(
            phone_frame,
            font=ENTRY_FONT,
            width=25,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_telephone.pack(side="left", padx=(10, 0))

        # WhatsApp
        tk.Label(
            self.main_frame,
            text="Numéro WhatsApp",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        whatsapp_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        whatsapp_frame.pack(fill="x", pady=(0, 15))
        self.entry_code_whatsapp = ttk.Combobox(
            whatsapp_frame,
            values=PHONE_CODES,
            width=8,
            state="readonly"
        )
        self.entry_code_whatsapp.set(DEFAULT_PHONE_CODE)
        self.entry_code_whatsapp.pack(side="left")
        self.entry_telephone_whatsapp = tk.Entry(
            whatsapp_frame,
            font=ENTRY_FONT,
            width=25,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_telephone_whatsapp.pack(side="left", padx=(10, 0))

        # ===== SECTION: ROOMING LIST =====
        section_label = tk.Label(
            self.main_frame,
            text="🛏️ ROOMING LIST (Répartition par chambre)",
            font=("Arial", 12, "bold"),
            fg="#2c3e50",
            bg=MAIN_BG_COLOR
        )
        section_label.pack(anchor="w", pady=(20, 10))

        rooming_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        rooming_frame.pack(fill="x", pady=(0, 15))

        room_types = [
            ("Single (SGL)", "sgl"),
            ("Double (DBL)", "dbl"),
            ("Twin (TWN)", "twn"),
            ("Triple (TPL)", "tpl"),
            ("Familiale (FML)", "fml")
        ]

        self.rooming_vars = {}
        self.rooming_entries = {}
        
        for label_text, key in room_types:
            col = tk.Frame(rooming_frame, bg=MAIN_BG_COLOR)
            col.pack(side="left", fill="x", expand=True, padx=(0, 10))
            
            # Checkbox pour sélectionner le type de chambre
            var = tk.BooleanVar()
            self.rooming_vars[key] = var
            
            checkbox = tk.Checkbutton(
                col,
                text=label_text,
                variable=var,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
                selectcolor="#4CAF50",
                font=LABEL_FONT,
                command=lambda v=var, k=key: self._on_room_toggle(k)
            )
            checkbox.pack(anchor="w")
            
            # Nombre de chambres
            count_frame = tk.Frame(col, bg=MAIN_BG_COLOR)
            count_frame.pack(fill="x", pady=(5, 0))
            
            tk.Label(
                count_frame,
                text="Nombre:",
                font=("Arial", 9),
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR
            ).pack(side="left")
            
            entry = tk.Entry(
                count_frame,
                font=ENTRY_FONT,
                width=6,
                bg=INPUT_BG_COLOR,
                fg=TEXT_COLOR,
                state="disabled"
            )
            entry.pack(side="left", padx=(5, 0))
            self.rooming_entries[key] = entry

        # ===== SECTION: AUTRES INFOS =====
        section_label = tk.Label(
            self.main_frame,
            text="ℹ️ AUTRES INFORMATIONS",
            font=("Arial", 12, "bold"),
            fg="#2c3e50",
            bg=MAIN_BG_COLOR
        )
        section_label.pack(anchor="w", pady=(20, 10))

        # Période
        tk.Label(
            self.main_frame,
            text="Période de voyage *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_periode = ttk.Combobox(
            self.main_frame,
            values=PERIODES,
            state="readonly",
            width=37
        )
        self.combo_periode.pack(fill="x", pady=(0, 15))

        # Restauration
        tk.Label(
            self.main_frame,
            text="Restauration *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_restauration = ttk.Combobox(
            self.main_frame,
            values=RESTAURATIONS,
            state="readonly",
            width=37
        )
        self.combo_restauration.pack(fill="x", pady=(0, 15))

        # Accommodation
        tk.Label(
            self.main_frame,
            text="Type d'hébergement *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_TypeHebergement = ttk.Combobox(
            self.main_frame,
            values=TYPE_HEBERGEMENTS,
            state="readonly",
            width=37
        )
        self.combo_TypeHebergement.pack(fill="x", pady=(0, 15))

        # Room type
        tk.Label(
            self.main_frame,
            text="Type de chambre *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_TypeChambre = ttk.Combobox(
            self.main_frame,
            values=TYPE_CHAMBRES,
            state="readonly",
            width=37
        )
        self.combo_TypeChambre.pack(fill="x", pady=(0, 15))

        # Package type
        tk.Label(
            self.main_frame,
            text="Type de forfait *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_forfait = ttk.Combobox(
            self.main_frame,
            values=FORFAITS,
            state="readonly",
            width=37
        )
        self.combo_forfait.pack(fill="x", pady=(0, 15))

        # Circuit type
        tk.Label(
            self.main_frame,
            text="Type de circuit *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_circuit = ttk.Combobox(
            self.main_frame,
            values=CIRCUITS,
            state="readonly",
            width=37
        )
        self.combo_circuit.pack(fill="x", pady=(0, 30))

        # Buttons
        btn_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        btn_frame.pack()
        
        if self.client_to_edit:
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
                text="VALIDER",
                command=self._validate,
                bg="#27ae60",
                fg="white",
                font=BUTTON_FONT,
                width=12
            ).pack(side="left", padx=5)

        # Populate fields if editing
        if self.client_to_edit:
            self._populate_fields()
        else:
            self._toggle_enfant()

    def _populate_fields(self):
        """Populate form fields with client data"""
        if not self.client_to_edit:
            return

        # Split telephone into code and number
        telephone = self.client_to_edit.get('telephone', '')
        code_pays = ""
        numero = ""
        if telephone.startswith('+'):
            for code in PHONE_CODES:
                if telephone.startswith(code):
                    code_pays = code
                    numero = telephone[len(code):]
                    break
        else:
            numero = telephone

        # Populate basic fields
        self.entry_ref_client.insert(0, self.client_to_edit.get('ref_client', ''))
        self.entry_prenom.insert(0, self.client_to_edit.get('prenom', ''))
        self.entry_nom.insert(0, self.client_to_edit.get('nom', ''))
        self.entry_date_arrivee.config(state="normal")
        self.entry_date_arrivee.insert(0, self.client_to_edit.get('date_arrivee', ''))
        self.entry_date_arrivee.config(state="readonly")
        self.entry_date_depart.config(state="normal")
        self.entry_date_depart.insert(0, self.client_to_edit.get('date_depart', ''))
        self.entry_date_depart.config(state="readonly")
        
        self.entry_nombre_participants.insert(0, self.client_to_edit.get('nombre_participants', ''))
        self.entry_adultes.insert(0, self.client_to_edit.get('nombre_adultes', ''))
        self.entry_enfants_2_12.insert(0, self.client_to_edit.get('nombre_enfants_2_12', ''))
        self.entry_bebes_0_2.insert(0, self.client_to_edit.get('nombre_bebes_0_2', ''))
        
        self.entry_email.insert(0, self.client_to_edit.get('email', ''))
        self.entry_code_pays.set(code_pays)
        self.entry_telephone.insert(0, numero)
        
        # WhatsApp
        telephone_whatsapp = self.client_to_edit.get('telephone_whatsapp', '')
        code_whatsapp = ""
        numero_whatsapp = ""
        if telephone_whatsapp.startswith('+'):
            for code in PHONE_CODES:
                if telephone_whatsapp.startswith(code):
                    code_whatsapp = code
                    numero_whatsapp = telephone_whatsapp[len(code):]
                    break
        else:
            numero_whatsapp = telephone_whatsapp
        
        self.entry_code_whatsapp.set(code_whatsapp)
        self.entry_telephone_whatsapp.insert(0, numero_whatsapp)
        
        # Rooming list - populate with checkboxes and counts
        room_keys = ['sgl', 'dbl', 'twn', 'tpl', 'fml']
        for key in room_keys:
            count = self.client_to_edit.get(f'{key}_count', '')
            if count:
                # Enable checkbox
                self.rooming_vars[key].set(True)
                self.rooming_entries[key].config(state="normal")
                self.rooming_entries[key].insert(0, count)
            else:
                # Keep disabled
                self.rooming_vars[key].set(False)
                self.rooming_entries[key].config(state="disabled")
        
        # Combo boxes
        self.combo_type_client.set(self.client_to_edit.get('type_client', ''))
        self.combo_periode.set(self.client_to_edit.get('periode', ''))
        self.combo_restauration.set(self.client_to_edit.get('restauration', ''))
        self.combo_TypeHebergement.set(self.client_to_edit.get('hebergement', ''))
        self.combo_TypeChambre.set(self.client_to_edit.get('chambre', ''))
        self.combo_forfait.set(self.client_to_edit.get('forfait', ''))
        self.combo_circuit.set(self.client_to_edit.get('circuit', ''))

        # Handle children
        enfant = self.client_to_edit.get('enfant', '')
        if enfant.lower() == 'oui':
            self.var_enfant.set(True)
            self._toggle_enfant()
        else:
            self.var_enfant.set(False)
            self._toggle_enfant()

        self._update_participants_total()

    def _open_calendar(self, entry_widget):
        """Open calendar dialog for date selection"""
        cal_dialog = CalendarDialog(self.parent, "Choisir une date")
        self.parent.wait_window(cal_dialog)
        
        if cal_dialog.selected_date:
            entry_widget.config(state="normal")
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, cal_dialog.selected_date.strftime("%d/%m/%Y"))
            entry_widget.config(state="readonly")
            # Auto-calculate duration if both dates are set
            self._calculate_duration()

    def _calculate_duration(self):
        """Calculate stay duration based on arrival and departure dates"""
        try:
            date_arr = self.entry_date_arrivee.get()
            date_dep = self.entry_date_depart.get()
            
            if date_arr and date_dep:
                arrival = datetime.strptime(date_arr, "%d/%m/%Y")
                departure = datetime.strptime(date_dep, "%d/%m/%Y")
                duration = (departure - arrival).days
                
                self.entry_duree_sejour.config(state="normal")
                self.entry_duree_sejour.delete(0, tk.END)
                self.entry_duree_sejour.insert(0, f"{duration} jours")
                self.entry_duree_sejour.config(state="readonly")
        except:
            pass

    def _toggle_enfant(self):
        """Show/hide child age field"""
        if self.var_enfant.get():
            self.entry_enfants_2_12.config(state="normal")
            self.entry_bebes_0_2.config(state="normal")
        else:
            self.entry_enfants_2_12.config(state="normal")
            self.entry_enfants_2_12.delete(0, tk.END)
            self.entry_enfants_2_12.insert(0, "0")
            self.entry_enfants_2_12.config(state="disabled")
            self.entry_bebes_0_2.config(state="normal")
            self.entry_bebes_0_2.delete(0, tk.END)
            self.entry_bebes_0_2.insert(0, "0")
            self.entry_bebes_0_2.config(state="disabled")

        self._update_participants_total()

    def _update_participants_total(self):
        """Auto-update total participants from adults + children + babies."""
        def _to_int(value):
            try:
                return int(str(value).strip())
            except Exception:
                return 0

        adults = _to_int(self.entry_adultes.get())
        enfants = _to_int(self.entry_enfants_2_12.get())
        bebes = _to_int(self.entry_bebes_0_2.get())

        total = adults + enfants + bebes
        self.entry_nombre_participants.config(state="normal")
        self.entry_nombre_participants.delete(0, tk.END)
        self.entry_nombre_participants.insert(0, str(total))
        self.entry_nombre_participants.config(state="readonly")

    def _on_room_toggle(self, room_key):
        """Enable/disable room entry field when checkbox is toggled"""
        if self.rooming_vars[room_key].get():
            # Enable the entry field
            self.rooming_entries[room_key].config(state="normal")
        else:
            # Disable and clear the entry field
            self.rooming_entries[room_key].config(state="normal")
            self.rooming_entries[room_key].delete(0, tk.END)
            self.rooming_entries[room_key].config(state="disabled")

    def _validate(self):
        """Validate and save client data"""
        form_data = {
            'timestamp': datetime.now().strftime('%d/%m/%Y %H:%M'),
            'date_jour': self.entry_date_jour.get(),
            'ref_client': self.entry_ref_client.get().strip(),
            'type_client': self.combo_type_client.get(),
            'prenom': self.entry_prenom.get().strip(),
            'nom': self.entry_nom.get().strip(),
            'date_arrivee': self.entry_date_arrivee.get(),
            'date_depart': self.entry_date_depart.get(),
            'duree_sejour': self.entry_duree_sejour.get(),
            'nombre_participants': self.entry_nombre_participants.get().strip(),
            'nombre_adultes': self.entry_adultes.get().strip(),
            'nombre_enfants_2_12': self.entry_enfants_2_12.get().strip(),
            'nombre_bebes_0_2': self.entry_bebes_0_2.get().strip(),
            'telephone': self.entry_code_pays.get() + self.entry_telephone.get(),
            'telephone_whatsapp': self.entry_code_whatsapp.get() + self.entry_telephone_whatsapp.get(),
            'email': self.entry_email.get().strip(),
            'periode': self.combo_periode.get(),
            'restauration': self.combo_restauration.get(),
            'hebergement': self.combo_TypeHebergement.get(),
            'chambre': self.combo_TypeChambre.get(),
            'enfant': 'Oui' if self.var_enfant.get() else 'Non',
            'age_enfant': '',
            'forfait': self.combo_forfait.get(),
            'circuit': self.combo_circuit.get(),
            'sgl_count': self.rooming_entries['sgl'].get().strip(),
            'dbl_count': self.rooming_entries['dbl'].get().strip(),
            'twn_count': self.rooming_entries['twn'].get().strip(),
            'tpl_count': self.rooming_entries['tpl'].get().strip(),
            'fml_count': self.rooming_entries['fml'].get().strip()
        }

        # Create client data object
        client = ClientData.from_form_data(form_data)

        # Validate
        errors = client.validate()

        # Additional validations
        if not validate_email(client.email):
            errors.append("Email invalide")
        if not validate_phone_number(self.entry_code_pays.get(), self.entry_telephone.get()):
            errors.append("Numéro téléphone invalide")

        if errors:
            messagebox.showerror("❌ Erreur", "\n".join(errors))
            return

        # Save or update to Excel
        try:
            if self.client_to_edit:
                # Update existing client
                from utils.excel_handler import update_client_in_excel
                success = update_client_in_excel(self.client_to_edit['row_number'], client.to_dict())
                if success:
                    messagebox.showinfo("✅ SUCCÈS", f"Client {client.nom} modifié avec succès !")
                    logger.info(f"Client updated: {client.ref_client} - {client.nom}")
                    if self.on_save_callback:
                        self.on_save_callback()
                else:
                    error_msg = "Erreur lors de la modification du client. Voir les logs."
                    messagebox.showerror("❌ Erreur Excel", error_msg)
                    logger.error(f"Failed to update client: {client.ref_client}")
            else:
                # Save new client
                row = save_client_to_excel(client.to_dict())
                if row > 0:
                    messagebox.showinfo("✅ SUCCÈS", f"Client {client.nom} sauvé ligne Excel {row} !")
                    logger.info(f"New client saved: {client.ref_client} - {client.nom} at row {row}")
                    self._reset_form()
                else:
                    error_msg = "Erreur lors de la sauvegarde du client. Voir les logs."
                    messagebox.showerror("❌ Erreur Excel", error_msg)
                    logger.error(f"Failed to save new client: {client.ref_client}")
        except Exception as e:
            error_msg = f"Erreur Excel: {str(e)}\n\nDétails dans les logs de l'application."
            messagebox.showerror("❌ Erreur", error_msg)
            logger.error(f"Exception during client save/update: {e}", exc_info=True)

    def _cancel(self):
        """Cancel editing and return to list"""
        if self.on_save_callback:
            self.on_save_callback()

    def _reset_form(self):
        """Reset all form fields"""
        self.entry_ref_client.delete(0, tk.END)
        self.entry_prenom.delete(0, tk.END)
        self.entry_nom.delete(0, tk.END)
        self.entry_date_arrivee.config(state="normal")
        self.entry_date_arrivee.delete(0, tk.END)
        self.entry_date_arrivee.config(state="readonly")
        self.entry_date_depart.config(state="normal")
        self.entry_date_depart.delete(0, tk.END)
        self.entry_date_depart.config(state="readonly")
        self.entry_duree_sejour.config(state="normal")
        self.entry_duree_sejour.delete(0, tk.END)
        self.entry_duree_sejour.config(state="readonly")
        self.entry_nombre_participants.config(state="normal")
        self.entry_nombre_participants.delete(0, tk.END)
        self.entry_nombre_participants.config(state="readonly")
        self.entry_adultes.delete(0, tk.END)
        self.entry_enfants_2_12.delete(0, tk.END)
        self.entry_bebes_0_2.delete(0, tk.END)
        self.entry_telephone.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_telephone_whatsapp.delete(0, tk.END)
        
        # Reset rooming list - uncheck all and disable fields
        for key in self.rooming_entries.keys():
            self.rooming_vars[key].set(False)
            self.rooming_entries[key].config(state="normal")
            self.rooming_entries[key].delete(0, tk.END)
            self.rooming_entries[key].config(state="disabled")
        
        self.combo_type_client.set("")
        self.combo_periode.set("")
        self.combo_restauration.set("")
        self.combo_TypeHebergement.set("")
        self.combo_TypeChambre.set("")
        self.combo_forfait.set("")
        self.combo_circuit.set("")
        self.var_enfant.set(False)
        self._toggle_enfant()
