"""
Client form GUI component - Version améliorée
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from config import *
from models.client_data import ClientData
from utils.validators import validate_email, validate_phone_number
from utils.excel_handler import save_client_to_excel
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
    Client request form component
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
        self.check_enfant_widget = None
        self.canvas = None
        self.form_frame = None

        self._create_form()

    def _create_form(self):
        """Create the client form"""
        # Clear parent
        for widget in self.parent.winfo_children():
            widget.destroy()

        # Title
        title_text = "MODIFIER CLIENT" if self.client_to_edit else "FORMULAIRE DEMANDE CLIENT"
        title = tk.Label(
            self.parent,
            text=title_text,
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        )
        title.pack(pady=(20, 10))

        # Rounded frame
        self.form_frame = self._create_rounded_frame(self.parent, 600, 900, 20)
        main_frame = self.form_frame

        # Date du jour (automatique)
        tk.Label(
            main_frame,
            text="Date du jour",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_date_jour = tk.Entry(
            main_frame,
            font=ENTRY_FONT,
            width=40,
            bg="#e8f5e9",  # Light green background to show it's auto-filled
            fg=TEXT_COLOR,
            state="readonly"
        )
        self.entry_date_jour.pack(fill="x", pady=(0, 15))
        # Auto-fill with current date
        current_date = datetime.now().strftime("%d/%m/%Y")
        self.entry_date_jour.config(state="normal")
        self.entry_date_jour.insert(0, current_date)
        self.entry_date_jour.config(state="readonly")

        # Date reservation
        tk.Label(
            main_frame,
            text="Date réservation *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        date_frame = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
        date_frame.pack(fill="x", pady=(0, 15))
        self.entry_date_reservation = tk.Entry(
            date_frame,
            font=ENTRY_FONT,
            width=25,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
            state="readonly"
        )
        self.entry_date_reservation.pack(side="left")
        tk.Button(
            date_frame,
            text="📅",
            font=("Arial", 12),
            width=3,
            bg="#27ae60",
            fg="white",
            command=self._open_calendar
        ).pack(side="left", padx=(10, 0))

        # Client reference
        tk.Label(
            main_frame,
            text="Référence client *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_ref_client = tk.Entry(
            main_frame,
            font=ENTRY_FONT,
            width=40,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_ref_client.pack(fill="x", pady=(0, 15))

        # Client name
        tk.Label(
            main_frame,
            text="Nom du client *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_nom = tk.Entry(
            main_frame,
            font=ENTRY_FONT,
            width=40,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_nom.pack(fill="x", pady=(0, 15))

        # Phone number
        tk.Label(
            main_frame,
            text="Numéro téléphone *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        phone_frame = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
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

        # Email
        tk.Label(
            main_frame,
            text="Adresse email *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.entry_email = tk.Entry(
            main_frame,
            font=ENTRY_FONT,
            width=40,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        self.entry_email.pack(fill="x", pady=(0, 15))

        # Period
        tk.Label(
            main_frame,
            text="Période de voyage *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_periode = ttk.Combobox(
            main_frame,
            values=PERIODES,
            state="readonly",
            width=37
        )
        self.combo_periode.pack(fill="x", pady=(0, 15))

        # Restauration
        tk.Label(
            main_frame,
            text="Restauration *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_restauration = ttk.Combobox(
            main_frame,
            values=RESTAURATIONS,
            state="readonly",
            width=37
        )
        self.combo_restauration.pack(fill="x", pady=(0, 15))

        # Accommodation
        tk.Label(
            main_frame,
            text="Type d'hébergement *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_TypeHebergement = ttk.Combobox(
            main_frame,
            values=TYPE_HEBERGEMENTS,
            state="readonly",
            width=37
        )
        self.combo_TypeHebergement.pack(fill="x", pady=(0, 15))

        # Room type
        tk.Label(
            main_frame,
            text="Type de chambre *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_TypeChambre = ttk.Combobox(
            main_frame,
            values=TYPE_CHAMBRES,
            state="readonly",
            width=37
        )
        self.combo_TypeChambre.pack(fill="x", pady=(0, 15))

        # Children
        self.var_enfant = tk.BooleanVar()
        self.check_enfant_widget = tk.Checkbutton(
            main_frame,
            text="Voyage avec enfant(s)",
            variable=self.var_enfant,
            command=self._toggle_enfant,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            selectcolor="#4CAF50"
        )
        self.check_enfant_widget.pack(anchor="w", pady=(0, 15))
        self.frame_age_enfant = tk.Frame(main_frame, bg=MAIN_BG_COLOR)

        # Package type
        tk.Label(
            main_frame,
            text="Type de forfait *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_forfait = ttk.Combobox(
            main_frame,
            values=FORFAITS,
            state="readonly",
            width=37
        )
        self.combo_forfait.pack(fill="x", pady=(0, 15))

        # Circuit type
        tk.Label(
            main_frame,
            text="Type de circuit *",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(anchor="w")
        self.combo_circuit = ttk.Combobox(
            main_frame,
            values=CIRCUITS,
            state="readonly",
            width=37
        )
        self.combo_circuit.pack(fill="x", pady=(0, 30))

        # Buttons
        btn_frame = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
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

    def _populate_fields(self):
        """Populate form fields with client data"""
        if not self.client_to_edit:
            return

        # Split telephone into code and number
        telephone = self.client_to_edit.get('telephone', '')
        code_pays = ""
        numero = ""
        if telephone.startswith('+'):
            # Find the country code
            for code in PHONE_CODES:
                if telephone.startswith(code):
                    code_pays = code
                    numero = telephone[len(code):]
                    break
        else:
            numero = telephone

        # Populate fields
        self.entry_ref_client.insert(0, self.client_to_edit.get('ref_client', ''))
        self.entry_nom.insert(0, self.client_to_edit.get('nom', ''))
        self.entry_code_pays.set(code_pays)
        self.entry_telephone.insert(0, numero)
        self.entry_email.insert(0, self.client_to_edit.get('email', ''))
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
            if hasattr(self, 'combo_age_enfant'):
                self.combo_age_enfant.set(self.client_to_edit.get('age_enfant', ''))

    def _create_rounded_frame(self, parent, width, height, radius=20):
        """Create a frame with rounded corners"""
        frame = tk.Frame(parent, bg=MAIN_BG_COLOR)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.canvas = tk.Canvas(
            frame,
            width=width,
            height=height,
            bg=MAIN_BG_COLOR,
            highlightthickness=0
        )
        self.canvas.pack(fill="both", expand=True)

        # Rounded rectangle
        self.canvas.create_oval(10, 10, 10+2*radius, 10+2*radius,
                               fill="#E0E0E0", outline="#D0D0D0", width=2)
        self.canvas.create_oval(width-10-2*radius, 10, width-10, 10+2*radius,
                               fill="#E0E0E0", outline="#D0D0D0", width=2)
        self.canvas.create_oval(width-10-2*radius, height-10-2*radius, width-10, height-10,
                               fill="#E0E0E0", outline="#D0D0D0", width=2)
        self.canvas.create_oval(10, height-10-2*radius, 10+2*radius, height-10,
                               fill="#E0E0E0", outline="#D0D0D0", width=2)

        self.canvas.create_rectangle(radius, 10, width-radius-10, height-10,
                                    fill=MAIN_BG_COLOR, outline="#D0D0D0", width=2)
        self.canvas.create_rectangle(10, radius, width-10, height-radius-10,
                                    fill=MAIN_BG_COLOR, outline="#D0D0D0", width=2)

        # Content frame
        self.form_frame = tk.Frame(frame, bg=MAIN_BG_COLOR)
        self.form_frame.place(x=20, y=20, width=width-40, height=height-40)
        return self.form_frame

    def _open_calendar(self):
        """Open calendar dialog for date selection"""
        cal_dialog = CalendarDialog(self.parent, "Choisir la date de réservation")
        self.parent.wait_window(cal_dialog)
        
        if cal_dialog.selected_date:
            self.entry_date_reservation.config(state="normal")
            self.entry_date_reservation.delete(0, tk.END)
            self.entry_date_reservation.insert(0, cal_dialog.selected_date.strftime("%d/%m/%Y"))
            self.entry_date_reservation.config(state="readonly")

    def _toggle_enfant(self):
        """Show/hide child age field"""
        if self.var_enfant.get():
            self.frame_age_enfant.pack(fill="x", pady=(0, 15), after=self.check_enfant_widget)
            if not hasattr(self, 'combo_age_enfant'):
                tk.Label(
                    self.frame_age_enfant,
                    text="Âge enfant(s):",
                    font=LABEL_FONT,
                    fg=TEXT_COLOR,
                    bg=MAIN_BG_COLOR
                ).pack(anchor="w")
                self.combo_age_enfant = ttk.Combobox(
                    self.frame_age_enfant,
                    values=AGES_ENFANTS,
                    state="readonly",
                    width=37
                )
                self.combo_age_enfant.pack(fill="x")
        else:
            self.frame_age_enfant.pack_forget()

    def _validate(self):
        """Validate and save client data"""
        age_enfant = None
        if self.var_enfant.get() and hasattr(self, 'combo_age_enfant'):
            age_enfant = self.combo_age_enfant.get()

        form_data = {
            'timestamp': datetime.now().strftime('%d/%m/%Y %H:%M'),
            'date_jour': self.entry_date_jour.get(),
            'date_reservation': self.entry_date_reservation.get(),
            'ref_client': self.entry_ref_client.get().strip(),
            'nom': self.entry_nom.get().strip(),
            'telephone': self.entry_code_pays.get() + self.entry_telephone.get(),
            'email': self.entry_email.get().strip(),
            'periode': self.combo_periode.get(),
            'restauration': self.combo_restauration.get(),
            'hebergement': self.combo_TypeHebergement.get(),
            'chambre': self.combo_TypeChambre.get(),
            'enfant': 'Oui' if self.var_enfant.get() else 'Non',
            'age_enfant': age_enfant or '',
            'forfait': self.combo_forfait.get(),
            'circuit': self.combo_circuit.get()
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
        
        # Validate reservation date
        if not self.entry_date_reservation.get():
            errors.append("Date de réservation requise")

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
                    if self.on_save_callback:
                        self.on_save_callback()
                else:
                    messagebox.showerror("❌ Erreur Excel", "Erreur lors de la modification")
            else:
                # Save new client
                row = save_client_to_excel(client.to_dict())
                messagebox.showinfo("✅ SUCCÈS", f"Client {client.nom} sauvé ligne Excel {row} !")
                self._reset_form()
        except Exception as e:
            messagebox.showerror("❌ Erreur Excel", f"Erreur: {str(e)}")

    def _cancel(self):
        """Cancel editing and return to list"""
        if self.on_save_callback:
            self.on_save_callback()

    def _reset_form(self):
        """Reset all form fields"""
        self.entry_ref_client.delete(0, tk.END)
        self.entry_nom.delete(0, tk.END)
        self.entry_telephone.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_date_reservation.config(state="normal")
        self.entry_date_reservation.delete(0, tk.END)
        self.entry_date_reservation.config(state="readonly")
        
        # Reset date du jour to current date
        self.entry_date_jour.config(state="normal")
        self.entry_date_jour.delete(0, tk.END)
        self.entry_date_jour.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.entry_date_jour.config(state="readonly")

        for combo in [self.combo_periode, self.combo_restauration, self.combo_TypeHebergement,
                     self.combo_TypeChambre, self.combo_forfait, self.combo_circuit]:
            combo.set("")
        self.var_enfant.set(False)
        if hasattr(self, 'combo_age_enfant'):
            self.combo_age_enfant.set("")