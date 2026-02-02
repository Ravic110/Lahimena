"""
Client form GUI component
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from config import *
from models.client_data import ClientData
from utils.validators import validate_email, validate_phone_number
from utils.excel_handler import save_client_to_excel


class ClientForm:
    """
    Client request form component
    """

    def __init__(self, parent):
        """
        Initialize client form

        Args:
            parent: Parent widget
        """
        self.parent = parent
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
        title = tk.Label(
            self.parent,
            text="FORMULAIRE DEMANDE CLIENT",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        )
        title.pack(pady=(20, 10))

        # Rounded frame
        self.form_frame = self._create_rounded_frame(self.parent, 600, 850, 20)
        main_frame = self.form_frame

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
        tk.Button(
            btn_frame,
            text="VALIDER",
            command=self._validate,
            bg="#27ae60",
            fg="white",
            font=BUTTON_FONT,
            width=12
        ).pack(side="left", padx=5)
        tk.Button(
            btn_frame,
            text="MODIFIER",
            command=self._modify,
            bg="#f39c12",
            fg="white",
            font=BUTTON_FONT,
            width=12
        ).pack(side="left", padx=5)
        tk.Button(
            btn_frame,
            text="SUPPRIMER",
            command=self._delete,
            bg="#e74c3c",
            fg="white",
            font=BUTTON_FONT,
            width=12
        ).pack(side="left", padx=5)

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
        """Open simple calendar for date selection"""
        cal_window = tk.Toplevel(self.parent)
        cal_window.title("Choisir date")
        cal_window.geometry("300x300")
        cal_window.configure(bg=MAIN_BG_COLOR)

        today = datetime.now()
        self.entry_date_reservation.config(state="normal")
        self.entry_date_reservation.delete(0, tk.END)
        self.entry_date_reservation.insert(0, today.strftime("%d/%m/%Y"))
        self.entry_date_reservation.config(state="readonly")
        cal_window.destroy()

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

        if errors:
            messagebox.showerror("❌ Erreur", "\n".join(errors))
            return

        # Save to Excel
        try:
            row = save_client_to_excel(client.to_dict())
            messagebox.showinfo("✅ SUCCÈS", f"Client {client.nom} sauvé ligne Excel {row} !")
            self._reset_form()
        except Exception as e:
            messagebox.showerror("❌ Erreur Excel", f"Erreur: {str(e)}")

    def _modify(self):
        """Modify client data - placeholder"""
        messagebox.showinfo("Modifier", "Fonction modifier - À implémenter")

    def _delete(self):
        """Delete client data - placeholder"""
        messagebox.showinfo("Supprimer", "Fonction supprimer - À implémenter")

    def _reset_form(self):
        """Reset all form fields"""
        self.entry_ref_client.delete(0, tk.END)
        self.entry_nom.delete(0, tk.END)
        self.entry_telephone.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_date_reservation.config(state="normal")
        self.entry_date_reservation.delete(0, tk.END)
        self.entry_date_reservation.config(state="readonly")

        for combo in [self.combo_periode, self.combo_restauration, self.combo_TypeHebergement,
                     self.combo_TypeChambre, self.combo_forfait, self.combo_circuit]:
            combo.set("")
        self.var_enfant.set(False)
        if hasattr(self, 'combo_age_enfant'):
            self.combo_age_enfant.set("")