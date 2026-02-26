"""
Paramétrage form GUI component
"""

import tkinter as tk
from tkinter import messagebox, ttk

from config import (
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_RED,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.excel_handler import (
    get_parametrage_headers,
    load_all_parametrages,
    save_parametrage_to_excel,
    update_parametrage_in_excel,
    delete_parametrage_from_excel,
)
from utils.logger import logger


class ParametrageForm:
    """Form for application parameters with two fields: PARAMETRE and VALEUR."""

    DEFAULT_HEADERS = ["PARAMETRE", "VALEUR"]
    DEFAULT_PARAMETERS = ["Prix Essence", "Prix Gasoil"]

    def __init__(self, parent, edit_data=None, row_number=None, callback_on_save=None):
        self.parent = parent
        self.edit_data = edit_data
        self.row_number = row_number
        self.callback_on_save = callback_on_save

        self.headers = []
        self.field_vars = {}
        self.field_widgets = {}

        self._load_headers()
        self._create_form()
        if self.edit_data:
            self._load_edit_data()

    def _load_headers(self):
        headers = get_parametrage_headers()
        self.headers = headers if headers else self.DEFAULT_HEADERS.copy()

        for required in self.DEFAULT_HEADERS:
            if required not in self.headers:
                self.headers.append(required)

    def _load_parameter_choices(self):
        values = set(self.DEFAULT_PARAMETERS)
        try:
            for row in load_all_parametrages():
                label = str(row.get("PARAMETRE") or "").strip()
                if label:
                    values.add(label)
        except Exception as e:
            logger.error(f"Error loading parameter choices: {e}", exc_info=True)
        return [""] + sorted(values, key=lambda v: v.lower())

    def _create_form(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        title_text = "MODIFIER PARAMÉTRAGE" if self.edit_data else "PARAMÉTRAGE APPLICATION"
        tk.Label(
            self.parent,
            text=title_text,
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(20, 10), fill="x")

        tk.Label(
            self.parent,
            text="Définissez les paramètres globaux (ex: prix essence, prix gasoil).",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 10))

        form_frame = tk.LabelFrame(
            self.parent,
            text="Informations paramètre",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        form_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.field_vars = {}
        self.field_widgets = {}

        for row, header in enumerate(self.DEFAULT_HEADERS):
            tk.Label(
                form_frame,
                text=f"{header} :",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=6)

            field_var = tk.StringVar()
            if header == "PARAMETRE":
                widget = ttk.Combobox(
                    form_frame,
                    textvariable=field_var,
                    values=self._load_parameter_choices(),
                    font=ENTRY_FONT,
                    width=38,
                    state="normal",
                )
            else:
                widget = tk.Entry(
                    form_frame,
                    textvariable=field_var,
                    font=ENTRY_FONT,
                    width=40,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                )

            widget.grid(row=row, column=1, sticky="w", padx=(0, 12), pady=6)
            self.field_vars[header] = field_var
            self.field_widgets[header] = widget

        button_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        button_frame.pack(fill="x", padx=20, pady=(0, 12))

        tk.Button(
            button_frame,
            text="💾 Enregistrer",
            command=self._save,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            button_frame,
            text="🧹 Réinitialiser",
            command=self._clear,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=5,
        ).pack(side="left", padx=(0, 8))

        if self.edit_data:
            tk.Button(
                button_frame,
                text="🗑️ Supprimer",
                command=self._delete,
                bg=BUTTON_RED,
                fg="white",
                font=BUTTON_FONT,
                padx=12,
                pady=5,
            ).pack(side="left", padx=(0, 8))

            tk.Button(
                button_frame,
                text="↩ Annuler",
                command=self._cancel,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
                padx=12,
                pady=5,
            ).pack(side="left")

    def _load_edit_data(self):
        try:
            self.field_vars["PARAMETRE"].set(str(self.edit_data.get("PARAMETRE", "")).strip())
            self.field_vars["VALEUR"].set(str(self.edit_data.get("VALEUR", "")).strip())
        except Exception as e:
            logger.error(f"Error loading parameter edit data: {e}", exc_info=True)

    def _collect_form_data(self):
        return {
            "PARAMETRE": self.field_vars["PARAMETRE"].get().strip(),
            "VALEUR": self.field_vars["VALEUR"].get().strip(),
        }

    def _save(self):
        form_data = self._collect_form_data()

        if not form_data["PARAMETRE"]:
            messagebox.showwarning("Validation", "Le champ PARAMETRE est obligatoire.")
            return
        if not form_data["VALEUR"]:
            messagebox.showwarning("Validation", "Le champ VALEUR est obligatoire.")
            return

        if self.edit_data and self.row_number is not None:
            result = update_parametrage_in_excel(self.row_number, form_data)
            if result == -2:
                messagebox.showerror("Fichier verrouillé", "Fermez data-hotel.xlsx puis réessayez.")
                return
            if result == -1:
                messagebox.showerror("Erreur", "Échec de la mise à jour du paramètre.")
                return
            messagebox.showinfo("Succès", "Paramètre modifié avec succès.")
            if self.callback_on_save:
                self.callback_on_save()
            return

        row = save_parametrage_to_excel(form_data)
        if row == -2:
            messagebox.showerror("Fichier verrouillé", "Fermez data-hotel.xlsx puis réessayez.")
            return
        if row == -1:
            messagebox.showerror("Erreur", "Échec de l'enregistrement du paramètre.")
            return

        messagebox.showinfo("Succès", f"Paramètre enregistré à la ligne {row}.")
        if self.callback_on_save:
            self.callback_on_save()
        else:
            self._clear()

    def _delete(self):
        if self.row_number is None:
            return

        confirm = messagebox.askyesno(
            "Confirmation",
            "Êtes-vous sûr de vouloir supprimer ce paramètre ?\n\nCette action ne peut pas être annulée.",
        )
        if not confirm:
            return

        success = delete_parametrage_from_excel(self.row_number)
        if not success:
            messagebox.showerror("Erreur", "Impossible de supprimer. Vérifiez que data-hotel.xlsx n'est pas ouvert.")
            return

        messagebox.showinfo("Succès", "Paramètre supprimé avec succès.")
        if self.callback_on_save:
            self.callback_on_save()

    def _cancel(self):
        if self.callback_on_save:
            self.callback_on_save()

    def _clear(self):
        self.field_vars["PARAMETRE"].set("")
        self.field_vars["VALEUR"].set("")
