"""Collective expense DB form (data-hotel.xlsx / Frais collectifs)."""

import tkinter as tk
from tkinter import messagebox

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
    save_collective_expense_db_row,
    update_collective_expense_db_row,
)


class CollectiveExpenseDBForm:
    def __init__(self, parent, edit_data=None, row_number=None, callback_on_done=None):
        self.parent = parent
        self.edit_data = edit_data
        self.row_number = row_number
        self.callback_on_done = callback_on_done
        self.vars = {}
        self._create_form()
        if self.edit_data:
            self._load_edit_data()

    def _create_form(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        tk.Label(
            self.parent,
            text="MODIFIER FRAIS COLLECTIFS (DB)" if self.edit_data else "FRAIS COLLECTIFS (DB)",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(20, 10))

        frame = tk.LabelFrame(
            self.parent,
            text="Données source (data-hotel.xlsx / Frais collectifs)",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=12,
            pady=12,
        )
        frame.pack(fill="x", padx=20, pady=(0, 10))

        fields = [
            ("forfait", "FORFAIT"),
            ("prestataire", "PRESTATAIRES"),
            ("designation", "DESIGNATION"),
            ("montant", "MONTANT"),
            ("id_circuit", "ID circuit"),
        ]

        for idx, (key, label) in enumerate(fields):
            tk.Label(
                frame,
                text=f"{label} :",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=idx, column=0, sticky="w", padx=(0, 8), pady=6)

            var = tk.StringVar()
            entry = tk.Entry(
                frame,
                textvariable=var,
                font=ENTRY_FONT,
                width=45,
                bg=INPUT_BG_COLOR,
                fg=TEXT_COLOR,
            )
            entry.grid(row=idx, column=1, sticky="we", pady=6)
            self.vars[key] = var

        frame.grid_columnconfigure(1, weight=1)

        btn_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))

        tk.Button(
            btn_frame,
            text="💾 Modifications" if self.edit_data else "💾 Enregistrer",
            command=self._save,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=16,
            pady=6,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btn_frame,
            text="❌ Annuler",
            command=self._cancel,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=16,
            pady=6,
        ).pack(side="left", padx=(0, 8))

        if not self.edit_data:
            tk.Button(
                btn_frame,
                text="🧹 Réinitialiser",
                command=self._clear,
                bg=BUTTON_RED,
                fg="white",
                font=BUTTON_FONT,
                padx=16,
                pady=6,
            ).pack(side="left")

    def _load_edit_data(self):
        for key, var in self.vars.items():
            var.set(str(self.edit_data.get(key, "")))

    def _collect_data(self):
        return {key: var.get().strip() for key, var in self.vars.items()}

    def _save(self):
        data = self._collect_data()
        if not any(data.values()):
            messagebox.showwarning("Validation", "Veuillez renseigner au moins un champ.")
            return

        if self.edit_data and self.row_number is not None:
            result = update_collective_expense_db_row(self.row_number, data)
            if result == -2:
                messagebox.showerror("Fichier verrouillé", "Fermez data-hotel.xlsx puis réessayez.")
                return
            if result == -1:
                messagebox.showerror("Erreur", "Échec de la mise à jour dans Frais collectifs.")
                return
            messagebox.showinfo("Succès", "Ligne mise à jour avec succès.")
            if self.callback_on_done:
                self.callback_on_done()
            return

        row = save_collective_expense_db_row(data)
        if row == -2:
            messagebox.showerror("Fichier verrouillé", "Fermez data-hotel.xlsx puis réessayez.")
            return
        if row == -1:
            messagebox.showerror("Erreur", "Échec de l'enregistrement dans Frais collectifs.")
            return

        messagebox.showinfo("Succès", f"Ligne enregistrée à la ligne {row}.")
        self._clear()

    def _cancel(self):
        if self.callback_on_done:
            self.callback_on_done()

    def _clear(self):
        for var in self.vars.values():
            var.set("")
