"""
Paramétrage summary GUI component
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
from utils.excel_handler import delete_parametrage_from_excel, load_all_parametrages
from utils.logger import logger


class ParametrageSummary:
    """Display list/summary for application parameters."""

    def __init__(self, parent, callback_edit=None, callback_add=None):
        self.parent = parent
        self.callback_edit = callback_edit
        self.callback_add = callback_add
        self.rows = []
        self.search_var = tk.StringVar()
        self.tree = None

        self._load_rows()
        self._create_interface()

    def _load_rows(self):
        try:
            self.rows = load_all_parametrages()
        except Exception as e:
            logger.error(f"Error loading parameters: {e}", exc_info=True)
            self.rows = []

    def _normalize(self, value):
        return str(value or "").strip().lower()

    def _create_interface(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        tk.Label(
            self.parent,
            text="RÉSUMÉ PARAMÉTRAGE",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(20, 10), fill="x")

        self.main_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        top_controls = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        top_controls.pack(fill="x", padx=8, pady=(0, 10))

        tk.Label(
            top_controls,
            text="Recherche:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        search_entry = tk.Entry(
            top_controls,
            textvariable=self.search_var,
            font=ENTRY_FONT,
            width=32,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        )
        search_entry.pack(side="left", padx=(8, 16))
        self.search_var.trace("w", self._refresh_display)

        tk.Button(
            top_controls,
            text="🔄 Rafraîchir",
            command=self._on_refresh_clicked,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left")

        action_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        action_frame.pack(fill="x", padx=8, pady=(0, 10))

        tk.Button(
            action_frame,
            text="➕ Ajouter",
            command=self._on_add_clicked,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            action_frame,
            text="✏️ Modifier",
            command=self._on_edit_clicked,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            action_frame,
            text="🗑️ Supprimer",
            command=self._on_delete_clicked,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left")

        self.content_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        self.content_frame.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure(
            "Treeview",
            background=INPUT_BG_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR,
        )
        style.map("Treeview", background=[("selected", BUTTON_GREEN)])

        self._refresh_display()

    def _on_refresh_clicked(self):
        self._load_rows()
        self._refresh_display()

    def _on_add_clicked(self):
        if self.callback_add:
            self.callback_add()

    def _on_edit_clicked(self):
        if self.tree is None:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à modifier.")
            return

        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à modifier.")
            return

        selected_item = selection[0]
        try:
            row_number = int(selected_item)
        except (TypeError, ValueError):
            return

        selected_row = next((row for row in self.rows if row.get("row_number") == row_number), None)
        if selected_row is None:
            return

        if self.callback_edit:
            self.callback_edit(selected_row, row_number)

    def _on_delete_clicked(self):
        if self.tree is None:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à supprimer.")
            return

        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à supprimer.")
            return

        confirm = messagebox.askyesno(
            "Confirmation",
            "Êtes-vous sûr de vouloir supprimer ce paramètre ?\n\nCette action ne peut pas être annulée.",
        )
        if not confirm:
            return

        selected_item = selection[0]
        try:
            row_number = int(selected_item)
        except (TypeError, ValueError):
            return

        success = delete_parametrage_from_excel(row_number)
        if success:
            messagebox.showinfo("Succès", "Paramètre supprimé avec succès.")
            self._on_refresh_clicked()
        else:
            messagebox.showerror("Erreur", "Impossible de supprimer. Vérifiez que data-hotel.xlsx n'est pas ouvert.")

    def _filtered_rows(self):
        query = self._normalize(self.search_var.get())
        if not query:
            return self.rows

        filtered = []
        for row in self.rows:
            if query in self._normalize(row.get("PARAMETRE")) or query in self._normalize(row.get("VALEUR")):
                filtered.append(row)
        return filtered

    def _clear_content(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def _refresh_display(self, *args):
        self._clear_content()
        rows = self._filtered_rows()

        if not rows:
            tk.Label(
                self.content_frame,
                text="Aucun paramètre trouvé",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            self.tree = None
            return

        headers = ["PARAMETRE", "VALEUR"]
        self.tree = ttk.Treeview(self.content_frame, columns=headers, show="headings", selectmode="browse")

        for header in headers:
            self.tree.heading(header, text=header)
            self.tree.column(header, width=280, anchor="w")

        scrollbar_y = ttk.Scrollbar(self.content_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(self.content_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        for row in rows:
            values = [row.get("PARAMETRE", ""), row.get("VALEUR", "")]
            iid = str(row.get("row_number") or "")
            if not iid:
                iid = None
            self.tree.insert("", "end", iid=iid, values=values)

        self.tree.pack(fill="both", expand=True, side="left")
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
