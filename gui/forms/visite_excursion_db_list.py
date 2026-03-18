"""Visite & Excursion DB list (data-hotel.xlsx / Visite_excursion)."""

import tkinter as tk
from tkinter import messagebox, ttk

from config import (
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_ORANGE,
    BUTTON_RED,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.excel_handler import (
    delete_visite_excursion_db_row,
    get_visite_excursion_db_headers,
    load_visite_excursion_db_rows,
)


class VisiteExcursionDBList:
    def __init__(self, parent, on_edit_row=None, on_new_row=None):
        self.parent = parent
        self.on_edit_row = on_edit_row
        self.on_new_row = on_new_row
        self.headers = []
        self.rows = []
        self.filtered_rows = []
        self._create_list()

    def _create_list(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        tk.Label(
            self.parent,
            text="LISTE VISITE & EXCURSION (DB)",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(20, 10))

        tk.Label(
            self.parent,
            text="Source: data-hotel.xlsx / feuille Visite_excursion",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 8))

        search_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        search_frame.pack(fill="x", padx=20, pady=(0, 10))

        tk.Label(
            search_frame,
            text="Rechercher:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        self.search_var = tk.StringVar()
        self.search_var.trace("w", self._on_filter_change)
        tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=ENTRY_FONT,
            width=35,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        ).pack(side="left", padx=(8, 0))

        btn_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        btn_frame.pack(fill="x", padx=20, pady=(0, 10))

        tk.Button(
            btn_frame,
            text="🔄 Actualiser",
            command=self._load_rows,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
        ).pack(side="left", padx=5)

        tk.Button(
            btn_frame,
            text="➕ Ajouter",
            command=self._new_row,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
        ).pack(side="left", padx=5)

        self.btn_edit = tk.Button(
            btn_frame,
            text="✏️ Modifier",
            command=self._edit_selected,
            bg=BUTTON_ORANGE,
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
        )
        self.btn_edit.pack(side="left", padx=5)

        self.btn_delete = tk.Button(
            btn_frame,
            text="🗑️ Supprimer",
            command=self._delete_selected,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
        )
        self.btn_delete.pack(side="left", padx=5)

        tree_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal")

        style = ttk.Style()
        style.configure(
            "Treeview",
            background=INPUT_BG_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR,
        )
        style.map("Treeview", background=[("selected", BUTTON_GREEN)])

        self.tree = ttk.Treeview(
            tree_frame,
            columns=(),
            show="headings",
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            style="Treeview",
        )
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)

        self.tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        self.tree.bind("<<TreeviewSelect>>", self._on_selection_change)
        self.tree.bind("<Double-1>", lambda e: self._edit_selected())

        self.status_label = tk.Label(
            self.parent,
            text="",
            font=("Poppins", 10),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        self.status_label.pack(anchor="w", padx=20, pady=(0, 10))

        self._load_rows()

    def _configure_tree_columns(self):
        self.tree.delete(*self.tree.get_children())

        columns = ["row_number"] + self.headers
        self.tree.configure(columns=columns)

        self.tree.heading("row_number", text="N°")
        self.tree.column("row_number", width=60, minwidth=60)

        for header in self.headers:
            self.tree.heading(header, text=header)
            self.tree.column(header, width=160, minwidth=120)

    def _load_rows(self):
        self.headers = get_visite_excursion_db_headers()
        self.rows = load_visite_excursion_db_rows()
        self._configure_tree_columns()
        self._apply_filters()
        self._update_treeview()
        self.tree.selection_remove(self.tree.selection())
        self._on_selection_change()

    def _apply_filters(self):
        query = self.search_var.get().strip().lower()
        if not query:
            self.filtered_rows = list(self.rows)
            return

        filtered = []
        for row in self.rows:
            hay = " ".join(str(row.get(h, "")) for h in self.headers).lower()
            if query in hay:
                filtered.append(row)
        self.filtered_rows = filtered

    def _update_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for row in self.filtered_rows:
            values = [row.get("row_number", "")]
            values.extend(row.get(h, "") for h in self.headers)
            self.tree.insert("", "end", iid=str(row.get("row_number")), values=values)

        self.status_label.config(
            text=f"Affichage de {len(self.filtered_rows)} ligne(s) sur {len(self.rows)}"
        )

    def _on_filter_change(self, *args):
        self._apply_filters()
        self._update_treeview()

    def _on_selection_change(self, event=None):
        selected = bool(self.tree.selection())
        self.btn_edit.config(state="normal" if selected else "disabled")
        self.btn_delete.config(state="normal" if selected else "disabled")

    def _new_row(self):
        if self.on_new_row:
            self.on_new_row()

    def _edit_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à modifier.")
            return

        try:
            row_number = int(selection[0])
        except Exception:
            return

        row_data = next((r for r in self.rows if r.get("row_number") == row_number), None)
        if row_data and self.on_edit_row:
            self.on_edit_row(row_data, row_number)

    def _delete_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à supprimer.")
            return

        confirm = messagebox.askyesno(
            "Confirmation",
            "Supprimer cette ligne de la base Visite_excursion (data-hotel.xlsx) ?",
        )
        if not confirm:
            return

        try:
            row_number = int(selection[0])
        except Exception:
            return

        success = delete_visite_excursion_db_row(row_number)
        if success:
            messagebox.showinfo("Succès", "Ligne supprimée avec succès.")
            self._load_rows()
        else:
            messagebox.showerror("Erreur", "Suppression impossible. Vérifiez que data-hotel.xlsx n'est pas ouvert.")
