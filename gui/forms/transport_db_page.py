"""Transport DB management page (TRANSPORT + KM_MADA)."""

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
    delete_km_mada_db_row,
    delete_transport_db_row,
    get_km_mada_db_headers,
    get_transport_db_headers,
    load_km_mada_db_rows,
    load_transport_db_rows,
    save_km_mada_db_row,
    save_transport_db_row,
    update_km_mada_db_row,
    update_transport_db_row,
)


class _SheetCrudPanel:
    def __init__(
        self,
        parent,
        title,
        source_text,
        get_headers,
        load_rows,
        save_row,
        update_row,
        delete_row,
    ):
        self.parent = parent
        self.title = title
        self.source_text = source_text
        self.get_headers_fn = get_headers
        self.load_rows_fn = load_rows
        self.save_row_fn = save_row
        self.update_row_fn = update_row
        self.delete_row_fn = delete_row

        self.headers = []
        self.rows = []
        self.filtered_rows = []
        self.vars = {}
        self.selected_row_number = None
        self.search_var = tk.StringVar()

        self.tree = None
        self.form_frame = None
        self.status_label = None
        self.btn_edit = None
        self.btn_delete = None

        self._build_ui()
        self._load_data()

    def _build_ui(self):
        tk.Label(
            self.parent,
            text=self.title,
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(10, 6))

        tk.Label(
            self.parent,
            text=self.source_text,
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 6))

        search_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        search_frame.pack(fill="x", padx=16, pady=(0, 8))

        tk.Label(
            search_frame,
            text="Rechercher:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        self.search_var.trace("w", self._on_filter_change)
        tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=ENTRY_FONT,
            width=38,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        ).pack(side="left", padx=(8, 0))

        btn_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        btn_frame.pack(fill="x", padx=16, pady=(0, 8))

        tk.Button(
            btn_frame,
            text="🔄 Actualiser",
            command=self._load_data,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
        ).pack(side="left", padx=4)

        tk.Button(
            btn_frame,
            text="➕ Ajouter",
            command=self._new_row,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
        ).pack(side="left", padx=4)

        self.btn_edit = tk.Button(
            btn_frame,
            text="✏️ Modifier",
            command=self._edit_selected,
            bg=BUTTON_ORANGE,
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
        )
        self.btn_edit.pack(side="left", padx=4)

        self.btn_delete = tk.Button(
            btn_frame,
            text="🗑️ Supprimer",
            command=self._delete_selected,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
        )
        self.btn_delete.pack(side="left", padx=4)

        tree_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

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
        self.tree.bind("<Double-1>", lambda _e: self._edit_selected())

        self.status_label = tk.Label(
            self.parent,
            text="",
            font=("Arial", 10),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        self.status_label.pack(anchor="w", padx=16, pady=(0, 6))

        self.form_frame = tk.LabelFrame(
            self.parent,
            text="Formulaire",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        self.form_frame.pack(fill="x", padx=16, pady=(0, 12))

    def _configure_form(self):
        for widget in self.form_frame.winfo_children():
            widget.destroy()

        self.vars = {}
        for idx, header in enumerate(self.headers):
            row = idx // 2
            col_group = idx % 2
            label_col = col_group * 2
            entry_col = label_col + 1

            tk.Label(
                self.form_frame,
                text=f"{header} :",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).grid(row=row, column=label_col, sticky="w", padx=(0, 8), pady=4)

            var = tk.StringVar()
            tk.Entry(
                self.form_frame,
                textvariable=var,
                font=ENTRY_FONT,
                width=35,
                bg=INPUT_BG_COLOR,
                fg=TEXT_COLOR,
            ).grid(row=row, column=entry_col, sticky="we", padx=(0, 10), pady=4)
            self.vars[header] = var

        max_row = (len(self.headers) + 1) // 2
        actions_row = max_row + 1

        btns = tk.Frame(self.form_frame, bg=MAIN_BG_COLOR)
        btns.grid(row=actions_row, column=0, columnspan=4, sticky="w", pady=(8, 0))

        tk.Button(
            btns,
            text="💾 Enregistrer",
            command=self._save_form,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            btns,
            text="❌ Annuler",
            command=self._clear_form,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=12,
            pady=4,
        ).pack(side="left", padx=(0, 8))

        for col in range(4):
            self.form_frame.grid_columnconfigure(col, weight=1 if col % 2 == 1 else 0)

    def _configure_tree_columns(self):
        self.tree.delete(*self.tree.get_children())
        columns = ["row_number"] + self.headers
        self.tree.configure(columns=columns)

        self.tree.heading("row_number", text="N°")
        self.tree.column("row_number", width=70, minwidth=70)

        for header in self.headers:
            self.tree.heading(header, text=header)
            self.tree.column(header, width=150, minwidth=120)

    def _load_data(self):
        self.headers = self.get_headers_fn() or []
        self.rows = self.load_rows_fn() or []
        self._configure_tree_columns()
        self._configure_form()
        self._apply_filters()
        self._update_treeview()
        self._clear_form()

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

    def _on_filter_change(self, *_args):
        self._apply_filters()
        self._update_treeview()

    def _on_selection_change(self, _event=None):
        selected = bool(self.tree.selection())
        self.btn_edit.config(state="normal" if selected else "disabled")
        self.btn_delete.config(state="normal" if selected else "disabled")

    def _new_row(self):
        self._clear_form()

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
        if not row_data:
            return

        self.selected_row_number = row_number
        for header in self.headers:
            self.vars[header].set(str(row_data.get(header, "")))

    def _delete_selected(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Info", "Veuillez sélectionner une ligne à supprimer.")
            return

        confirm = messagebox.askyesno(
            "Confirmation",
            f"Supprimer cette ligne de la base {self.title} ?",
        )
        if not confirm:
            return

        try:
            row_number = int(selection[0])
        except Exception:
            return

        success = self.delete_row_fn(row_number)
        if success:
            messagebox.showinfo("Succès", "Ligne supprimée avec succès.")
            self._load_data()
        else:
            messagebox.showerror(
                "Erreur",
                "Suppression impossible. Vérifiez que data-hotel.xlsx n'est pas ouvert.",
            )

    def _collect_form_data(self):
        return {header: self.vars[header].get().strip() for header in self.headers}

    def _save_form(self):
        if not self.headers:
            messagebox.showerror("Erreur", "Aucun en-tête détecté pour cette feuille.")
            return

        data = self._collect_form_data()
        if not any(data.values()):
            messagebox.showwarning("Validation", "Veuillez renseigner au moins un champ.")
            return

        if self.selected_row_number is not None:
            result = self.update_row_fn(self.selected_row_number, data)
            if result == -2:
                messagebox.showerror("Fichier verrouillé", "Fermez data-hotel.xlsx puis réessayez.")
                return
            if result == -1:
                messagebox.showerror("Erreur", "Échec de la mise à jour.")
                return
            messagebox.showinfo("Succès", "Ligne mise à jour avec succès.")
            self._load_data()
            return

        row = self.save_row_fn(data)
        if row == -2:
            messagebox.showerror("Fichier verrouillé", "Fermez data-hotel.xlsx puis réessayez.")
            return
        if row == -1:
            messagebox.showerror("Erreur", "Échec de l'enregistrement.")
            return

        messagebox.showinfo("Succès", f"Ligne enregistrée à la ligne {row}.")
        self._load_data()

    def _clear_form(self):
        self.selected_row_number = None
        for var in self.vars.values():
            var.set("")
        self.tree.selection_remove(self.tree.selection())
        self._on_selection_change()


class TransportDBPage:
    """Transport DB management page for TRANSPORT and KM_MADA sheets."""

    def __init__(self, parent, on_back_to_db=None):
        self.parent = parent
        self.on_back_to_db = on_back_to_db
        self._create_page()

    def _create_page(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        tk.Label(
            self.parent,
            text="GESTION BASE TRANSPORT (DB)",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(16, 6))

        tk.Label(
            self.parent,
            text="Gestion des feuilles TRANSPORT et KM_MADA (Ajout, Modification, Suppression)",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(pady=(0, 8))

        if self.on_back_to_db:
            tk.Button(
                self.parent,
                text="⬅ Retour vers Bases de données",
                command=self._go_back_to_db,
                bg=BUTTON_BLUE,
                fg="white",
                font=BUTTON_FONT,
            ).pack(anchor="w", padx=16, pady=(0, 8))

        container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        container.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        notebook = ttk.Notebook(container)
        notebook.pack(fill="both", expand=True)

        transport_tab = tk.Frame(notebook, bg=MAIN_BG_COLOR)
        km_mada_tab = tk.Frame(notebook, bg=MAIN_BG_COLOR)

        notebook.add(transport_tab, text="TRANSPORT")
        notebook.add(km_mada_tab, text="KM_MADA")

        _SheetCrudPanel(
            transport_tab,
            title="TRANSPORT (DB)",
            source_text="Source: data-hotel.xlsx / feuille TRANSPORT",
            get_headers=get_transport_db_headers,
            load_rows=load_transport_db_rows,
            save_row=save_transport_db_row,
            update_row=update_transport_db_row,
            delete_row=delete_transport_db_row,
        )

        _SheetCrudPanel(
            km_mada_tab,
            title="KM_MADA (DB)",
            source_text="Source: data-hotel.xlsx / feuille KM_MADA",
            get_headers=get_km_mada_db_headers,
            load_rows=load_km_mada_db_rows,
            save_row=save_km_mada_db_row,
            update_row=update_km_mada_db_row,
            delete_row=delete_km_mada_db_row,
        )

    def _go_back_to_db(self):
        if self.on_back_to_db:
            self.on_back_to_db()
