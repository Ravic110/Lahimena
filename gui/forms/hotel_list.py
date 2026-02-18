"""
Hotel list GUI component
"""

import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from config import *
from utils.excel_handler import load_all_hotels, delete_hotel_from_excel


class HotelList:
    """
    Hotel list component with search and management features
    """

    def __init__(self, parent, on_edit_hotel=None, on_new_hotel=None):
        """
        Initialize hotel list

        Args:
            parent: Parent widget
            on_edit_hotel: Callback for editing a hotel (receives hotel dict)
            on_new_hotel: Callback for creating a new hotel
        """
        self.parent = parent
        self.on_edit_hotel = on_edit_hotel
        self.on_new_hotel = on_new_hotel
        self.hotels = []
        self.filtered_hotels = []

        self._create_list()

    def _create_list(self):
        """Create the hotel list interface"""
        # Clear parent
        for widget in self.parent.winfo_children():
            widget.destroy()

        # Title
        title = tk.Label(
            self.parent,
            text="LISTE DES HÔTELS",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        )
        title.pack(pady=(20, 10))

        # Search frame
        search_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        search_frame.pack(fill="x", padx=20, pady=(0, 10))

        # Search by name
        tk.Label(
            search_frame,
            text="Rechercher:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(side="left")

        self.search_var = tk.StringVar()
        self.search_var.trace("w", self._on_filter_change)
        search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=ENTRY_FONT,
            width=25,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR
        )
        search_entry.pack(side="left", padx=(10, 20))

        # Filter by city
        tk.Label(
            search_frame,
            text="Ville:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(side="left")

        self.city_var = tk.StringVar()
        self.city_var.trace("w", self._on_filter_change)
        self.city_combo = ttk.Combobox(
            search_frame,
            textvariable=self.city_var,
            font=ENTRY_FONT,
            width=20,
            state="readonly"
        )
        self.city_combo.pack(side="left", padx=(10, 0))
        self.city_combo['values'] = ["Toutes les villes"]

        # Buttons frame
        btn_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        btn_frame.pack(fill="x", padx=20, pady=(0, 10))

        tk.Button(
            btn_frame,
            text="🔄 Actualiser",
            command=self._load_hotels,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT
        ).pack(side="left", padx=5)

        tk.Button(
            btn_frame,
            text="➕ Nouvel hôtel",
            command=self._new_hotel,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT
        ).pack(side="left", padx=5)

        self.btn_edit = tk.Button(
            btn_frame,
            text="✏️ Modifier",
            command=self._edit_selected,
            bg=BUTTON_ORANGE,
            fg="white",
            font=BUTTON_FONT,
            state="disabled"
        )
        self.btn_edit.pack(side="left", padx=5)

        self.btn_delete = tk.Button(
            btn_frame,
            text="🗑️ Supprimer",
            command=self._delete_selected,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            state="disabled"
        )
        self.btn_delete.pack(side="left", padx=5)

        # Treeview frame
        tree_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal")

        # Style for better selection appearance
        style = ttk.Style()
        style.configure(
            "Treeview",
            background=INPUT_BG_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR
        )
        style.map("Treeview", background=[("selected", BUTTON_GREEN)])

        # Treeview
        columns = ("row", "id", "nom", "lieu", "type_hebergement", "categorie",
                  "chambre_single", "chambre_double", "contact")
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            style="Treeview"
        )

        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)

        # Configure columns
        column_headers = {
            "row": "N°",
            "id": "ID",
            "nom": "Nom",
            "lieu": "Lieu",
            "type_hebergement": "Type",
            "categorie": "Catégorie",
            "chambre_single": "Single",
            "chambre_double": "Double",
            "contact": "Contact"
        }

        column_widths = {
            "row": 50,
            "id": 80,
            "nom": 150,
            "lieu": 120,
            "type_hebergement": 100,
            "categorie": 80,
            "chambre_single": 80,
            "chambre_double": 80,
            "contact": 120
        }

        for col in columns:
            self.tree.heading(col, text=column_headers[col])
            self.tree.column(col, width=column_widths[col], minwidth=column_widths[col])

        # Pack treeview and scrollbars
        self.tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        # Context menu
        self.context_menu = tk.Menu(self.parent, tearoff=0)
        self.context_menu.add_command(label="Modifier", command=self._edit_selected)
        self.context_menu.add_command(label="Supprimer", command=self._delete_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Voir détails", command=self._view_details)

        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<Double-1>", lambda e: self._edit_selected())
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_change)

        # Status label
        self.status_label = tk.Label(
            self.parent,
            text="",
            font=("Arial", 10),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        )
        self.status_label.pack(anchor="w", padx=20, pady=(0, 10))

        # Load hotels
        self._load_hotels()

    def _load_hotels(self):
        """Load and display all hotels"""
        self.hotels = load_all_hotels()
        
        # Extract unique cities for filter
        cities = sorted(set(hotel['lieu'] for hotel in self.hotels if hotel['lieu']))
        self.city_combo['values'] = ["Toutes les villes"] + cities
        if not self.city_var.get():
            self.city_var.set("Toutes les villes")
        
        self.filtered_hotels = self.hotels.copy()
        self._apply_filters()
        self._update_treeview()
        # Clear selection and disable buttons after reload
        self.tree.selection_remove(self.tree.selection())
        self._on_selection_change()

    def _fmt_ar(self, val):
        """Format a value as Ariary with thousand separators, or return empty/string safely."""
        if val is None or val == "":
            return ""
        try:
            return f"{int(val):,} Ar"
        except (ValueError, TypeError):
            try:
                return f"{float(val):,} Ar"
            except Exception:
                return str(val)

    def _update_treeview(self):
        """Update the treeview with filtered hotels"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Add filtered hotels
        for hotel in self.filtered_hotels:
            values = (
                hotel['row_number'],
                hotel['id'],
                hotel['nom'],
                hotel['lieu'],
                hotel['type_hebergement'],
                hotel['categorie'],
                self._fmt_ar(hotel.get('chambre_single')),
                self._fmt_ar(hotel.get('chambre_double')),
                hotel['contact']
            )
            self.tree.insert("", "end", values=values, tags=(str(hotel['row_number']),))

        # Update status label
        total_hotels = len(self.hotels)
        filtered_count = len(self.filtered_hotels)
        
        status_parts = []
        if self.search_var.get():
            status_parts.append(f"recherche: '{self.search_var.get()}'")
        if self.city_var.get() != "Toutes les villes":
            status_parts.append(f"ville: {self.city_var.get()}")
        
        if status_parts:
            filter_text = " | ".join(status_parts)
            self.status_label.config(text=f"Affichage de {filtered_count} hôtel(s) sur {total_hotels} (filtré: {filter_text})")
        else:
            self.status_label.config(text=f"Total: {total_hotels} hôtel(s)")

    def _on_filter_change(self, *args):
        """Handle filter changes (search and city)"""
        self._apply_filters()
        self._update_treeview()

    def _apply_filters(self):
        """Apply search and city filters"""
        search_text = self.search_var.get().lower()
        selected_city = self.city_var.get()
        
        self.filtered_hotels = []
        for hotel in self.hotels:
            # City filter
            city_match = (selected_city == "Toutes les villes" or 
                         hotel['lieu'] == selected_city)
            
            # Search filter
            search_match = (not search_text or
                           search_text in hotel['nom'].lower() or
                           search_text in hotel['lieu'].lower() or
                           search_text in hotel['id'].lower() or
                           search_text in hotel['contact'].lower())
            
            if city_match and search_match:
                self.filtered_hotels.append(hotel)

    def _on_selection_change(self, event=None):
        """Handle selection change in treeview"""
        selection = self.tree.selection()
        if selection:
            # Enable buttons when a hotel is selected
            self.btn_edit.config(state="normal")
            self.btn_delete.config(state="normal")
        else:
            # Disable buttons when no hotel is selected
            self.btn_edit.config(state="disabled")
            self.btn_delete.config(state="disabled")

    def _show_context_menu(self, event):
        """Show context menu on right-click"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def _get_selected_hotel(self):
        """Get the currently selected hotel"""
        selection = self.tree.selection()
        if not selection:
            return None

        item = selection[0]
        values = self.tree.item(item, "values")
        row_number = int(values[0])

        # Find hotel by row number
        for hotel in self.hotels:
            if hotel['row_number'] == row_number:
                return hotel
        return None

    def _edit_selected(self):
        """Edit the selected hotel"""
        hotel = self._get_selected_hotel()
        if hotel and self.on_edit_hotel:
            self.on_edit_hotel(hotel)
        else:
            messagebox.showwarning("Aucun hôtel sélectionné", "Veuillez sélectionner un hôtel à modifier.")

    def _delete_selected(self):
        """Delete the selected hotel"""
        hotel = self._get_selected_hotel()
        if not hotel:
            messagebox.showwarning("Aucun hôtel sélectionné", "Veuillez sélectionner un hôtel à supprimer.")
            return

        if messagebox.askyesno("Confirmation",
                              f"Êtes-vous sûr de vouloir supprimer l'hôtel {hotel['nom']} ?"):
            if delete_hotel_from_excel(hotel['row_number']):
                messagebox.showinfo("Succès", "Hôtel supprimé avec succès !")
                self._load_hotels()
            else:
                messagebox.showerror("Erreur", "Erreur lors de la suppression de l'hôtel")

    def _view_details(self):
        """View detailed information of selected hotel"""
        hotel = self._get_selected_hotel()
        if not hotel:
            return

        details = f"""Détails de l'hôtel:

    Nom: {hotel['nom']}
    ID: {hotel['id']}
    Lieu: {hotel['lieu']}
    Type: {hotel['type_hebergement']}
    Catégorie: {hotel['categorie']}

    TARIFICATION:
    Chambre Single: {self._fmt_ar(hotel.get('chambre_single'))}
    Chambre Double: {self._fmt_ar(hotel.get('chambre_double'))}
    Chambre Familiale: {self._fmt_ar(hotel.get('chambre_familiale'))}
    Lit Supplémentaire: {self._fmt_ar(hotel.get('lit_supp'))}
    Day Use: {self._fmt_ar(hotel.get('day_use'))}

    FRAIS ADDITIONNELS:
    Vignette: {self._fmt_ar(hotel.get('vignette'))}
    Taxe de séjour: {self._fmt_ar(hotel.get('taxe_sejour'))}

    REPAS:
    Petit déjeuner: {self._fmt_ar(hotel.get('petit_dejeuner'))}
    Déjeuner: {self._fmt_ar(hotel.get('dejeuner'))}
    Dîner: {self._fmt_ar(hotel.get('diner'))}

    CONTACT:
    {hotel['contact']}
    {hotel['email']}

    Description: {hotel['description']}"""

        messagebox.showinfo("Détails hôtel", details)

    def _new_hotel(self):
        """Create a new hotel"""
        if self.on_new_hotel:
            self.on_new_hotel()
