"""
Client list GUI component
"""

import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

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
from models.client_data import ClientData
from utils.excel_handler import (
    delete_client_from_excel,
    load_all_clients,
    update_client_statut,
)


class ClientList:
    """
    Client list component with search and management features
    """

    def __init__(self, parent, on_edit_client=None, on_new_client=None):
        """
        Initialize client list

        Args:
            parent: Parent widget
            on_edit_client: Callback for editing a client (receives client dict)
            on_new_client: Callback for creating a new client
        """
        self.parent = parent
        self.on_edit_client = on_edit_client
        self.on_new_client = on_new_client
        self.clients = []
        self.filtered_clients = []

        self._create_list()

    def _create_list(self):
        """Create the client list interface"""
        # Clear parent
        for widget in self.parent.winfo_children():
            widget.destroy()

        # Title
        title = tk.Label(
            self.parent,
            text="LISTE DES CLIENTS",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        title.pack(pady=(20, 10))

        # ── Barre de recherche + filtre statut ──────────────────────────
        search_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        search_frame.pack(fill="x", padx=20, pady=(0, 6))

        tk.Label(
            search_frame,
            text="Rechercher:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        self.search_var = tk.StringVar()
        self.search_var.trace("w", self._on_filter_change)
        search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=ENTRY_FONT,
            width=28,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        )
        search_entry.pack(side="left", padx=(8, 20))

        # Filtre par statut
        tk.Label(
            search_frame,
            text="Statut:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        self._STATUT_COLORS = {
            "Tous":       None,
            "En cours":   "#00BCD4",
            "Accepté":    "#4CAF50",
            "En circuit": "#FF9800",
            "Annulé":     "#F44336",
        }
        self._STATUT_ROW_BG = {
            "En cours":   "#E0F7FA",
            "Accepté":    "#E8F5E9",
            "En circuit": "#FFF3E0",
            "Annulé":     "#FFEBEE",
        }

        self.statut_filter_var = tk.StringVar(value="Tous")
        self.statut_filter_var.trace("w", self._on_filter_change)
        statut_combo = ttk.Combobox(
            search_frame,
            textvariable=self.statut_filter_var,
            values=list(self._STATUT_COLORS.keys()),
            state="readonly",
            width=14,
            font=ENTRY_FONT,
        )
        statut_combo.pack(side="left", padx=(8, 0))

        # Buttons frame
        btn_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        btn_frame.pack(fill="x", padx=20, pady=(0, 10))

        tk.Button(
            btn_frame,
            text="🔄 Actualiser",
            command=self._load_clients,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
        ).pack(side="left", padx=5)

        tk.Button(
            btn_frame,
            text="➕ Nouveau client",
            command=self._new_client,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
        ).pack(side="left", padx=5)

        self.btn_voir = tk.Button(
            btn_frame,
            text="👁 Voir Résa",
            command=self._edit_selected,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
        )
        self.btn_voir.pack(side="left", padx=5)

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

        self.btn_annuler = tk.Button(
            btn_frame,
            text="🚫 Annuler",
            command=self._annuler_selected,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
        )
        self.btn_annuler.pack(side="left", padx=5)

        self.btn_reintegrer = tk.Button(
            btn_frame,
            text="↩ Réintégrer",
            command=self._reintegrer_selected,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
        )
        self.btn_reintegrer.pack(side="left", padx=5)

        self.btn_delete = tk.Button(
            btn_frame,
            text="🗑️ Supprimer",
            command=self._delete_selected,
            bg="#555555",
            fg="white",
            font=BUTTON_FONT,
            state="disabled",
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
            fieldbackground=INPUT_BG_COLOR,
        )
        style.map("Treeview", background=[("selected", BUTTON_GREEN)])

        # Treeview
        columns = (
            "statut",
            "row",
            "timestamp",
            "ref_client",
            "nom",
            "telephone",
            "email",
            "periode",
            "circuit",
            "date_arrivee",
            "date_depart",
            "duree_sejour",
        )
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            style="Treeview",
        )

        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)

        # Configure columns
        column_headers = {
            "statut":       "Statut",
            "row":          "N°",
            "timestamp":    "Date création",
            "ref_client":   "Réf. Client",
            "nom":          "Nom",
            "telephone":    "Téléphone",
            "email":        "Email",
            "periode":      "Période",
            "circuit":      "Circuit",
            "date_arrivee": "Arrivée",
            "date_depart":  "Départ",
            "duree_sejour": "Durée",
        }

        for col in columns:
            self.tree.heading(col, text=column_headers[col])
            self.tree.column(col, width=100, minwidth=80)

        self.tree.column("statut",      width=110, minwidth=90)
        self.tree.column("row",         width=40,  minwidth=30)
        self.tree.column("ref_client",  width=120, minwidth=100)
        self.tree.column("nom",         width=160, minwidth=120)
        self.tree.column("email",       width=180, minwidth=140)
        self.tree.column("circuit",     width=180, minwidth=140)
        self.tree.column("date_arrivee",width=100, minwidth=80)
        self.tree.column("date_depart", width=100, minwidth=80)
        self.tree.column("duree_sejour",width=70,  minwidth=50)

        # Color tags per statut
        for s, bg in self._STATUT_ROW_BG.items():
            self.tree.tag_configure(s, background=bg)

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
        self.tree.bind("<Double-1>", lambda e: self._voir_or_edit())
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_change)

        # Status label
        self.status_label = tk.Label(
            self.parent, text="", font=("Poppins", 10), fg=TEXT_COLOR, bg=MAIN_BG_COLOR
        )
        self.status_label.pack(anchor="w", padx=20, pady=(0, 10))

        # Load clients
        self._load_clients()

    def _on_selection_change(self, event=None):
        """Handle selection change in treeview"""
        selection = self.tree.selection()
        if selection:
            client = self._get_selected_client()
            is_annule = client and (client.get("statut") or "En cours") == "Annulé"
            self.btn_voir.config(state="normal")
            # Dossier annulé : lecture seule, pas modifiable
            self.btn_edit.config(state="disabled" if is_annule else "normal")
            self.btn_annuler.config(state="disabled" if is_annule else "normal")
            self.btn_reintegrer.config(state="normal" if is_annule else "disabled")
            self.btn_delete.config(state="normal")
        else:
            for btn in (self.btn_voir, self.btn_edit, self.btn_annuler,
                        self.btn_reintegrer, self.btn_delete):
                btn.config(state="disabled")

    def _load_clients(self):
        """Load and display all clients"""
        self.clients = load_all_clients()
        self.filtered_clients = self.clients.copy()
        self._update_treeview()
        # Clear selection and disable buttons after reload
        self.tree.selection_remove(self.tree.selection())
        self._on_selection_change()

    def _update_treeview(self):
        """Update the treeview with filtered clients"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Add filtered clients
        for client in self.filtered_clients:
            statut = client.get("statut") or "En cours"
            values = (
                statut,
                client["row_number"],
                client["timestamp"],
                client["ref_client"],
                client["nom"],
                client["telephone"],
                client["email"],
                client["periode"],
                client["circuit"],
                client.get("date_arrivee", ""),
                client.get("date_depart", ""),
                client.get("duree_sejour", ""),
            )
            self.tree.insert(
                "", "end", values=values, tags=(statut,)
            )

        # Update status label
        total_clients = len(self.clients)
        filtered_count = len(self.filtered_clients)
        if self.search_var.get():
            self.status_label.config(
                text=f"Affichage de {filtered_count} client(s) sur {total_clients} (filtré)"
            )
        else:
            self.status_label.config(text=f"Total: {total_clients} client(s)")

    def _on_filter_change(self, *args):
        """Filter list by search text and/or statut."""
        search_text = self.search_var.get().lower()
        statut_filter = self.statut_filter_var.get()

        self.filtered_clients = [
            client for client in self.clients
            if (
                (statut_filter == "Tous" or (client.get("statut") or "En cours") == statut_filter)
                and (
                    not search_text
                    or search_text in client["nom"].lower()
                    or search_text in client["ref_client"].lower()
                    or search_text in client["email"].lower()
                    or search_text in client["telephone"].lower()
                    or search_text in str(client.get("circuit", "")).lower()
                    or search_text in str(client.get("numero_dossier", "")).lower()
                )
            )
        ]
        self._update_treeview()

    def _show_context_menu(self, event):
        """Show context menu on right-click"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def _get_selected_client(self):
        """Get the currently selected client"""
        selection = self.tree.selection()
        if not selection:
            return None

        item = selection[0]
        values = self.tree.item(item, "values")
        row_number = int(values[1])  # col 0 = statut, col 1 = row_number

        # Find client by row number
        for client in self.clients:
            if client["row_number"] == row_number:
                return client
        return None

    def _voir_or_edit(self):
        """Double-clic : ouvrir en lecture seule si annulé, en édition sinon."""
        client = self._get_selected_client()
        if not client:
            return
        if (client.get("statut") or "En cours") == "Annulé":
            self._view_details()
        else:
            self._edit_selected()

    def _edit_selected(self):
        """Edit the selected client"""
        client = self._get_selected_client()
        if client and self.on_edit_client:
            self.on_edit_client(client)
        else:
            messagebox.showwarning(
                "Aucun client sélectionné",
                "Veuillez sélectionner un client à modifier.",
            )

    def _delete_selected(self):
        """Delete the selected client"""
        client = self._get_selected_client()
        if not client:
            messagebox.showwarning(
                "Aucun client sélectionné",
                "Veuillez sélectionner un client à supprimer.",
            )
            return

        if messagebox.askyesno(
            "Confirmation",
            f"Êtes-vous sûr de vouloir supprimer le client {client['nom']} ?",
        ):
            if delete_client_from_excel(client["row_number"]):
                messagebox.showinfo("Succès", "Client supprimé avec succès !")
                self._load_clients()
            else:
                messagebox.showerror(
                    "Erreur", "Erreur lors de la suppression du client"
                )

    def _annuler_selected(self):
        """Annuler la réservation sélectionnée — passe le statut à 'Annulé'."""
        client = self._get_selected_client()
        if not client:
            return
        if not messagebox.askyesno(
            "Confirmation d'annulation",
            f"Voulez-vous annuler la réservation de {client['nom']} ?\n"
            "Le dossier sera fermé et ne pourra plus être modifié\n"
            "sauf réintégration.",
            icon="warning",
        ):
            return
        if update_client_statut(client["row_number"], "Annulé"):
            messagebox.showinfo("Annulé", f"Dossier {client['nom']} annulé.")
            self._load_clients()
        else:
            messagebox.showerror("Erreur", "Impossible d'annuler le dossier.")

    def _reintegrer_selected(self):
        """Réintégrer un dossier annulé — passe le statut à 'En cours'."""
        client = self._get_selected_client()
        if not client:
            return
        if not messagebox.askyesno(
            "Confirmation de réintégration",
            f"Voulez-vous réintégrer la réservation de {client['nom']} ?\n"
            "Le statut repassera à 'En cours'.",
        ):
            return
        if update_client_statut(client["row_number"], "En cours"):
            messagebox.showinfo("Réintégré", f"Dossier {client['nom']} réintégré.")
            self._load_clients()
        else:
            messagebox.showerror("Erreur", "Impossible de réintégrer le dossier.")

    def _view_details(self):
        """View detailed information of selected client"""
        client = self._get_selected_client()
        if not client:
            return

        details = f"""Détails du client:

Nom: {client['nom']}
Référence: {client['ref_client']}
Téléphone: {client['telephone']}
Email: {client['email']}
Période: {client['periode']}
Restauration: {client['restauration']}
Hébergement: {client['hebergement']}
Chambre: {client['chambre']}
Enfant: {client['enfant']}
Âge enfant: {client['age_enfant']}
Forfait: {client['forfait']}
Circuit: {client['circuit']}
ID Circuit: {client.get('id_circuit', '')}
Durée Circuit: {client.get('duree_circuit', '')}
Condition Physique: {client.get('condition_physique_circuit', '')}
Type de voiture: {client.get('type_voiture_circuit', '')}
Activité: {client.get('activite_circuit', '')}
Date d'ajout: {client['timestamp']}"""

        messagebox.showinfo("Détails client", details)

    def _new_client(self):
        """Create a new client"""
        if self.on_new_client:
            self.on_new_client()
