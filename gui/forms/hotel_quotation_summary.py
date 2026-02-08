"""
Hotel quotation summary and grouping GUI component
"""

import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from config import *
from utils.excel_handler import (
    load_all_hotel_quotations, 
    get_quotations_grouped_by_client,
    get_quotations_by_city
)
from utils.logger import logger


class HotelQuotationSummary:
    """
    Component to display hotel quotations grouped by client or city with totals
    """

    def __init__(self, parent):
        """
        Initialize quotation summary component

        Args:
            parent: Parent widget
        """
        self.parent = parent
        self.quotations = []
        self.grouped_by_client = {}
        self.grouped_by_city = {}
        self.current_view = "by_client"  # 'by_client' or 'by_city'

        self._load_quotations()
        self._create_interface()

    def _load_quotations(self):
        """Load quotations from Excel"""
        try:
            self.quotations = load_all_hotel_quotations()
            self.grouped_by_client = get_quotations_grouped_by_client()
            self.grouped_by_city = get_quotations_by_city()
            logger.info(f"Loaded {len(self.quotations)} quotations")
        except Exception as e:
            logger.error(f"Error loading quotations: {e}", exc_info=True)
            self.quotations = []
            self.grouped_by_client = {}
            self.grouped_by_city = {}

    def _create_interface(self):
        """Create the quotation summary interface"""
        # Clear parent
        for widget in self.parent.winfo_children():
            widget.destroy()

        # Title
        title = tk.Label(
            self.parent,
            text="RÉSUMÉ DES COTATIONS HÔTEL",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        )
        title.pack(pady=(20, 10), fill="x")

        # View selection frame
        view_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        view_frame.pack(fill="x", padx=8, pady=(0, 10))

        tk.Label(
            view_frame,
            text="Afficher par:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(side="left")

        self.view_var = tk.StringVar(value="by_client")
        view_combo = ttk.Combobox(
            view_frame,
            textvariable=self.view_var,
            values=["Par client", "Par ville"],
            font=ENTRY_FONT,
            width=20,
            state="readonly"
        )
        view_combo.pack(side="left", padx=(10, 20))
        view_combo.bind("<<ComboboxSelected>>", self._on_view_changed)

        refresh_btn = tk.Button(
            view_frame,
            text="🔄 Rafraîchir",
            command=self._refresh_data,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=5
        )
        refresh_btn.pack(side="left")

        # Main content frame
        self.content_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.content_frame.pack(fill="both", expand=True, padx=0, pady=(0, 0))

        # Display the current view
        self._display_by_client()

    def _on_view_changed(self, event=None):
        """Handle view change"""
        view = self.view_var.get()
        if "client" in view.lower():
            self._display_by_client()
        else:
            self._display_by_city()

    def _display_by_client(self):
        """Display quotations grouped by client"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        if not self.grouped_by_client:
            tk.Label(
                self.content_frame,
                text="Aucune cotation trouvée",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR
            ).pack(pady=20)
            return

        # Create scrollable frame
        canvas = tk.Canvas(
            self.content_frame,
            bg=MAIN_BG_COLOR,
            highlightthickness=0
        )
        scrollbar = ttk.Scrollbar(self.content_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=MAIN_BG_COLOR)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        # Make the scrollable_frame match the canvas width so content fills available space
        canvas.bind(
            "<Configure>",
            lambda e, cid=window_id: canvas.itemconfig(cid, width=e.width)
        )

        # Grand total
        grand_total = sum(
            client_data['total'] 
            for client_data in self.grouped_by_client.values()
        )
        currency = list(self.grouped_by_client.values())[0]['currency'] if self.grouped_by_client else 'Ariary'

        grand_total_frame = tk.Frame(scrollable_frame, bg="#E8F4F8", bd=2, relief="ridge")
        grand_total_frame.pack(fill="x", pady=(0, 10), padx=0)

        tk.Label(
            grand_total_frame,
            text="TOTAL GÉNÉRAL",
            font=("Arial", 12, "bold"),
            fg="#003366",
            bg="#E8F4F8"
        ).pack(side="left", padx=15, pady=5)

        tk.Label(
            grand_total_frame,
            text=f"{grand_total:,.2f} {currency}",
            font=("Arial", 12, "bold"),
            fg="#003366",
            bg="#E8F4F8"
        ).pack(side="right", padx=15, pady=5)

        # Display each client's quotations
        for client_id, client_data in sorted(self.grouped_by_client.items()):
            self._create_client_frame(scrollable_frame, client_id, client_data)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _create_client_frame(self, parent, client_id, client_data):
        """Create a frame for a single client's quotations"""
        client_frame = tk.LabelFrame(
            parent,
            text=f"{client_data['client_name']} (ID: {client_id})",
            font=("Arial", 11, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10
        )
        client_frame.pack(fill="x", pady=10, padx=0)

        # Create treeview for quotations
        tree_frame = tk.Frame(client_frame, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True)

        columns = ("hotel", "city", "nights", "adults", "children", "meal_plan", "price")
        tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            height=min(8, len(client_data['quotations'])),
            show="headings"
        )

        # Define column headings
        tree.heading("hotel", text="Hôtel")
        tree.heading("city", text="Ville")
        tree.heading("nights", text="Nuits")
        tree.heading("adults", text="Adultes")
        tree.heading("children", text="Enfants")
        tree.heading("meal_plan", text="Restauration")
        tree.heading("price", text="Total")

        # Define column widths
        tree.column("hotel", width=150)
        tree.column("city", width=80)
        tree.column("nights", width=50)
        tree.column("adults", width=50)
        tree.column("children", width=50)
        tree.column("meal_plan", width=120)
        tree.column("price", width=100)

        # Add data to tree
        for quotation in client_data['quotations']:
            tree.insert(
                "",
                "end",
                values=(
                    quotation['hotel_name'],
                    quotation['city'],
                    quotation['nights'],
                    quotation['adults'],
                    quotation['children'],
                    quotation['meal_plan'],
                    f"{quotation['total_price']:,.2f} {quotation['currency']}"
                )
            )

        tree.pack(fill="both", expand=True)

        # Client subtotal
        subtotal_frame = tk.Frame(client_frame, bg=MAIN_BG_COLOR)
        subtotal_frame.pack(fill="x", pady=(10, 0))

        tk.Label(
            subtotal_frame,
            text=f"Sous-total {client_data['client_name']}:",
            font=("Arial", 10, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(side="left")

        tk.Label(
            subtotal_frame,
            text=f"{client_data['total']:,.2f} {client_data['currency']}",
            font=("Arial", 10, "bold"),
            fg=BUTTON_GREEN,
            bg=MAIN_BG_COLOR
        ).pack(side="right")

    def _display_by_city(self):
        """Display quotations grouped by city"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        if not self.grouped_by_city:
            tk.Label(
                self.content_frame,
                text="Aucune cotation trouvée",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR
            ).pack(pady=20)
            return

        # Create scrollable frame
        canvas = tk.Canvas(
            self.content_frame,
            bg=MAIN_BG_COLOR,
            highlightthickness=0
        )
        scrollbar = ttk.Scrollbar(self.content_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=MAIN_BG_COLOR)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind(
            "<Configure>",
            lambda e, cid=window_id: canvas.itemconfig(cid, width=e.width)
        )

        # Grand total
        grand_total = sum(
            city_data['total'] 
            for city_data in self.grouped_by_city.values()
        )
        currency = list(self.grouped_by_city.values())[0]['currency'] if self.grouped_by_city else 'Ariary'

        grand_total_frame = tk.Frame(scrollable_frame, bg="#E8F4F8", bd=2, relief="ridge")
        grand_total_frame.pack(fill="x", pady=(0, 10), padx=0)

        tk.Label(
            grand_total_frame,
            text="TOTAL GÉNÉRAL",
            font=("Arial", 12, "bold"),
            fg="#003366",
            bg="#E8F4F8"
        ).pack(side="left", padx=15, pady=5)

        tk.Label(
            grand_total_frame,
            text=f"{grand_total:,.2f} {currency}",
            font=("Arial", 12, "bold"),
            fg="#003366",
            bg="#E8F4F8"
        ).pack(side="right", padx=15, pady=5)

        # Display each city's quotations
        for city, city_data in sorted(self.grouped_by_city.items()):
            self._create_city_frame(scrollable_frame, city, city_data)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _create_city_frame(self, parent, city, city_data):
        """Create a frame for a single city's quotations"""
        city_frame = tk.LabelFrame(
            parent,
            text=f"Ville: {city}",
            font=("Arial", 11, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10
        )
        city_frame.pack(fill="x", pady=10, padx=0)

        # Create treeview for quotations
        tree_frame = tk.Frame(city_frame, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True)

        columns = ("hotel", "client", "nights", "adults", "meal_plan", "price")
        tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            height=min(8, len(city_data['quotations'])),
            show="headings"
        )

        # Define column headings
        tree.heading("hotel", text="Hôtel")
        tree.heading("client", text="Client")
        tree.heading("nights", text="Nuits")
        tree.heading("adults", text="Adultes")
        tree.heading("meal_plan", text="Restauration")
        tree.heading("price", text="Total")

        # Define column widths
        tree.column("hotel", width=150)
        tree.column("client", width=120)
        tree.column("nights", width=50)
        tree.column("adults", width=50)
        tree.column("meal_plan", width=120)
        tree.column("price", width=100)

        # Add data to tree
        for quotation in city_data['quotations']:
            tree.insert(
                "",
                "end",
                values=(
                    quotation['hotel_name'],
                    quotation['client_name'],
                    quotation['nights'],
                    quotation['adults'],
                    quotation['meal_plan'],
                    f"{quotation['total_price']:,.2f} {quotation['currency']}"
                )
            )

        tree.pack(fill="both", expand=True)

        # City subtotal
        subtotal_frame = tk.Frame(city_frame, bg=MAIN_BG_COLOR)
        subtotal_frame.pack(fill="x", pady=(10, 0))

        tk.Label(
            subtotal_frame,
            text=f"Sous-total {city}:",
            font=("Arial", 10, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR
        ).pack(side="left")

        tk.Label(
            subtotal_frame,
            text=f"{city_data['total']:,.2f} {city_data['currency']}",
            font=("Arial", 10, "bold"),
            fg=BUTTON_GREEN,
            bg=MAIN_BG_COLOR
        ).pack(side="right")

    def _refresh_data(self):
        """Refresh quotation data"""
        self._load_quotations()
        view = self.view_var.get()
        if "client" in view.lower():
            self._display_by_client()
        else:
            self._display_by_city()
        messagebox.showinfo("✅ Succès", "Les données ont été actualisées.")
