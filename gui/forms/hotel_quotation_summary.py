"""
Hotel quotation summary and grouping GUI component
"""

import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    ACCENT_TEXT_COLOR,
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    CARD_BG_COLOR,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.excel_handler import (
    get_quotations_by_city,
    get_quotations_grouped_by_client,
    load_all_hotel_quotations,
)
from utils.logger import logger


class HotelQuotationSummary:
    """
    Component to display hotel quotations in a simple summary table.
    """

    def __init__(self, parent):
        """
        Initialize quotation summary component

        Args:
            parent: Parent widget
        """
        self.parent = parent
        self.quotations = []
        self.tree = None
        self.search_var = None

        self._load_quotations()
        self._create_interface()

    def _load_quotations(self):
        """Load quotations from Excel"""
        try:
            self.quotations = load_all_hotel_quotations()
            logger.info(f"Loaded {len(self.quotations)} quotations")
        except Exception as e:
            logger.error(f"Error loading quotations: {e}", exc_info=True)
            self.quotations = []

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
            bg=MAIN_BG_COLOR,
        )
        title.pack(pady=(20, 10), fill="x")

        # Use a main frame to occupy full height (consistent with client form)
        self.main_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        style = ttk.Style()
        style.configure(
            "Treeview",
            background=INPUT_BG_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR,
        )
        style.map("Treeview", background=[("selected", BUTTON_GREEN)])

        # Search frame
        search_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        search_frame.pack(fill="x", padx=8, pady=(0, 8))

        tk.Label(
            search_frame,
            text="Recherche:",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        self.search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=ENTRY_FONT,
            width=30,
            bg=INPUT_BG_COLOR,
            fg=TEXT_COLOR,
        )
        search_entry.pack(side="left", padx=(10, 20))
        self.search_var.trace("w", self._on_search_changed)

        refresh_btn = tk.Button(
            search_frame,
            text="🔄 Rafraîchir",
            command=self._refresh_data,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=5,
        )
        refresh_btn.pack(side="left")

        self.kpi_label = tk.Label(
            self.main_frame,
            text="",
            font=LABEL_FONT,
            fg=ACCENT_TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            anchor="w",
        )
        self.kpi_label.pack(fill="x", padx=8, pady=(0, 8))

        # Main content frame
        self.content_frame = tk.Frame(self.main_frame, bg=MAIN_BG_COLOR)
        self.content_frame.pack(fill="both", expand=True, padx=0, pady=(0, 0))

        self._display_simple_table()

    def _get_quotations_by_hotel(self):
        """Group quotations by hotel with subtotals"""
        grouped = {}
        for quotation in self.quotations:
            hotel_name = quotation.get("hotel_name") or ""
            if hotel_name not in grouped:
                grouped[hotel_name] = {
                    "quotations": [],
                    "total": 0,
                    "currency": quotation.get("currency") or "Ariary",
                }
            grouped[hotel_name]["quotations"].append(quotation)
            grouped[hotel_name]["total"] += quotation.get("total_price", 0)
        return grouped

    def _get_search_query(self):
        """Return normalized search query"""
        if not self.search_var:
            return ""
        return self.search_var.get().strip().lower()

    def _on_view_changed(self, event=None):
        """Handle view change"""
        view = self.view_var.get()
        if "client" in view.lower():
            self._display_by_client()
        elif "ville" in view.lower():
            self._display_by_city()
        else:
            self._display_by_hotel()

    def _on_search_changed(self, *args):
        """Handle search query change"""
        self._display_simple_table()

    def _format_total(self, value, currency):
        try:
            return f"{float(value):,.2f} {currency}"
        except Exception:
            return f"{value} {currency}"

    def _display_simple_table(self):
        """Display quotations as a simple table with key information only."""
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        query = self._get_search_query()
        rows = []
        for q in self.quotations:
            if query:
                haystack = " ".join(
                    [
                        str(q.get("quote_date", "")),
                        str(q.get("client_id", "")),
                        str(q.get("client_name", "")),
                        str(q.get("hotel_name", "")),
                        str(q.get("city", "")),
                        str(q.get("room_type", "")),
                    ]
                ).lower()
                if query not in haystack:
                    continue
            rows.append(q)

        total_amount = sum(float(q.get("total_price", 0) or 0) for q in rows)
        self.kpi_label.config(
            text=f"Lignes: {len(rows)}   |   Total: {total_amount:,.2f}"
        )

        if not rows:
            tk.Label(
                self.content_frame,
                text="Aucune cotation trouvée",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            self.tree = None
            return

        table_frame = tk.Frame(self.content_frame, bg=MAIN_BG_COLOR)
        table_frame.pack(fill="both", expand=True, padx=8, pady=(0, 0))

        columns = ("date", "client", "ville", "hotel", "nuits", "total")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.tree.heading("date", text="Date")
        self.tree.heading("client", text="Client")
        self.tree.heading("ville", text="Ville")
        self.tree.heading("hotel", text="Hôtel")
        self.tree.heading("nuits", text="Nuits")
        self.tree.heading("total", text="Total")

        self.tree.column("date", width=130, anchor="w")
        self.tree.column("client", width=170, anchor="w")
        self.tree.column("ville", width=120, anchor="w")
        self.tree.column("hotel", width=220, anchor="w")
        self.tree.column("nuits", width=70, anchor="center")
        self.tree.column("total", width=130, anchor="e")

        for q in rows:
            self.tree.insert(
                "",
                "end",
                values=(
                    q.get("quote_date", ""),
                    q.get("client_name", ""),
                    q.get("city", ""),
                    q.get("hotel_name", ""),
                    q.get("nights", ""),
                    self._format_total(q.get("total_price", 0), q.get("currency", "")),
                ),
            )

        scrollbar_y = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.tree.yview
        )
        scrollbar_x = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(
            yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set
        )
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

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
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            return

        query = self._get_search_query()
        filtered_clients = []
        for client_id, client_data in sorted(self.grouped_by_client.items()):
            if not query:
                filtered_clients.append((client_id, client_data))
                continue
            client_name = (client_data.get("client_name") or "").lower()
            client_ref = str(client_id).lower()
            if query in client_name or query in client_ref:
                filtered_clients.append((client_id, client_data))

        if not filtered_clients:
            tk.Label(
                self.content_frame,
                text="Aucun résultat pour cette recherche",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            return

        # Use parent scrollable area (full height like client form)
        scrollable_frame = self.content_frame

        # Grand total
        grand_total = sum(client_data["total"] for _, client_data in filtered_clients)
        currency = filtered_clients[0][1]["currency"] if filtered_clients else "Ariary"

        grand_total_frame = tk.Frame(
            scrollable_frame, bg=CARD_BG_COLOR, bd=2, relief="ridge"
        )
        grand_total_frame.pack(fill="x", pady=(0, 10), padx=0)

        tk.Label(
            grand_total_frame,
            text="TOTAL GÉNÉRAL",
            font=("Poppins", 12, "bold"),
            fg=ACCENT_TEXT_COLOR,
            bg=CARD_BG_COLOR,
        ).pack(side="left", padx=15, pady=5)

        tk.Label(
            grand_total_frame,
            text=f"{grand_total:,.2f} {currency}",
            font=("Poppins", 12, "bold"),
            fg=ACCENT_TEXT_COLOR,
            bg=CARD_BG_COLOR,
        ).pack(side="right", padx=15, pady=5)

        # Display each client's quotations
        for client_id, client_data in filtered_clients:
            self._create_client_frame(scrollable_frame, client_id, client_data)

        # Scrolling handled by the parent scrollable frame

    def _create_client_frame(self, parent, client_id, client_data):
        """Create a frame for a single client's quotations"""
        client_frame = tk.LabelFrame(
            parent,
            text=f"{client_data['client_name']} (ID: {client_id})",
            font=("Poppins", 11, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        client_frame.pack(fill="x", pady=10, padx=0)

        # Create treeview for quotations
        tree_frame = tk.Frame(client_frame, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True)

        columns = (
            "hotel",
            "city",
            "nights",
            "adults",
            "children",
            "meal_plan",
            "price",
        )
        tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            height=min(8, len(client_data["quotations"])),
            show="headings",
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
        for quotation in client_data["quotations"]:
            tree.insert(
                "",
                "end",
                values=(
                    quotation["hotel_name"],
                    quotation["city"],
                    quotation["nights"],
                    quotation["adults"],
                    quotation["children"],
                    quotation["meal_plan"],
                    f"{quotation['total_price']:,.2f} {quotation['currency']}",
                ),
            )

        tree.pack(fill="both", expand=True)

        # Client subtotal
        subtotal_frame = tk.Frame(client_frame, bg=MAIN_BG_COLOR)
        subtotal_frame.pack(fill="x", pady=(10, 0))

        tk.Label(
            subtotal_frame,
            text=f"Sous-total {client_data['client_name']}:",
            font=("Poppins", 10, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        tk.Label(
            subtotal_frame,
            text=f"{client_data['total']:,.2f} {client_data['currency']}",
            font=("Poppins", 10, "bold"),
            fg=BUTTON_GREEN,
            bg=MAIN_BG_COLOR,
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
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            return

        query = self._get_search_query()
        filtered_cities = []
        for city, city_data in sorted(self.grouped_by_city.items()):
            if not query:
                filtered_cities.append((city, city_data))
                continue
            if query in (city or "").lower():
                filtered_cities.append((city, city_data))

        if not filtered_cities:
            tk.Label(
                self.content_frame,
                text="Aucun résultat pour cette recherche",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            return

        # Use parent scrollable area (full height like client form)
        scrollable_frame = self.content_frame

        # Grand total
        grand_total = sum(city_data["total"] for _, city_data in filtered_cities)
        currency = filtered_cities[0][1]["currency"] if filtered_cities else "Ariary"

        grand_total_frame = tk.Frame(
            scrollable_frame, bg=CARD_BG_COLOR, bd=2, relief="ridge"
        )
        grand_total_frame.pack(fill="x", pady=(0, 10), padx=0)

        tk.Label(
            grand_total_frame,
            text="TOTAL GÉNÉRAL",
            font=("Poppins", 12, "bold"),
            fg=ACCENT_TEXT_COLOR,
            bg=CARD_BG_COLOR,
        ).pack(side="left", padx=15, pady=5)

        tk.Label(
            grand_total_frame,
            text=f"{grand_total:,.2f} {currency}",
            font=("Poppins", 12, "bold"),
            fg=ACCENT_TEXT_COLOR,
            bg=CARD_BG_COLOR,
        ).pack(side="right", padx=15, pady=5)

        # Display each city's quotations
        for city, city_data in filtered_cities:
            self._create_city_frame(scrollable_frame, city, city_data)

        # Scrolling handled by the parent scrollable frame

    def _create_city_frame(self, parent, city, city_data):
        """Create a frame for a single city's quotations"""
        city_frame = tk.LabelFrame(
            parent,
            text=f"Ville: {city}",
            font=("Poppins", 11, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        city_frame.pack(fill="x", pady=10, padx=0)

        # Create treeview for quotations
        tree_frame = tk.Frame(city_frame, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True)

        columns = ("hotel", "client", "nights", "adults", "meal_plan", "price")
        tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            height=min(8, len(city_data["quotations"])),
            show="headings",
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
        for quotation in city_data["quotations"]:
            tree.insert(
                "",
                "end",
                values=(
                    quotation["hotel_name"],
                    quotation["client_name"],
                    quotation["nights"],
                    quotation["adults"],
                    quotation["meal_plan"],
                    f"{quotation['total_price']:,.2f} {quotation['currency']}",
                ),
            )

        tree.pack(fill="both", expand=True)

        # City subtotal
        subtotal_frame = tk.Frame(city_frame, bg=MAIN_BG_COLOR)
        subtotal_frame.pack(fill="x", pady=(10, 0))

        tk.Label(
            subtotal_frame,
            text=f"Sous-total {city}:",
            font=("Poppins", 10, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        tk.Label(
            subtotal_frame,
            text=f"{city_data['total']:,.2f} {city_data['currency']}",
            font=("Poppins", 10, "bold"),
            fg=BUTTON_GREEN,
            bg=MAIN_BG_COLOR,
        ).pack(side="right")

    def _display_by_hotel(self):
        """Display quotations grouped by hotel"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        if not self.grouped_by_hotel:
            tk.Label(
                self.content_frame,
                text="Aucune cotation trouvée",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            return

        query = self._get_search_query()
        filtered_hotels = []
        for hotel, hotel_data in sorted(self.grouped_by_hotel.items()):
            if not query:
                filtered_hotels.append((hotel, hotel_data))
                continue
            if query in (hotel or "").lower():
                filtered_hotels.append((hotel, hotel_data))

        if not filtered_hotels:
            tk.Label(
                self.content_frame,
                text="Aucun résultat pour cette recherche",
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).pack(pady=20)
            return

        # Use parent scrollable area (full height like client form)
        scrollable_frame = self.content_frame

        # Grand total
        grand_total = sum(hotel_data["total"] for _, hotel_data in filtered_hotels)
        currency = filtered_hotels[0][1]["currency"] if filtered_hotels else "Ariary"

        grand_total_frame = tk.Frame(
            scrollable_frame, bg=CARD_BG_COLOR, bd=2, relief="ridge"
        )
        grand_total_frame.pack(fill="x", pady=(0, 10), padx=0)

        tk.Label(
            grand_total_frame,
            text="TOTAL GÉNÉRAL",
            font=("Poppins", 12, "bold"),
            fg=ACCENT_TEXT_COLOR,
            bg=CARD_BG_COLOR,
        ).pack(side="left", padx=15, pady=5)

        tk.Label(
            grand_total_frame,
            text=f"{grand_total:,.2f} {currency}",
            font=("Poppins", 12, "bold"),
            fg=ACCENT_TEXT_COLOR,
            bg=CARD_BG_COLOR,
        ).pack(side="right", padx=15, pady=5)

        # Display each hotel's quotations
        for hotel, hotel_data in filtered_hotels:
            self._create_hotel_frame(scrollable_frame, hotel, hotel_data)

    def _create_hotel_frame(self, parent, hotel, hotel_data):
        """Create a frame for a single hotel's quotations"""
        hotel_frame = tk.LabelFrame(
            parent,
            text=f"Hôtel: {hotel}",
            font=("Poppins", 11, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        hotel_frame.pack(fill="x", pady=10, padx=0)

        # Create treeview for quotations
        tree_frame = tk.Frame(hotel_frame, bg=MAIN_BG_COLOR)
        tree_frame.pack(fill="both", expand=True)

        columns = (
            "client",
            "city",
            "nights",
            "adults",
            "children",
            "meal_plan",
            "price",
        )
        tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            height=min(8, len(hotel_data["quotations"])),
            show="headings",
        )

        # Define column headings
        tree.heading("client", text="Client")
        tree.heading("city", text="Ville")
        tree.heading("nights", text="Nuits")
        tree.heading("adults", text="Adultes")
        tree.heading("children", text="Enfants")
        tree.heading("meal_plan", text="Restauration")
        tree.heading("price", text="Total")

        # Define column widths
        tree.column("client", width=120)
        tree.column("city", width=80)
        tree.column("nights", width=50)
        tree.column("adults", width=50)
        tree.column("children", width=50)
        tree.column("meal_plan", width=120)
        tree.column("price", width=100)

        # Add data to tree
        for quotation in hotel_data["quotations"]:
            tree.insert(
                "",
                "end",
                values=(
                    quotation.get("client_name", ""),
                    quotation.get("city", ""),
                    quotation.get("nights", ""),
                    quotation.get("adults", ""),
                    quotation.get("children", ""),
                    quotation.get("meal_plan", ""),
                    f"{quotation.get('total_price', 0):,.2f} {quotation.get('currency', '')}",
                ),
            )

        tree.pack(fill="both", expand=True)

        # Hotel subtotal
        subtotal_frame = tk.Frame(hotel_frame, bg=MAIN_BG_COLOR)
        subtotal_frame.pack(fill="x", pady=(10, 0))

        tk.Label(
            subtotal_frame,
            text=f"Sous-total {hotel}:",
            font=("Poppins", 10, "bold"),
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        ).pack(side="left")

        tk.Label(
            subtotal_frame,
            text=f"{hotel_data['total']:,.2f} {hotel_data['currency']}",
            font=("Poppins", 10, "bold"),
            fg=BUTTON_GREEN,
            bg=MAIN_BG_COLOR,
        ).pack(side="right")

    def _refresh_data(self):
        """Refresh quotation data"""
        self._load_quotations()
        self._display_simple_table()
