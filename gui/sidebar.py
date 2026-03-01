"""
Sidebar GUI component
"""

import customtkinter as ctk

try:
    from PIL import Image

    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

from config import (
    BUTTON_BLUE,
    BUTTON_GREEN,
    BUTTON_GREEN_HOVER,
    CURRENT_THEME,
    LOGO_PATH,
    SIDEBAR_BG_COLOR,
    TEXT_COLOR,
    apply_theme,
)


class Sidebar:
    """
    Sidebar component with navigation menu
    """

    def __init__(self, parent, main_content_callback, on_theme_change=None):
        """
        Initialize sidebar

        Args:
            parent: Parent widget
            main_content_callback: Callback function for main content updates
        """
        self.parent = parent
        self.main_content_callback = main_content_callback
        self.on_theme_change = on_theme_change
        self.menu_states = {}

        # Create scrollable sidebar
        self.sidebar_scroll = None
        self._build_sidebar()

    def _build_sidebar(self):
        """Setup sidebar content"""
        if self.sidebar_scroll is not None:
            self.sidebar_scroll.destroy()

        self.sidebar_scroll = ctk.CTkScrollableFrame(
            self.parent, width=250, fg_color=SIDEBAR_BG_COLOR, corner_radius=0
        )
        self.sidebar_scroll.grid(row=0, column=0, sticky="nswe")

        # Logo
        try:
            if PIL_AVAILABLE:
                logo_label = ctk.CTkLabel(
                    self.sidebar_scroll,
                    image=ctk.CTkImage(
                        light_image=Image.open(LOGO_PATH), size=(150, 150)
                    ),
                    text="",
                )
            else:
                raise ImportError
        except (ImportError, FileNotFoundError, Exception):
            # Fallback if logo not found or PIL not available
            logo_label = ctk.CTkLabel(
                self.sidebar_scroll,
                text="🏨 Lahimena Tours",
                font=("Arial", 16, "bold"),
            )
        logo_label.pack(pady=30)

        theme_frame = ctk.CTkFrame(self.sidebar_scroll, fg_color="transparent")
        theme_frame.pack(fill="x", padx=20, pady=(0, 10))

        theme_label = ctk.CTkLabel(
            theme_frame,
            text="Thème sombre",
            text_color=TEXT_COLOR,
            font=("Arial", 12, "bold"),
        )
        theme_label.pack(side="left")

        self.theme_switch = ctk.CTkSwitch(
            theme_frame,
            text="",
            command=self._toggle_theme,
            fg_color=BUTTON_BLUE,
            progress_color=BUTTON_GREEN,
        )
        if CURRENT_THEME == "dark":
            self.theme_switch.select()
        else:
            self.theme_switch.deselect()
        self.theme_switch.pack(side="right")

        # Menu buttons
        self._create_menu_buttons()

    def _create_menu_buttons(self):
        """Create all menu buttons"""
        # Home
        self._create_button("🏠 Accueil", self._show_home)

        # Client Information (direct access)
        _btn1 = self._create_button("🏨 Information client", self._show_client_page)

        # Hotel Quotation (direct access)
        _btn2 = self._create_button("🏨 Cotation hôtel", self._show_hotel_quotation_page)

        # Services
        _btn_collective = self._create_button(
            "🎯 Frais collectifs", self._show_collective_expense_page
        )
        _btn_transport = self._create_button("🎯 Transport", self._show_transport_page)
        _btn_air = self._create_button("🎯 Billet Avion", self._show_air_ticket_page)

        # Client Quotations (direct access)
        _btn4 = self._create_button("👥 Devis clients", self._show_client_quotes_page)

        # Client Invoices (direct access)
        _btn5 = self._create_button("💰 Facture clients", self._show_current_invoices)

        # Expenses
        btn6 = self._create_button("📉 Dépenses ▶", None)
        _submenu6_frame = self._create_submenu(
            btn6,
            [
                ("📄 Factures actuelles", self._show_current_expenses),
                ("✏️ Historiques factures", self._show_expense_history),
            ],
        )

        # Unified Databases section
        btn_db = self._create_button("🏨 Bases de données (BDD) ▶", None)
        _submenu_db_frame = self._create_submenu(
            btn_db,
            [
                ("📋 Hôtels", self._show_hotel_list),
                ("➕ Ajout hôtel", self._show_add_hotel),
                ("📋 Frais collectifs", self._show_collective_expense_list),
                ("➕ Ajout frais collectif", self._show_add_collective_expense),
                ("📋 Transport", self._show_transport_db_page),
                ("📋 Billets avion", self._show_air_ticket_list_db),
                ("➕ Ajout billet avion", self._show_add_air_ticket_db),
            ],
        )

        # Financial Statements (single entry point, no sidebar submenu)
        _btn10 = self._create_button("📊 Etat Financier", self._show_financial_home)

        # Accounting Entry
        _btn11 = self._create_button("📊 Saisie comptable", self._show_accounting_entry)

        # Excel Editor
        _btn12 = self._create_button(
            "📊 Éditeur Excel 'calcul' ▶", self._show_excel_editor
        )

    def _create_button(self, text, command=None):
        """Create a sidebar button"""
        btn = ctk.CTkButton(
            self.sidebar_scroll,
            text=text,
            fg_color=BUTTON_GREEN,
            hover_color=BUTTON_GREEN_HOVER,
            height=45,
            corner_radius=10,
            command=command,
        )
        btn.pack(pady=8, padx=20, fill="x")
        return btn

    def _create_submenu(self, parent_btn, items):
        """Create submenu for a button"""
        submenu_frame = ctk.CTkFrame(self.sidebar_scroll, fg_color="transparent")

        for text, command in items:
            ctk.CTkButton(
                submenu_frame,
                text=text,
                height=35,
                fg_color=(
                    "#10B981"
                    if "➕" in text
                    else (
                        "#007A93"
                        if "📝" in text
                        else "#D31F25" if "❌" in text else "#10B981"
                    )
                ),
                command=command,
            ).pack(pady=2, padx=10, fill="x")

        submenu_frame.pack_forget()
        self._create_toggle_menu(parent_btn, submenu_frame)
        return submenu_frame

    def _create_toggle_menu(self, btn, submenu_frame):
        """Create toggle functionality for expandable menus"""

        def toggle():
            btn_key = id(btn)
            if not self.menu_states.get(btn_key, False):
                submenu_frame.pack(after=btn, pady=(0, 5), padx=20, fill="x")
                btn.configure(text=btn.cget("text").replace("▶", "▼"))
                self.menu_states[btn_key] = True
            else:
                submenu_frame.pack_forget()
                btn.configure(text=btn.cget("text").replace("▼", "▶"))
                self.menu_states[btn_key] = False

        if "▶" not in btn.cget("text"):
            btn.configure(text=btn.cget("text") + " ▶")
        btn.configure(command=toggle)

    def _toggle_theme(self):
        theme_name = "dark" if self.theme_switch.get() == 1 else "light"
        apply_theme(theme_name)
        ctk.set_appearance_mode(theme_name)
        self._build_sidebar()
        if self.on_theme_change:
            self.on_theme_change()

    # Placeholder methods for menu actions
    def _show_home(self):
        self.main_content_callback("home")

    def _show_client_form(self):
        self.main_content_callback("client_page")

    def _show_client_list(self):
        self.main_content_callback("client_page")

    def _show_client_page(self):
        self.main_content_callback("client_page")

    def _show_hotel_list(self):
        self.main_content_callback("hotel_list")

    def _show_hotel_quotation(self):
        self.main_content_callback("hotel_quotation_page")

    def _show_hotel_quotation_summary(self):
        self.main_content_callback("hotel_quotation_page")

    def _show_hotel_quotation_page(self):
        self.main_content_callback("hotel_quotation_page")

    def _show_transport_page(self):
        self.main_content_callback("transport_page")

    def _show_collective_expense_quotation(self):
        self.main_content_callback("collective_expense_quotation")

    def _show_collective_expense_summary(self):
        self.main_content_callback("collective_expense_summary")

    def _show_collective_expense_page(self):
        self.main_content_callback("collective_expense_page")

    def _show_air_ticket_quotation(self):
        self.main_content_callback("air_ticket_quotation")

    def _show_air_ticket_summary(self):
        self.main_content_callback("air_ticket_summary")

    def _show_air_ticket_page(self):
        self.main_content_callback("air_ticket_page")

    def _show_visite_excursion_summary(self):
        self.main_content_callback("visite_excursion_summary")

    def _show_current_quotes(self):
        self.main_content_callback("client_quotes_page")

    def _show_quote_history(self):
        self.main_content_callback("client_quotes_page")

    def _show_client_quotes_page(self):
        self.main_content_callback("client_quotes_page")

    def _show_current_invoices(self):
        self.main_content_callback("current_invoices")

    def _show_invoice_history(self):
        self.main_content_callback("invoice_history")

    def _show_current_expenses(self):
        self.main_content_callback("current_expenses")

    def _show_expense_history(self):
        self.main_content_callback("expense_history")

    def _show_add_hotel(self):
        self.main_content_callback("hotel_form")

    def _show_add_collective_expense(self):
        self.main_content_callback("collective_expense_db_form")

    def _show_collective_expense_list(self):
        self.main_content_callback("collective_expense_db_list")

    def _show_transport_db_page(self):
        self.main_content_callback("transport_db_page")

    def _show_air_ticket_list_db(self):
        self.main_content_callback("air_ticket_db_list")

    def _show_add_air_ticket_db(self):
        self.main_content_callback("air_ticket_db_form")

    def _show_visite_excursion_list_db(self):
        self.main_content_callback("visite_excursion_db_list")

    def _show_add_visite_excursion_db(self):
        self.main_content_callback("visite_excursion_db_form")

    def _show_income_statement(self):
        self.main_content_callback("income_statement")

    def _show_financial_home(self):
        self.main_content_callback("financial_home")

    def _show_balance_sheet(self):
        self.main_content_callback("balance_sheet")

    def _show_cash_flow(self):
        self.main_content_callback("cash_flow")

    def _show_depreciation_table(self):
        self.main_content_callback("depreciation_table")

    def _show_fixed_assets(self):
        self.main_content_callback("fixed_assets")

    def _show_5year_forecast(self):
        self.main_content_callback("5year_forecast")

    def _show_12month_cash_forecast(self):
        self.main_content_callback("12month_cash_forecast")

    def _show_accounting_entry(self):
        self.main_content_callback("accounting_entry")

    def _show_excel_editor(self):
        self.main_content_callback("excel_editor")

    def _show_parametrage_page(self):
        self.main_content_callback("parametrage_page")
