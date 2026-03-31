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
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_GREEN_HOVER,
    BUTTON_RED,
    LOGO_PATH,
    SIDEBAR_BG_COLOR,
    TEXT_COLOR,
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
        self.active_button = None
        self.primary_buttons = []

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
                font=("Poppins", 16, "bold"),
                text_color=TEXT_COLOR,
            )
        logo_label.pack(pady=30)

        # Menu buttons
        self._create_menu_buttons()

    def _create_menu_buttons(self):
        """Create all menu buttons"""
        # Home
        home_btn = self._create_button("Accueil", self._show_home)
        self._set_active(home_btn)

        # Client Information (direct access)
        btn_clients = self._create_button("Infos Clients & Voyages", self._show_client_page)

        # Quotation module hub
        btn_cotation = self._create_button("Cotation", self._show_cotation_hub_page)

        # Invoices / Quotes hub
        btn_facture = self._create_button("Facture/Devis", self._show_billing_quotes_hub_page)

        # Unified Databases section (dedicated hub page)
        btn_bdd = self._create_button(
            "Bases de données (BDD)", self._show_database_hub_page
        )

        # Financial Statements (single entry point, no sidebar submenu)
        btn_financier = self._create_button("Etat financier", self._show_financial_home)

        # Marketing placeholder
        btn_marketing = self._create_button("Marketing", self._show_marketing_page)

        # Restrict navigation for comptable role
        try:
            from utils.auth_handler import current_role
            if current_role() == "comptable":
                for btn in self.primary_buttons:
                    if btn.cget("text") != "Etat financier":
                        btn.configure(
                            state="disabled",
                            fg_color="#AAAAAA",
                            hover_color="#AAAAAA",
                            text_color="#DDDDDD",
                        )
                # Set "Etat financier" as the active button
                self._set_active(btn_financier)
        except Exception:
            pass

    def _create_button(self, text, command=None):
        """Create a sidebar button"""
        btn = ctk.CTkButton(
            self.sidebar_scroll,
            text=text,
            fg_color=BUTTON_BLUE,
            hover_color=BUTTON_GREEN_HOVER,
            text_color="white",
            font=BUTTON_FONT,
            height=36,
            corner_radius=16,
        )
        def _on_click():
            self._set_active(btn)
            if command:
                command()

        btn.configure(command=_on_click)
        btn.pack(pady=8, padx=20, fill="x")
        self.primary_buttons.append(btn)
        return btn

    def _set_active(self, btn):
        if self.active_button is btn:
            return
        if self.active_button is not None:
            try:
                self.active_button.configure(fg_color=BUTTON_BLUE)
            except Exception:
                pass
        try:
            btn.configure(fg_color=BUTTON_RED)
        except Exception:
            pass
        self.active_button = btn

    def _create_submenu(self, parent_btn, items):
        """Create submenu for a button"""
        submenu_frame = ctk.CTkFrame(self.sidebar_scroll, fg_color="transparent")

        for text, command in items:
            label = text.lower()
            if "❌" in text or "supprimer" in label:
                button_color = BUTTON_RED
            elif "📝" in text or "modifier" in label:
                button_color = BUTTON_BLUE
            else:
                button_color = BUTTON_GREEN
            ctk.CTkButton(
                submenu_frame,
                text=text,
                height=35,
                fg_color=button_color,
                hover_color=BUTTON_GREEN_HOVER,
                text_color="white",
                font=BUTTON_FONT,
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

    def _show_cotation_hub_page(self):
        self.main_content_callback("cotation_hub_page")

    def _show_database_hub_page(self):
        self.main_content_callback("database_hub_page")

    def _show_billing_quotes_hub_page(self):
        self.main_content_callback("billing_quotes_hub_page")

    def _show_marketing_page(self):
        self.main_content_callback("marketing_page")

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

    def _show_circuit_db_page(self):
        self.main_content_callback("circuit_db_page")

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
