"""
Sidebar GUI component
"""

import customtkinter as ctk

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

from config import *


class Sidebar:
    """
    Sidebar component with navigation menu
    """

    def __init__(self, parent, main_content_callback):
        """
        Initialize sidebar

        Args:
            parent: Parent widget
            main_content_callback: Callback function for main content updates
        """
        self.parent = parent
        self.main_content_callback = main_content_callback
        self.menu_states = {}

        # Create scrollable sidebar
        self.sidebar_scroll = ctk.CTkScrollableFrame(
            parent,
            width=250,
            fg_color=SIDEBAR_BG_COLOR,
            corner_radius=0
        )
        self.sidebar_scroll.grid(row=0, column=0, sticky="nswe")

        self._setup_sidebar()

    def _setup_sidebar(self):
        """Setup sidebar content"""
        # Logo
        try:
            if PIL_AVAILABLE:
                logo_label = ctk.CTkLabel(
                    self.sidebar_scroll,
                    image=ctk.CTkImage(light_image=Image.open(LOGO_PATH), size=(150, 150)),
                    text=""
                )
            else:
                raise ImportError
        except:
            # Fallback if logo not found or PIL not available
            logo_label = ctk.CTkLabel(
                self.sidebar_scroll,
                text="🏨 Lahimena Tours",
                font=("Arial", 16, "bold")
            )
        logo_label.pack(pady=30)

        # Menu buttons
        self._create_menu_buttons()

    def _create_menu_buttons(self):
        """Create all menu buttons"""
        # Client Information
        btn1 = self._create_button("🏨 Information client ▶", self._show_client_form)
        submenu1_frame = self._create_submenu(btn1, [
            ("👤 Nouveau client", self._show_client_form),
            ("📋 Liste clients", self._show_client_list),
            ("🔍 Recherche", self._show_client_search)
        ])

        # Hotel Quotation
        btn2 = self._create_button("🏨 Cotation hôtel", self._show_hotel_quotation)

        # Service Quotation
        btn3 = self._create_button("🎯 Cotation prestation", self._show_service_quotation)

        # Client Quotations
        btn4 = self._create_button("👥 Devis clients ▶", None)
        submenu4_frame = self._create_submenu(btn4, [
            ("📄 Devis actuels", self._show_current_quotes),
            ("✏️ Historiques devis", self._show_quote_history)
        ])

        # Client Invoices
        btn5 = self._create_button("💰 Facture clients ▶", None)
        submenu5_frame = self._create_submenu(btn5, [
            ("📄 Factures actuelles", self._show_current_invoices),
            ("✏️ Historiques factures", self._show_invoice_history)
        ])

        # Expenses
        btn6 = self._create_button("📉 Dépenses ▶", None)
        submenu6_frame = self._create_submenu(btn6, [
            ("📄 Factures actuelles", self._show_current_expenses),
            ("✏️ Historiques factures", self._show_expense_history)
        ])

        # Hotel Database
        btn7 = self._create_button("🏨 Hôtel (DB) ▶", None)
        submenu7_frame = self._create_submenu(btn7, [
            ("➕ Ajout hôtel", self._show_add_hotel),
            ("📝 Mise à jour", self._show_update_hotel),
            ("❌ Suppression", self._show_delete_hotel)
        ])

        # Collective Expenses Database
        btn8 = self._create_button("🏨 Frais collectif (DB) ▶", None)
        submenu8_frame = self._create_submenu(btn8, [
            ("➕ Ajouter", self._show_add_collective_expense),
            ("📝 Mettre à jour", self._show_update_collective_expense),
            ("❌ Supprimer", self._show_delete_collective_expense)
        ])

        # Transport Database
        btn9 = self._create_button("🏨 Transport (DB) ▶", None)
        submenu9_frame = self._create_submenu(btn9, [
            ("➕ Ajouter", self._show_add_transport),
            ("📝 Mettre à jour", self._show_update_transport),
            ("❌ Supprimer", self._show_delete_transport)
        ])

        # Financial Statements
        btn10 = self._create_button("📊 Etat Financier ▶", None)
        submenu10_frame = self._create_submenu(btn10, [
            ("📊 Compte de résultat", self._show_income_statement),
            ("📊 Bilan", self._show_balance_sheet),
            ("📊 Tableau Flux de trésorerie", self._show_cash_flow),
            ("📊 Tableau d'amortissement", self._show_depreciation_table),
            ("📊 Liste des immobilisations", self._show_fixed_assets),
            ("📊 Prévisionnel sur 5 ans", self._show_5year_forecast),
            ("📊 Prévisionnel trésorerie 12 mois", self._show_12month_cash_forecast)
        ])

        # Accounting Entry
        btn11 = self._create_button("📊 Saisie comptable", self._show_accounting_entry)

        # Excel Editor
        btn12 = self._create_button("📊 Éditeur Excel 'calcul' ▶", self._show_excel_editor)

    def _create_button(self, text, command=None):
        """Create a sidebar button"""
        btn = ctk.CTkButton(
            self.sidebar_scroll,
            text=text,
            fg_color=BUTTON_GREEN,
            hover_color=BUTTON_GREEN_HOVER,
            height=45,
            corner_radius=10,
            command=command
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
                fg_color="#10B981" if "➕" in text else "#007A93" if "📝" in text else "#D31F25" if "❌" in text else "#10B981",
                command=command
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
    def _show_client_form(self):
        self.main_content_callback("client_form")

    def _show_client_list(self):
        self.main_content_callback("client_list")

    def _show_client_search(self):
        self.main_content_callback("client_search")

    def _show_hotel_quotation(self):
        self.main_content_callback("hotel_quotation")

    def _show_service_quotation(self):
        self.main_content_callback("service_quotation")

    def _show_current_quotes(self):
        self.main_content_callback("current_quotes")

    def _show_quote_history(self):
        self.main_content_callback("quote_history")

    def _show_current_invoices(self):
        self.main_content_callback("current_invoices")

    def _show_invoice_history(self):
        self.main_content_callback("invoice_history")

    def _show_current_expenses(self):
        self.main_content_callback("current_expenses")

    def _show_expense_history(self):
        self.main_content_callback("expense_history")

    def _show_add_hotel(self):
        self.main_content_callback("add_hotel")

    def _show_update_hotel(self):
        self.main_content_callback("update_hotel")

    def _show_delete_hotel(self):
        self.main_content_callback("delete_hotel")

    def _show_add_collective_expense(self):
        self.main_content_callback("add_collective_expense")

    def _show_update_collective_expense(self):
        self.main_content_callback("update_collective_expense")

    def _show_delete_collective_expense(self):
        self.main_content_callback("delete_collective_expense")

    def _show_add_transport(self):
        self.main_content_callback("add_transport")

    def _show_update_transport(self):
        self.main_content_callback("update_transport")

    def _show_delete_transport(self):
        self.main_content_callback("delete_transport")

    def _show_income_statement(self):
        self.main_content_callback("income_statement")

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