"""
Main content GUI component
"""

import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    BUTTON_BLUE,
    BUTTON_GREEN,
    BUTTON_RED,
    FINANCIAL_EXCEL_PATH,
    MAIN_BG_COLOR,
    PANEL_BG_COLOR,
    TEXT_COLOR,
    MUTED_TEXT_COLOR,
    INPUT_BG_COLOR,
    ENTRY_FONT,
    LABEL_FONT,
)


class MainContent:
    """
    Main content area component
    """

    def __init__(self, parent):
        """
        Initialize main content area

        Args:
            parent: Parent widget
        """
        self.parent = parent
        self.current_content_type = "home"
        self._tsarakonta_process = None
        self._embedded_tsarakonta = None

        # Outer container (topbar + scrollable content)
        self._container = ctk.CTkFrame(parent, fg_color=MAIN_BG_COLOR, corner_radius=0)
        self._container.grid(row=0, column=1, sticky="nswe")
        self._container.grid_rowconfigure(0, weight=0)
        self._container.grid_rowconfigure(1, weight=1)
        self._container.grid_columnconfigure(0, weight=1)

        # Topbar with 3 icons
        self._build_topbar()

        # Scrollable main content
        self.main_scroll = ctk.CTkScrollableFrame(
            self._container, corner_radius=0, fg_color=MAIN_BG_COLOR
        )
        self.main_scroll.grid(row=1, column=0, sticky="nswe")

        self._show_welcome()

    # ── Topbar ────────────────────────────────────────────────────────────

    def _build_topbar(self):
        """Create the persistent top bar with Bienvenue / Chercher / Résa LHM."""
        topbar = ctk.CTkFrame(
            self._container,
            fg_color=PANEL_BG_COLOR,
            corner_radius=0,
            height=44,
            border_width=0,
        )
        topbar.grid(row=0, column=0, sticky="ew")
        topbar.grid_propagate(False)

        # Right-side icons container
        icons = ctk.CTkFrame(topbar, fg_color="transparent")
        icons.pack(side="right", padx=12, pady=6)

        # ── Résa LHM ──────────────────────────────────────────────────
        ctk.CTkButton(
            icons,
            text="📋 Résa LHM",
            width=110,
            height=30,
            fg_color=BUTTON_BLUE,
            hover_color="#1565C0",
            corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: self.update_content("client_page"),
        ).pack(side="right", padx=(6, 0))

        # ── Chercher ──────────────────────────────────────────────────
        ctk.CTkButton(
            icons,
            text="🔍 Chercher",
            width=110,
            height=30,
            fg_color=BUTTON_GREEN,
            hover_color="#2E7D32",
            corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._open_search_dialog,
        ).pack(side="right", padx=(6, 0))

        # ── Bienvenue ─────────────────────────────────────────────────
        ctk.CTkButton(
            icons,
            text="🏠 Bienvenue",
            width=110,
            height=30,
            fg_color=BUTTON_RED,
            hover_color="#B71C1C",
            corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._go_home,
        ).pack(side="right", padx=(0, 0))

        # ── Comptes (admin uniquement) — côté gauche ─────────────────
        try:
            from utils.auth_handler import is_admin
            if is_admin():
                ctk.CTkButton(
                    topbar,
                    text="👤 Comptes",
                    width=100,
                    height=30,
                    fg_color="#455A64",
                    hover_color="#263238",
                    corner_radius=8,
                    font=ctk.CTkFont(size=12, weight="bold"),
                    command=self._open_account_management,
                ).pack(side="left", padx=(12, 0), pady=6)
        except Exception:
            pass

    def _go_home(self):
        """Navigate to home page with confirmation if not already there."""
        if self.current_content_type in ("welcome", "home"):
            return
        if messagebox.askyesno(
            "Retour à l'accueil",
            "Voulez-vous retourner à l'accueil ?\n"
            "Les modifications non enregistrées seront perdues.",
            icon="warning",
        ):
            self.update_content("welcome")

    def _open_search_dialog(self):
        """Open the client search dialog."""
        _SearchDialog(self._container, self.update_content)

    def _open_account_management(self):
        """Open account management window (admin only)."""
        from gui.forms.account_management import AccountManagementWindow
        AccountManagementWindow(self._container)

    # ── /Topbar ───────────────────────────────────────────────────────────

    def update_content(self, content_type, **kwargs):
        """
        Update main content based on selected menu item

        Args:
            content_type (str): Type of content to display
            **kwargs: Optional data forwarded to the target view (e.g. client_to_edit)
        """
        self.current_content_type = content_type
        self._nav_kwargs = kwargs
        # Clear current content
        for widget in self.main_scroll.winfo_children():
            widget.destroy()

        if content_type in ("welcome", "home"):
            self._show_welcome()
            return

        handlers = {
            "client_form": self._show_client_page,
            "client_list": self._show_client_page,
            "client_page": self._show_client_page,
            "billing_quotes_hub_page": self._show_billing_quotes_hub_page,
            "cotation_hub_page": self._show_cotation_hub_page,
            "database_hub_page": self._show_database_hub_page,
            "hotel_form": self._show_hotel_form,
            "hotel_list": self._show_hotel_list,
            "hotel_quotation": self._show_hotel_quotation_page,
            "hotel_quotation_summary": self._show_hotel_quotation_page,
            "hotel_quotation_page": self._show_hotel_quotation_page,
            "current_quotes": self._show_client_quotes_page,
            "quote_history": self._show_client_quotes_page,
            "client_quotes_page": self._show_client_quotes_page,
            "collective_expense_quotation": self._show_collective_expense_quotation,
            "collective_expense_summary": self._show_collective_expense_summary,
            "collective_expense_page": self._show_collective_expense_page,
            "expenses_page": self._show_expenses_page,
            "transport_page": self._show_transport_page,
            "transport_db_page": self._show_transport_db_page,
            "circuit_db_page": self._show_circuit_db_page,
            "air_ticket_page": self._show_air_ticket_page,
            "air_ticket_quotation": self._show_air_ticket_quotation,
            "air_ticket_summary": self._show_air_ticket_summary,
            "air_ticket_db_list": self._show_air_ticket_db_list,
            "air_ticket_db_form": self._show_air_ticket_db_form,
            "visite_excursion_quotation": self._show_visite_excursion_quotation,
            "visite_excursion_summary": self._show_visite_excursion_summary,
            "visite_excursion_db_list": self._show_visite_excursion_db_list,
            "visite_excursion_db_form": self._show_visite_excursion_db_form,
            "collective_expense_db_list": self._show_collective_expense_db_list,
            "collective_expense_db_form": self._show_collective_expense_db_form,
            "parametrage_page": self._show_parametrage_page,
            "current_invoices": self._show_invoice_management,
            "invoice_history": self._show_invoice_management,
        }
        handler = handlers.get(content_type)
        if handler:
            handler()
        elif content_type in (
            "financial_home",
            "income_statement",
            "balance_sheet",
            "cash_flow",
            "depreciation_table",
            "fixed_assets",
            "5year_forecast",
            "12month_cash_forecast",
            "accounting_entry",
        ):
            if content_type == "financial_home":
                self._show_financial_home()
            else:
                self._show_financial_view(content_type)
        else:
            self._show_placeholder(content_type)

    def refresh(self):
        """Re-render the current view after a theme change."""
        self._container.configure(fg_color=MAIN_BG_COLOR)
        self.main_scroll.configure(fg_color=MAIN_BG_COLOR)
        self.update_content(self.current_content_type)

    def _show_welcome(self):
        """Show home page"""
        from gui.forms.home_page import HomePage

        HomePage(self.main_scroll, self.update_content)

    def _show_client_form(self, client_to_edit=None):
        """Show client form"""
        from gui.forms.client_form import ClientForm

        ClientForm(self.main_scroll, client_to_edit, self._on_client_saved)

    def _show_client_list(self):
        """Show client list"""
        from gui.forms.client_list import ClientList

        ClientList(self.main_scroll, self._edit_client, self._new_client)

    def _show_client_page(self):
        """Show combined client form + list page."""
        from gui.forms.client_page import ClientPage

        client_to_edit = getattr(self, "_nav_kwargs", {}).get("client_to_edit")
        self._nav_kwargs = {}
        ClientPage(self.main_scroll, client_to_edit=client_to_edit)

    def _show_database_hub_page(self):
        """Show dedicated database hub page."""
        from gui.forms.database_hub_page import DatabaseHubPage

        DatabaseHubPage(self.main_scroll, self.update_content)

    def _show_cotation_hub_page(self):
        """Show dedicated quotation hub page."""
        from gui.forms.cotation_hub_page import CotationHubPage

        CotationHubPage(self.main_scroll, self.update_content)

    def _show_billing_quotes_hub_page(self):
        """Show dedicated hub for invoices and client quotes."""
        from gui.forms.billing_quotes_hub_page import BillingQuotesHubPage

        BillingQuotesHubPage(self.main_scroll, self.update_content)

    def _show_hotel_form(self, hotel_to_edit=None):
        """Show hotel form"""
        from gui.forms.hotel_form import HotelForm

        HotelForm(
            self.main_scroll,
            hotel_to_edit,
            self._on_hotel_saved,
            on_back_to_db=lambda: self.update_content("database_hub_page"),
        )

    def _show_hotel_list(self):
        """Show hotel list"""
        from gui.forms.hotel_list import HotelList

        HotelList(
            self.main_scroll,
            self._edit_hotel,
            self._new_hotel,
            on_back_to_db=lambda: self.update_content("database_hub_page"),
        )

    def _show_hotel_quotation(self):
        """Show hotel quotation form"""
        from gui.forms.hotel_quotation import HotelQuotation

        HotelQuotation(self.main_scroll)

    def _show_hotel_quotation_summary(self):
        """Show hotel quotation summary"""
        from gui.forms.hotel_quotation_summary import HotelQuotationSummary

        HotelQuotationSummary(self.main_scroll)

    def _show_hotel_quotation_page(self):
        """Show combined hotel quotation + summary page."""
        from gui.forms.hotel_quotation_page import HotelQuotationPage

        HotelQuotationPage(
            self.main_scroll,
            on_back_to_cotation=lambda: self.update_content("cotation_hub_page"),
        )

    def _show_client_quotation(self):
        """Show client quotation form"""
        from gui.forms.client_quotation import ClientQuotation

        ClientQuotation(self.main_scroll)

    def _show_client_quotation_history(self):
        """Show client quotation history"""
        from gui.forms.client_quotation_history import ClientQuotationHistory

        ClientQuotationHistory(self.main_scroll)

    def _show_client_quotes_page(self):
        """Show combined client quotation + history page."""
        from gui.forms.client_quotation_page import ClientQuotationPage

        ClientQuotationPage(
            self.main_scroll,
            on_back_to_hub=lambda: self.update_content("billing_quotes_hub_page"),
        )

    def _show_collective_expense_quotation(self):
        """Show collective expense quotation form"""
        from gui.forms.collective_expense_quotation import CollectiveExpenseQuotation

        CollectiveExpenseQuotation(self.main_scroll)

    def _show_collective_expense_page(self):
        """Show combined collective expense page (form + summary)."""
        from gui.forms.collective_expense_page import CollectiveExpensePage

        CollectiveExpensePage(
            self.main_scroll,
            on_back_to_hub=lambda: self.update_content("billing_quotes_hub_page"),
        )

    def _show_expenses_page(self):
        """Show temporary empty expenses page."""
        from gui.forms.expenses_page import ExpensesPage

        ExpensesPage(
            self.main_scroll,
            on_back_to_hub=lambda: self.update_content("billing_quotes_hub_page"),
        )

    def _show_collective_expense_quotation_for_edit(self, data, row_number):
        """Show collective expense quotation form in edit mode"""
        from gui.forms.collective_expense_quotation import CollectiveExpenseQuotation

        def on_edit_done():
            # Refresh summary after edit
            self.update_content("collective_expense_summary")

        # Create form in edit mode with callback
        CollectiveExpenseQuotation(
            self.main_scroll,
            edit_data=data,
            row_number=row_number,
            callback_on_save=on_edit_done,
        )

    def _on_add_collective_expense(self):
        """Navigate to add collective expense form"""
        self.update_content("collective_expense_quotation")

    def _show_collective_expense_summary(self):
        """Show collective expense quotations summary"""
        from gui.forms.collective_expense_quotation_summary import (
            CollectiveExpenseQuotationSummary,
        )

        CollectiveExpenseQuotationSummary(
            self.main_scroll,
            callback_edit=self._show_collective_expense_quotation_for_edit,
            callback_add=self._on_add_collective_expense,
        )

    def _show_air_ticket_quotation(self):
        """Show air ticket quotation form."""
        from gui.forms.air_ticket_quotation import AirTicketQuotation

        AirTicketQuotation(self.main_scroll)

    def _show_air_ticket_page(self):
        """Show combined air ticket page (form + summary)."""
        from gui.forms.air_ticket_page import AirTicketPage

        AirTicketPage(
            self.main_scroll,
            on_back_to_cotation=lambda: self.update_content("cotation_hub_page"),
        )

    def _show_air_ticket_quotation_for_edit(self, data, row_number):
        """Show air ticket quotation in edit mode."""
        from gui.forms.air_ticket_quotation import AirTicketQuotation

        def on_edit_done():
            self.update_content("air_ticket_summary")

        AirTicketQuotation(
            self.main_scroll,
            edit_data=data,
            row_number=row_number,
            callback_on_done=on_edit_done,
        )

    def _on_add_air_ticket(self):
        """Navigate to add air ticket form."""
        self.update_content("air_ticket_page")

    def _show_air_ticket_summary(self):
        """Show air ticket quotation summary."""
        from gui.forms.air_ticket_quotation_summary import AirTicketQuotationSummary

        AirTicketQuotationSummary(
            self.main_scroll,
            callback_edit=self._show_air_ticket_quotation_for_edit,
            callback_add=self._on_add_air_ticket,
        )

    def _show_visite_excursion_quotation(self):
        """Show visite & excursion quotation form."""
        from gui.forms.visite_excursion_quotation import VisiteExcursionQuotation

        VisiteExcursionQuotation(self.main_scroll)

    def _show_visite_excursion_quotation_for_edit(self, data, row_number):
        """Show visite & excursion quotation form in edit mode."""
        from gui.forms.visite_excursion_quotation import VisiteExcursionQuotation

        def on_edit_done():
            self.update_content("visite_excursion_summary")

        VisiteExcursionQuotation(
            self.main_scroll,
            edit_data=data,
            row_number=row_number,
            callback_on_done=on_edit_done,
        )

    def _on_add_visite_excursion(self):
        """Navigate to add visite & excursion quotation."""
        self.update_content("visite_excursion_quotation")

    def _show_visite_excursion_summary(self):
        """Show visite & excursion quotation summary."""
        from gui.forms.visite_excursion_quotation_summary import (
            VisiteExcursionQuotationSummary,
        )

        VisiteExcursionQuotationSummary(
            self.main_scroll,
            callback_edit=self._show_visite_excursion_quotation_for_edit,
            callback_add=self._on_add_visite_excursion,
        )

    def _show_transport_page(self):
        """Show combined transport page (form + summary)."""
        from gui.forms.transport_page import TransportPage

        TransportPage(
            self.main_scroll,
            navigate_callback=lambda page: self.update_content(page),
            on_back_to_cotation=lambda: self.update_content("cotation_hub_page"),
        )

    def _show_transport_db_page(self):
        """Show transport DB management page."""
        from gui.forms.transport_db_page import TransportDBPage

        TransportDBPage(
            self.main_scroll,
            on_back_to_db=lambda: self.update_content("database_hub_page"),
        )

    def _show_circuit_db_page(self):
        """Show circuit DB management page."""
        from gui.forms.circuit_db_page import CircuitDBPage

        CircuitDBPage(
            self.main_scroll,
            on_back_to_db=lambda: self.update_content("database_hub_page"),
        )

    def _show_parametrage_page(self):
        """Show combined parameter page (form + summary)."""
        from gui.forms.parametrage_page import ParametragePage

        ParametragePage(self.main_scroll)

    def _show_invoice_management(self):
        """Show client invoice management."""
        from gui.forms.invoice_management import InvoiceManagement

        InvoiceManagement(
            self.main_scroll,
            on_back_to_hub=lambda: self.update_content("billing_quotes_hub_page"),
        )

    def _show_collective_expense_db_form(self, row_to_edit=None, row_number=None):
        """Show collective expense DB form."""
        from gui.forms.collective_expense_db_form import CollectiveExpenseDBForm

        def on_done():
            self.update_content("collective_expense_db_list")

        CollectiveExpenseDBForm(
            self.main_scroll,
            edit_data=row_to_edit,
            row_number=row_number,
            callback_on_done=on_done,
            on_back_to_db=lambda: self.update_content("database_hub_page"),
        )

    def _show_collective_expense_db_list(self):
        """Show collective expense DB list."""
        from gui.forms.collective_expense_db_list import CollectiveExpenseDBList

        def on_edit(row_data, row_number):
            self._show_collective_expense_db_form(row_data, row_number)

        def on_new():
            self._show_collective_expense_db_form()

        CollectiveExpenseDBList(
            self.main_scroll,
            on_edit_row=on_edit,
            on_new_row=on_new,
            on_back_to_db=lambda: self.update_content("database_hub_page"),
        )

    def _show_air_ticket_db_form(self, row_to_edit=None, row_number=None):
        """Show air ticket DB form."""
        from gui.forms.air_ticket_db_form import AirTicketDBForm

        def on_done():
            self.update_content("air_ticket_db_list")

        AirTicketDBForm(
            self.main_scroll,
            edit_data=row_to_edit,
            row_number=row_number,
            callback_on_done=on_done,
            on_back_to_db=lambda: self.update_content("database_hub_page"),
        )

    def _show_air_ticket_db_list(self):
        """Show air ticket DB list."""
        from gui.forms.air_ticket_db_list import AirTicketDBList

        def on_edit(row_data, row_number):
            self._show_air_ticket_db_form(row_data, row_number)

        def on_new():
            self._show_air_ticket_db_form()

        AirTicketDBList(
            self.main_scroll,
            on_edit_row=on_edit,
            on_new_row=on_new,
            on_back_to_db=lambda: self.update_content("database_hub_page"),
        )

    def _show_visite_excursion_db_form(self, row_to_edit=None, row_number=None):
        """Show visite & excursion DB form."""
        from gui.forms.visite_excursion_db_form import VisiteExcursionDBForm

        def on_done():
            self.update_content("visite_excursion_db_list")

        VisiteExcursionDBForm(
            self.main_scroll,
            edit_data=row_to_edit,
            row_number=row_number,
            callback_on_done=on_done,
        )

    def _show_visite_excursion_db_list(self):
        """Show visite & excursion DB list."""
        from gui.forms.visite_excursion_db_list import VisiteExcursionDBList

        def on_edit(row_data, row_number):
            self._show_visite_excursion_db_form(row_data, row_number)

        def on_new():
            self._show_visite_excursion_db_form()

        VisiteExcursionDBList(
            self.main_scroll,
            on_edit_row=on_edit,
            on_new_row=on_new,
        )

    def _show_placeholder(self, content_type):
        """Show placeholder for unimplemented features"""
        title = ctk.CTkLabel(
            self.main_scroll,
            text=f"Fonction '{content_type}' - À implémenter",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.pack(pady=40)

    def _show_financial_view(self, content_type):
        view = self._get_financial_view_data(content_type)
        if not view:
            self._show_placeholder(content_type)
            return

        title = ctk.CTkLabel(
            self.main_scroll,
            text=view["title"],
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        title.pack(pady=(20, 8))

        subtitle = ctk.CTkLabel(
            self.main_scroll,
            text=view["description"],
            font=ctk.CTkFont(size=12),
        )
        subtitle.pack(pady=(0, 16))

        for label, etat in view["actions"]:
            ctk.CTkButton(
                self.main_scroll,
                text=label,
                command=lambda e=etat: self._launch_tsarakonta(etat=e),
                height=42,
            ).pack(pady=6, padx=20, fill="x")

        if view.get("note"):
            note = ctk.CTkLabel(
                self.main_scroll,
                text=view["note"],
                font=ctk.CTkFont(size=11),
            )
            note.pack(pady=(12, 0))

    def _show_financial_home(self):
        """Embed TsaraKonta directly in the current main window."""
        missing_deps = self._check_tsarakonta_dependencies()
        if missing_deps:
            deps = ", ".join(missing_deps)
            messagebox.showerror(
                "Dépendances TsaraKonta manquantes",
                "Impossible d'afficher TsaraKonta dans la fenêtre.\n\n"
                f"Modules manquants: {deps}\n"
                "Installe-les avec:\n"
                "pip install -r requirements-financial.txt",
            )
            return

        try:
            from finances.tsarakonta.ui.main import ComptabiliteApp
        except Exception as exc:
            messagebox.showerror(
                "Erreur import TsaraKonta",
                f"Impossible de charger TsaraKonta:\n{exc}",
            )
            return

        container = ctk.CTkFrame(self.main_scroll, fg_color=MAIN_BG_COLOR)
        container.pack(fill="both", expand=True, padx=0, pady=0)

        self._embedded_tsarakonta = ComptabiliteApp(
            container,
            fichier_excel=FINANCIAL_EXCEL_PATH,
            etat_initial=None,
        )
        self._embedded_tsarakonta.pack(side="top", fill="both", expand=True)

    def _get_financial_view_data(self, content_type):
        views = {
            "income_statement": {
                "title": "Compte de résultat",
                "description": "Choisis la variante à ouvrir dans TsaraKonta.",
                "actions": [
                    ("Ouvrir - Par nature", "Compte de résultat par nature"),
                    ("Ouvrir - Par fonction", "Compte de résultat par fonction"),
                ],
            },
            "balance_sheet": {
                "title": "Bilan",
                "description": "Sélectionne la partie à afficher.",
                "actions": [
                    ("Ouvrir - Bilan actif", "Bilan actif"),
                    ("Ouvrir - Bilan passif", "Bilan passif"),
                ],
            },
            "cash_flow": {
                "title": "Tableau de flux de trésorerie",
                "description": "Choisis la méthode à ouvrir.",
                "actions": [
                    (
                        "Ouvrir - Méthode directe",
                        "Tableau de flux de trésorerie (méthode directe)",
                    ),
                    (
                        "Ouvrir - Méthode indirecte",
                        "Tableau de flux de trésorerie (méthode indirecte)",
                    ),
                ],
            },
            "accounting_entry": {
                "title": "Saisie comptable",
                "description": "Ouvre TsaraKonta pour saisir ou modifier les écritures.",
                "actions": [("Ouvrir TsaraKonta", None)],
            },
            "depreciation_table": {
                "title": "Tableau d'amortissement",
                "description": "Cette fonctionnalité n'est pas encore disponible.",
                "actions": [("Ouvrir TsaraKonta", None)],
                "note": "Le module d'amortissement sera ajouté dans TsaraKonta.",
            },
            "fixed_assets": {
                "title": "Liste des immobilisations",
                "description": "Cette fonctionnalité n'est pas encore disponible.",
                "actions": [("Ouvrir TsaraKonta", None)],
                "note": "Le module des immobilisations sera ajouté dans TsaraKonta.",
            },
            "5year_forecast": {
                "title": "Prévisionnel sur 5 ans",
                "description": "Cette fonctionnalité n'est pas encore disponible.",
                "actions": [("Ouvrir TsaraKonta", None)],
                "note": "Le module prévisionnel 5 ans sera ajouté dans TsaraKonta.",
            },
            "12month_cash_forecast": {
                "title": "Prévisionnel trésorerie 12 mois",
                "description": "Cette fonctionnalité n'est pas encore disponible.",
                "actions": [("Ouvrir TsaraKonta", None)],
                "note": "Le module trésorerie 12 mois sera ajouté dans TsaraKonta.",
            },
        }
        return views.get(content_type)

    def _launch_tsarakonta(self, etat=None):
        tsarakonta_dir = os.path.join(
            os.path.dirname(os.path.dirname(__file__)),
            "finances",
            "tsarakonta",
        )
        tsarakonta_main = os.path.join(tsarakonta_dir, "main.py")

        if not os.path.exists(tsarakonta_main):
            messagebox.showerror(
                "TsaraKonta introuvable",
                f"Le fichier n'existe pas:\n{tsarakonta_main}",
            )
            return

        if self._tsarakonta_process and self._tsarakonta_process.poll() is None:
            messagebox.showinfo(
                "TsaraKonta déjà ouvert",
                "TsaraKonta est déjà lancé.",
            )
            return

        missing_deps = self._check_tsarakonta_dependencies()
        if missing_deps:
            deps = ", ".join(missing_deps)
            messagebox.showerror(
                "Dépendances TsaraKonta manquantes",
                "Impossible de lancer TsaraKonta.\n\n"
                f"Modules manquants: {deps}\n"
                "Installe-les avec:\n"
                "pip install -r requirements-financial.txt",
            )
            return

        try:
            cmd = [sys.executable, tsarakonta_main, "--excel", FINANCIAL_EXCEL_PATH]
            if etat:
                cmd.extend(["--etat", etat])

            self._tsarakonta_process = subprocess.Popen(cmd, cwd=tsarakonta_dir)
        except Exception as exc:
            messagebox.showerror(
                "Erreur lancement TsaraKonta",
                f"Impossible de lancer TsaraKonta:\n{exc}",
            )

    @staticmethod
    def _check_tsarakonta_dependencies():
        missing = []
        for module_name in ("pandas", "openpyxl"):
            try:
                __import__(module_name)
            except ImportError:
                missing.append(module_name)
        return missing

    def _edit_client(self, client):
        """Edit a client"""
        self._show_client_form(client)

    def _new_client(self):
        """Create a new client"""
        self._show_client_form()

    def _on_client_saved(self):
        """Callback after client is saved/updated"""
        self._show_client_list()

    def _edit_hotel(self, hotel):
        """Edit a hotel"""
        self._show_hotel_form(hotel)

    def _new_hotel(self):
        """Create a new hotel"""
        self._show_hotel_form()

    def _on_hotel_saved(self):
        """Callback after hotel is saved/updated"""
        self._show_hotel_list()


class _SearchDialog(tk.Toplevel):
    """Dialog de recherche de dossier client (Chercher)."""

    def __init__(self, parent, navigate_callback):
        super().__init__(parent)
        self.title("Rechercher un dossier client")
        self.geometry("560x420")
        self.configure(bg=PANEL_BG_COLOR)
        self.resizable(False, False)
        self.transient(parent)
        self._navigate = navigate_callback
        self._results = []

        self.after(0, self._safe_grab)
        self._build_ui()

    def _safe_grab(self):
        try:
            self.wait_visibility()
            self.grab_set()
        except Exception:
            pass

    def _build_ui(self):
        # Title
        ctk.CTkLabel(
            self,
            text="Recherche de dossier client",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(pady=(16, 4))

        ctk.CTkLabel(
            self,
            text="Entrez au moins le N° de dossier ou le Nom pour rechercher.",
            font=ctk.CTkFont(size=12),
            text_color=MUTED_TEXT_COLOR,
        ).pack(pady=(0, 12))

        fields_frame = ctk.CTkFrame(self, fg_color=PANEL_BG_COLOR)
        fields_frame.pack(fill="x", padx=20, pady=(0, 8))

        def _row(parent, label, attr):
            row = tk.Frame(parent, bg=PANEL_BG_COLOR)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label, font=("Poppins", 10), fg=TEXT_COLOR,
                     bg=PANEL_BG_COLOR, width=16, anchor="w").pack(side="left")
            var = tk.StringVar()
            entry = tk.Entry(row, textvariable=var, font=("Poppins", 10),
                             bg=INPUT_BG_COLOR, fg=TEXT_COLOR, relief="flat",
                             highlightthickness=1, highlightbackground="#9EC7CF")
            entry.pack(side="left", fill="x", expand=True)
            setattr(self, attr, var)
            return entry

        self._e_dossier = _row(fields_frame, "N° Dossier", "_var_dossier")
        self._e_nom = _row(fields_frame, "Nom client", "_var_nom")
        self._e_email = _row(fields_frame, "Email", "_var_email")
        self._e_dossier.focus_set()
        self._e_dossier.bind("<Return>", lambda e: self._search())
        self._e_nom.bind("<Return>", lambda e: self._search())

        ctk.CTkButton(
            self, text="🔍 Rechercher", width=160, height=32,
            fg_color=BUTTON_BLUE, hover_color="#1565C0",
            corner_radius=8, font=ctk.CTkFont(size=12, weight="bold"),
            command=self._search,
        ).pack(pady=(0, 10))

        # Results treeview
        tree_frame = tk.Frame(self, bg=PANEL_BG_COLOR)
        tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        cols = ("dossier", "nom", "email", "statut", "arrivee")
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=7)
        headers = {"dossier": "N° Dossier", "nom": "Nom", "email": "Email",
                   "statut": "Statut", "arrivee": "Arrivée"}
        widths = {"dossier": 120, "nom": 140, "email": 140, "statut": 90, "arrivee": 90}
        for c in cols:
            self._tree.heading(c, text=headers[c])
            self._tree.column(c, width=widths[c], minwidth=60)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._tree.bind("<Double-1>", self._open_selected)

        ctk.CTkButton(
            self, text="Ouvrir le dossier", width=160, height=30,
            fg_color=BUTTON_GREEN, hover_color="#2E7D32",
            corner_radius=8, font=ctk.CTkFont(size=12, weight="bold"),
            command=self._open_selected,
        ).pack(pady=(0, 12))

    def _search(self):
        from utils.excel_handler import load_all_clients
        dossier = self._var_dossier.get().strip().lower()
        nom = self._var_nom.get().strip().lower()
        email = self._var_email.get().strip().lower()

        if not dossier and not nom and not email:
            messagebox.showwarning(
                "Critère manquant",
                "Veuillez entrer au moins un critère de recherche.",
                parent=self,
            )
            return

        try:
            clients = load_all_clients()
        except Exception:
            clients = []

        self._results = [
            c for c in clients
            if (not dossier or dossier in str(c.get("numero_dossier", "")).lower()
                or dossier in str(c.get("ref_client", "")).lower())
            and (not nom or nom in str(c.get("nom", "")).lower())
            and (not email or email in str(c.get("email", "")).lower())
        ]

        for item in self._tree.get_children():
            self._tree.delete(item)

        for c in self._results:
            self._tree.insert("", "end", values=(
                c.get("numero_dossier") or c.get("ref_client", ""),
                c.get("nom", ""),
                c.get("email", ""),
                c.get("statut") or "En cours",
                c.get("date_arrivee", ""),
            ))

        if not self._results:
            messagebox.showinfo("Aucun résultat", "Aucun dossier trouvé.", parent=self)

    def _open_selected(self, event=None):
        selection = self._tree.selection()
        if not selection:
            return
        idx = self._tree.index(selection[0])
        if idx >= len(self._results):
            return
        client = self._results[idx]
        self.destroy()
        # Navigate to client form in edit mode via client_page
        # We pass the client via a global signal if supported, otherwise go to list
        self._navigate("client_page")
