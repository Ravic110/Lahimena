"""
Client form GUI component - Version améliorée avec nouveaux champs
"""

import calendar
import getpass
import os
import re
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

try:
    import customtkinter as ctk
except Exception:
    ctk = None

from config import (
    AGES_ENFANTS,
    CIRCUITS,
    CODE_TO_COUNTRY,
    COUNTRY_PHONE_MAP,
    DEFAULT_COUNTRY,
    DEFAULT_PHONE_CODE,
    FORFAITS,
    HOTEL_ARRIVAL_TYPES,
    MAIN_BG_COLOR,
    MUTED_TEXT_COLOR,
    BUTTON_BLUE,
    BUTTON_ORANGE,
    BUTTON_RED,
    BUTTON_GREEN,
    PANEL_BG_COLOR,
    CARD_BG_COLOR,
    LABEL_FONT,
    ENTRY_FONT,
    INPUT_BG_COLOR,
    PERIODES,
    PHONE_CODES,
    RESTAURATIONS,
    TEXT_COLOR,
    TITLE_FONT,
    TYPE_CHAMBRES,
    TYPE_HEBERGEMENTS,
)
from gui.ui_style import (
    action_button,
    configure_combobox_style,
    create_card,
    muted_label,
    row_two_columns,
    styled_entry,
    styled_label,
)
from models.client_data import ClientData
from utils.excel_handler import (
    load_all_circuits,
    load_all_hotels,
    load_circuit_catalog,
    save_client_to_excel,
)
from utils.logger import logger
from utils.validators import validate_email, validate_phone_number


class CalendarDialog(tk.Toplevel):
    """
    Custom calendar dialog for date selection
    """

    def __init__(self, parent, title="Choisir une date"):
        super().__init__(parent)
        self.title(title)
        self.geometry("350x400")
        self.configure(bg=MAIN_BG_COLOR)
        self.resizable(False, False)

        # Center window
        self.transient(parent)
        self.after(0, self._safe_grab)

        self.selected_date = None
        self.current_month = datetime.now().month
        self.current_year = datetime.now().year

        self._create_widgets()

    def _safe_grab(self):
        """Grab focus once the window is viewable."""
        try:
            self.wait_visibility()
            self.grab_set()
        except tk.TclError:
            pass

    def _create_widgets(self):
        """Create calendar widgets"""
        # Header frame
        header_frame = tk.Frame(self, bg=MAIN_BG_COLOR)
        header_frame.pack(fill="x", padx=10, pady=10)

        # Previous month button
        tk.Button(
            header_frame,
            text="◀",
            command=self._prev_month,
            bg=BUTTON_BLUE,
            fg="white",
            font=("Poppins", 12, "bold"),
            width=3,
        ).pack(side="left")

        # Month/Year label
        self.month_year_label = tk.Label(
            header_frame,
            text="",
            font=("Poppins", 14, "bold"),
            bg=MAIN_BG_COLOR,
            fg=TEXT_COLOR,
        )
        self.month_year_label.pack(side="left", expand=True)

        # Next month button
        tk.Button(
            header_frame,
            text="▶",
            command=self._next_month,
            bg=BUTTON_BLUE,
            fg="white",
            font=("Poppins", 12, "bold"),
            width=3,
        ).pack(side="right")

        # Calendar frame
        self.calendar_frame = tk.Frame(self, bg=MAIN_BG_COLOR)
        self.calendar_frame.pack(padx=10, pady=10)

        self._show_calendar()

        # Button frame
        btn_frame = tk.Frame(self, bg=MAIN_BG_COLOR)
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="Aujourd'hui",
            command=self._select_today,
            bg=BUTTON_GREEN,
            fg="white",
            font=("Poppins", 10),
            width=12,
        ).pack(side="left", padx=5)

        tk.Button(
            btn_frame,
            text="Annuler",
            command=self.destroy,
            bg=BUTTON_RED,
            fg="white",
            font=("Poppins", 10),
            width=12,
        ).pack(side="left", padx=5)

    def _show_calendar(self):
        """Display calendar for current month/year"""
        # Clear previous calendar
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()

        # Update month/year label
        months_fr = [
            "Janvier",
            "Février",
            "Mars",
            "Avril",
            "Mai",
            "Juin",
            "Juillet",
            "Août",
            "Septembre",
            "Octobre",
            "Novembre",
            "Décembre",
        ]
        self.month_year_label.config(
            text=f"{months_fr[self.current_month - 1]} {self.current_year}"
        )

        # Day headers
        days = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]
        for i, day in enumerate(days):
            tk.Label(
                self.calendar_frame,
                text=day,
                font=("Poppins", 10, "bold"),
                bg=BUTTON_BLUE,
                fg="white",
                width=5,
            ).grid(row=0, column=i, padx=1, pady=1, sticky="nsew")

        # Get calendar data
        cal = calendar.monthcalendar(self.current_year, self.current_month)

        # Today's date
        today = datetime.now()

        # Create day buttons
        for week_num, week in enumerate(cal, start=1):
            for day_num, day in enumerate(week):
                if day == 0:
                    # Empty cell
                    tk.Label(
                        self.calendar_frame, text="", bg=MAIN_BG_COLOR, width=5
                    ).grid(row=week_num, column=day_num, padx=1, pady=1)
                else:
                    # Check if it's today
                    is_today = (
                        day == today.day
                        and self.current_month == today.month
                        and self.current_year == today.year
                    )

                    bg_color = BUTTON_GREEN if is_today else INPUT_BG_COLOR
                    fg_color = "white" if is_today else TEXT_COLOR

                    btn = tk.Button(
                        self.calendar_frame,
                        text=str(day),
                        bg=bg_color,
                        fg=fg_color,
                        font=("Poppins", 10),
                        width=5,
                        command=lambda d=day: self._select_date(d),
                    )
                    btn.grid(
                        row=week_num, column=day_num, padx=1, pady=1, sticky="nsew"
                    )

    def _prev_month(self):
        """Go to previous month"""
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self._show_calendar()

    def _next_month(self):
        """Go to next month"""
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self._show_calendar()

    def _select_date(self, day):
        """Select a specific date"""
        self.selected_date = datetime(self.current_year, self.current_month, day)
        self.destroy()

    def _select_today(self):
        """Select today's date"""
        today = datetime.now()
        self.selected_date = today
        self.destroy()


class ClientForm:
    """
    Client request form component with extended fields
    """

    def __init__(self, parent, client_to_edit=None, on_save_callback=None):
        """
        Initialize client form

        Args:
            parent: Parent widget
            client_to_edit: Optional client dict to edit
            on_save_callback: Optional callback after saving
        """
        self.parent = parent
        self.client_to_edit = client_to_edit
        self.on_save_callback = on_save_callback
        self.client_data = {}
        self.canvas = None
        self.canvas_window = None
        self.form_frame = None
        self.main_frame = None
        self.city_options = self._load_city_options()
        self.circuit_catalog = self._load_circuit_catalog()
        self.circuit_map = {
            circuit.get("nom", ""): circuit
            for circuit in self.circuit_catalog
            if circuit.get("nom")
        }
        self.circuit_options = self._load_circuit_options()
        self.itinerary_rows = []
        self._itin_widget_rows = []
        self._itin_canvas = None

        self._create_form()

    def _load_city_options(self):
        """Load unique city options from hotel database"""
        try:
            hotels = load_all_hotels()
            cities = sorted(
                {
                    (hotel.get("lieu") or "").strip()
                    for hotel in hotels
                    if hotel.get("lieu")
                }
            )
            return cities
        except Exception as e:
            logger.warning(f"Failed to load city options: {e}")
            return []

    def _load_circuit_options(self):
        """Load circuit options from the Circuits sheet when available"""
        if self.circuit_catalog:
            return [circuit.get("nom", "") for circuit in self.circuit_catalog if circuit.get("nom")]
        try:
            circuits = load_all_circuits()
            return circuits if circuits else CIRCUITS
        except Exception as e:
            logger.warning(f"Failed to load circuits: {e}")
            return CIRCUITS

    def _load_circuit_catalog(self):
        """Load detailed circuits metadata from Excel when available."""
        try:
            return load_circuit_catalog()
        except Exception as e:
            logger.warning(f"Failed to load circuit catalog: {e}")
            return []

    def _apply_placeholder(self, entry, text):
        if entry is None:
            return
        try:
            entry.configure(placeholder_text=text)
            entry._placeholder_text = text
            return
        except Exception:
            pass

        entry._placeholder_text = text
        entry._placeholder_active = False
        placeholder_color = MUTED_TEXT_COLOR
        normal_color = TEXT_COLOR

        def _show_placeholder():
            if entry.get():
                return
            entry._placeholder_active = True
            entry.config(fg=placeholder_color)
            entry.delete(0, tk.END)
            entry.insert(0, text)

        def _hide_placeholder():
            if not getattr(entry, "_placeholder_active", False):
                return
            entry._placeholder_active = False
            entry.delete(0, tk.END)
            entry.config(fg=normal_color)

        def _on_focus_in(event):
            _hide_placeholder()

        def _on_focus_out(event):
            if not entry.get():
                _show_placeholder()

        entry.bind("<FocusIn>", _on_focus_in, add=True)
        entry.bind("<FocusOut>", _on_focus_out, add=True)
        entry._placeholder_show = _show_placeholder
        entry._placeholder_hide = _hide_placeholder
        _show_placeholder()

    def _clear_placeholder(self, entry):
        if hasattr(entry, "_placeholder_hide"):
            entry._placeholder_hide()

    def _restore_placeholder(self, entry):
        if hasattr(entry, "_placeholder_show"):
            entry._placeholder_show()

    def _get_entry_value(self, entry):
        if entry is None:
            return ""
        if getattr(entry, "_placeholder_active", False):
            return ""
        value = entry.get()
        if getattr(entry, "_placeholder_text", None) == value:
            return ""
        return value

    def _make_calendar_badge(self, parent, entry_widget, label_text="📅"):
        """Create a simple calendar icon button."""
        badge = tk.Canvas(
            parent,
            width=28,
            height=28,
            bg=PANEL_BG_COLOR,
            highlightthickness=0,
            cursor="hand2",
        )
        badge.create_text(
            14, 14,
            text="📅",
            font=("Poppins", 13),
        )
        badge.bind("<Button-1>", lambda e: self._open_calendar(entry_widget))
        return badge

    def _create_form(self):
        """Create the client form with scrollable area for many fields"""
        # Clear parent
        for widget in self.parent.winfo_children():
            widget.destroy()

        # Create main container that will expand
        container = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        container.pack(fill="both", expand=True, padx=0, pady=0)
        self.container = container

        top_bar = tk.Frame(container, bg=MAIN_BG_COLOR)
        top_bar.pack(fill="x", pady=(6, 4), padx=16)

        def _top_action(parent, label, icon):
            item = tk.Frame(parent, bg=MAIN_BG_COLOR)
            item.pack(side="left", padx=6)
            canvas = tk.Canvas(
                item,
                width=18,
                height=18,
                bg=MAIN_BG_COLOR,
                highlightthickness=0,
            )
            canvas.create_oval(1, 1, 17, 17, fill=BUTTON_RED, outline=BUTTON_RED)
            canvas.create_text(
                9,
                9,
                text=icon,
                fill="white",
                font=("Poppins", 8, "bold"),
            )
            canvas.pack()
            tk.Label(
                item,
                text=label,
                font=("Poppins", 7),
                fg=MUTED_TEXT_COLOR,
                bg=MAIN_BG_COLOR,
            ).pack()

        use_ctk = ctk is not None and hasattr(ctk, "CTkFrame")
        if use_ctk:
            page_card = ctk.CTkFrame(
                container,
                fg_color=CARD_BG_COLOR,
                corner_radius=22,
                border_width=1,
                border_color="#C9DDE3",
            )
        else:
            page_card = tk.Frame(
                container,
                bg=CARD_BG_COLOR,
                highlightbackground="#C9DDE3",
                highlightthickness=1,
                bd=0,
            )
        page_card.pack(fill="both", expand=True, padx=16, pady=(4, 16))
        if use_ctk:
            page_inner = ctk.CTkFrame(page_card, fg_color="transparent")
        else:
            page_inner = tk.Frame(page_card, bg=CARD_BG_COLOR)
        page_inner.pack(fill="both", expand=True, padx=12, pady=10)

        header = tk.Frame(page_inner, bg=CARD_BG_COLOR)
        header.pack(fill="x", pady=(4, 10), padx=8)
        header.grid_columnconfigure(0, weight=1)

        title_block = tk.Frame(header, bg=CARD_BG_COLOR)
        title_block.grid(row=0, column=0)
        title_text = (
            "MODIFICATION DEMANDE CLIENT"
            if self.client_to_edit
            else "FORMULAIRE DEMANDES CLIENT"
        )
        tk.Label(
            title_block,
            text=title_text.upper(),
            font=("Poppins", 19, "bold"),
            fg="#667780",
            bg=CARD_BG_COLOR,
        ).pack(anchor="center")

        try:
            from utils.auth_handler import get_current_user
            _u = get_current_user()
            creator = f"{_u['username']} ({_u['role']})" if _u else (os.getenv("LHM_USER") or getpass.getuser() or "-")
        except Exception:
            creator = os.getenv("LHM_USER") or getpass.getuser() or "-"
        current_datetime = datetime.now().strftime("%d/%m/%Y à %H:%M:%S")
        tk.Label(
            title_block,
            text=f"Création du dossier : {current_datetime}  |  Par : {creator}",
            font=("Poppins", 8),
            fg=MUTED_TEXT_COLOR,
            bg=CARD_BG_COLOR,
        ).pack(anchor="center")
        tk.Frame(page_inner, bg="#D12E2E", height=2).pack(
            fill="x", padx=18, pady=(2, 12)
        )

        current_date = datetime.now().strftime("%d/%m/%Y")

        # Hidden date field for persistence
        self.entry_date_jour = styled_entry(page_inner, readonly=True, width=40)
        self.entry_date_jour.configure(state="normal")
        self.entry_date_jour.insert(0, current_date)
        self.entry_date_jour.configure(state="readonly")

        configure_combobox_style(page_inner)

        # No internal scrollbar: rely on MainContent scroll only.
        self.main_frame = tk.Frame(page_inner, bg=CARD_BG_COLOR)
        self.main_frame.pack(fill="both", expand=True, padx=0, pady=0)

        content = tk.Frame(self.main_frame, bg=CARD_BG_COLOR)
        content.pack(fill="both", expand=True, padx=14, pady=6)
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)

        left_col = tk.Frame(content, bg=CARD_BG_COLOR)
        right_col = tk.Frame(content, bg=CARD_BG_COLOR)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(6, 10))
        right_col.grid(row=0, column=1, sticky="nsew", padx=(10, 6))

        def two_column_row(parent, left_label, left_widget, right_label, right_widget):
            row_two_columns(parent, left_label, left_widget, right_label, right_widget)

        def full_width_row(parent, label, widget_builder):
            row = tk.Frame(parent, bg=PANEL_BG_COLOR)
            row.pack(fill="x", pady=(0, 6))
            if label:
                tk.Label(
                    row,
                    text=label,
                    font=LABEL_FONT,
                    fg=TEXT_COLOR,
                    bg=PANEL_BG_COLOR,
                ).pack(anchor="w")
            widget_builder(row)

        def inline_row(parent, label, widget_builder, pady=(0, 6)):
            """Label and widget on the same horizontal line."""
            row = tk.Frame(parent, bg=PANEL_BG_COLOR)
            row.pack(fill="x", pady=pady)
            tk.Label(
                row,
                text=label,
                font=LABEL_FONT,
                fg=TEXT_COLOR,
                bg=PANEL_BG_COLOR,
            ).pack(side="left", padx=(0, 8))
            widget_builder(row)

        def three_column_row(parent, columns):
            row = tk.Frame(parent, bg=PANEL_BG_COLOR)
            row.pack(fill="x", pady=(0, 6))
            for index, (label, widget_builder) in enumerate(columns):
                if index == 0:
                    padx = (0, 8)
                elif index == 1:
                    padx = (8, 8)
                else:
                    padx = (8, 0)
                column_frame = tk.Frame(row, bg=PANEL_BG_COLOR)
                column_frame.pack(
                    side="left", fill="x", expand=True, padx=padx
                )
                if label:
                    tk.Label(
                        column_frame,
                        text=label,
                        font=LABEL_FONT,
                        fg=TEXT_COLOR,
                        bg=PANEL_BG_COLOR,
                    ).pack(anchor="w")
                if widget_builder:
                    widget_builder(column_frame)
                elif label:
                    tk.Frame(column_frame, bg=PANEL_BG_COLOR, height=24).pack(
                        fill="x"
                    )

        # ===== CARD: INFOS CLIENTS =====
        info_card = create_card(
            left_col,
            title=None,
            tabs=None,
            show_controls=False,
        )

        # ── Onglets cliquables ──────────────────────────────────────────
        _tab_header = tk.Frame(info_card, bg=PANEL_BG_COLOR)
        _tab_header.pack(fill="x")

        _lbl_tab_clients = tk.Label(
            _tab_header, text="Clients", bg=BUTTON_RED, fg="white",
            font=("Poppins", 10, "bold"), padx=10, pady=3, cursor="hand2",
        )
        _lbl_tab_clients.pack(side="left", padx=(0, 6))
        _lbl_tab_compl = tk.Label(
            _tab_header, text="Compléments", bg=BUTTON_BLUE, fg="white",
            font=("Poppins", 10, "bold"), padx=10, pady=3, cursor="hand2",
        )
        _lbl_tab_compl.pack(side="left")
        tk.Frame(info_card, bg=BUTTON_RED, height=2).pack(fill="x", pady=(6, 8))

        clients_panel = tk.Frame(info_card, bg=PANEL_BG_COLOR)
        clients_panel.pack(fill="x")
        complements_panel = tk.Frame(info_card, bg=PANEL_BG_COLOR)

        def _switch_tab(tab):
            if tab == "clients":
                _lbl_tab_clients.configure(bg=BUTTON_RED)
                _lbl_tab_compl.configure(bg=BUTTON_BLUE)
                complements_panel.pack_forget()
                clients_panel.pack(fill="x")
            else:
                _lbl_tab_clients.configure(bg=BUTTON_BLUE)
                _lbl_tab_compl.configure(bg=BUTTON_RED)
                clients_panel.pack_forget()
                complements_panel.pack(fill="x")

        _lbl_tab_clients.bind("<Button-1>", lambda e: _switch_tab("clients"))
        _lbl_tab_compl.bind("<Button-1>", lambda e: _switch_tab("complements"))

        def _field_type_client(parent):
            values = ["Mr", "Mme", "CIE"]
            if ctk is not None:
                self.combo_type_client = ctk.CTkComboBox(
                    parent,
                    values=values,
                    state="readonly",
                    width=90,
                    height=28,
                    fg_color=INPUT_BG_COLOR,
                    text_color=TEXT_COLOR,
                    border_color="#9EC7CF",
                    corner_radius=14,
                    font=ENTRY_FONT,
                    dropdown_fg_color=INPUT_BG_COLOR,
                    dropdown_text_color=TEXT_COLOR,
                    dropdown_hover_color=BUTTON_GREEN,
                )
                self.combo_type_client.set(values[0])
            else:
                self.combo_type_client = ttk.Combobox(
                    parent,
                    values=values,
                    state="readonly",
                    width=7,
                )
                self.combo_type_client.current(0)
            self.combo_type_client.pack(fill="x")

        def _field_nom(parent):
            self.entry_nom = styled_entry(parent)
            self._apply_placeholder(self.entry_nom, "Nom")
            self.entry_nom.pack(fill="x")

            def _force_upper(event=None):
                try:
                    # Ne rien faire si le placeholder est actif (tk.Entry)
                    if getattr(self.entry_nom, "_placeholder_active", False):
                        return
                    cur = self.entry_nom.get()
                    # Pour CTK, get() retourne "" quand placeholder affiché → pas de risque
                    upper = cur.upper()
                    if cur != upper:
                        try:
                            pos = self.entry_nom.index(tk.INSERT)
                        except Exception:
                            pos = tk.END
                        self.entry_nom.delete(0, tk.END)
                        self.entry_nom.insert(0, upper)
                        try:
                            self.entry_nom.icursor(pos)
                        except Exception:
                            pass
                except Exception:
                    pass
                self._update_rooming_identity()

            self.entry_nom.bind("<KeyRelease>", lambda e: _force_upper())

        def _field_prenom(parent):
            self.entry_prenom = styled_entry(parent)
            self._apply_placeholder(self.entry_prenom, "Prénom")
            self.entry_prenom.pack(fill="x")

        def _field_email(parent):
            self.entry_email = styled_entry(parent)
            self.entry_email.pack(side="left", fill="x", expand=True)

        def _field_mobile(parent):
            country_names = list(COUNTRY_PHONE_MAP.keys())

            # StringVar holding the dial code (e.g. "+261") – used for save/load
            self._code_pays_var = tk.StringVar(value=DEFAULT_PHONE_CODE)

            def _on_country_selected(event=None):
                country = self.combo_pays.get()
                code = COUNTRY_PHONE_MAP.get(country, DEFAULT_PHONE_CODE)
                self._code_pays_var.set(code)
                self.whatsapp_code_var.set(code)

            phone_frame = tk.Frame(parent, bg=PANEL_BG_COLOR)
            phone_frame.pack(side="left", fill="x", expand=True)

            # ── Pays (country name dropdown) ──────────────────────────────
            # Always use ttk.Combobox for the dropdown: CTkComboBox loses focus
            # immediately and the list disappears before the user can pick.
            pays_col = tk.Frame(phone_frame, bg=PANEL_BG_COLOR)
            pays_col.pack(side="left", padx=(0, 4))
            tk.Label(
                pays_col, text="Pays",
                font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR,
            ).pack(anchor="w")
            style = ttk.Style()
            style.configure(
                "Phone.TCombobox",
                font=ENTRY_FONT,
                foreground=TEXT_COLOR,
                fieldbackground=INPUT_BG_COLOR,
                background=INPUT_BG_COLOR,
                arrowcolor=TEXT_COLOR,
                padding=4,
                relief="flat",
            )
            style.map(
                "Phone.TCombobox",
                foreground=[("readonly", TEXT_COLOR)],
                fieldbackground=[("readonly", INPUT_BG_COLOR)],
                selectforeground=[("readonly", TEXT_COLOR)],
                selectbackground=[("readonly", INPUT_BG_COLOR)],
            )
            self.combo_pays = ttk.Combobox(
                pays_col,
                values=country_names,
                width=16,
                state="readonly",
                style="Phone.TCombobox",
            )
            self.combo_pays.set(DEFAULT_COUNTRY)
            self.combo_pays.bind("<<ComboboxSelected>>", _on_country_selected)
            self.combo_pays.pack()

            # ── Code pays (readonly, auto-filled) ────────────────────────
            code_col = tk.Frame(phone_frame, bg=PANEL_BG_COLOR)
            code_col.pack(side="left", padx=(0, 4))
            tk.Label(
                code_col, text="Code",
                font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR,
            ).pack(anchor="w")
            if ctk is not None:
                self._code_pays_entry = ctk.CTkEntry(
                    code_col,
                    width=60,
                    height=28,
                    fg_color=INPUT_BG_COLOR,
                    text_color=TEXT_COLOR,
                    border_color="#9EC7CF",
                    corner_radius=14,
                    font=ENTRY_FONT,
                    textvariable=self._code_pays_var,
                    state="readonly",
                )
            else:
                self._code_pays_entry = tk.Entry(
                    code_col,
                    width=6,
                    font=ENTRY_FONT,
                    bg=INPUT_BG_COLOR,
                    fg=TEXT_COLOR,
                    textvariable=self._code_pays_var,
                    state="readonly",
                    relief="flat",
                    highlightthickness=1,
                    highlightbackground="#9EC7CF",
                    bd=0,
                )
            self._code_pays_entry.pack()

            # Proxy: entry_code_pays.get() → _code_pays_var.get()
            class _CodeProxy:
                def __init__(self_, var):
                    self_._var = var
                def get(self_):
                    return self_._var.get()
                def set(self_, val):
                    self_._var.set(val)
                    country = CODE_TO_COUNTRY.get(val, "")
                    if country:
                        self.combo_pays.set(country)
            self.entry_code_pays = _CodeProxy(self._code_pays_var)

            # ── Numéro ───────────────────────────────────────────────────
            num_col = tk.Frame(phone_frame, bg=PANEL_BG_COLOR)
            num_col.pack(side="left", fill="x", expand=True)
            tk.Label(
                num_col, text="Numéro",
                font=("Poppins", 9), fg=MUTED_TEXT_COLOR, bg=PANEL_BG_COLOR,
            ).pack(anchor="w")
            self.entry_telephone = styled_entry(num_col)
            self.entry_telephone.pack(fill="x")

            # WhatsApp mirrors (hidden, kept for data compatibility)
            self.whatsapp_code_var = tk.StringVar(value=DEFAULT_PHONE_CODE)
            self.whatsapp_number_var = tk.StringVar()
            self.combo_whatsapp_code = ttk.Combobox(
                phone_frame, values=PHONE_CODES, width=8,
                state="readonly", textvariable=self.whatsapp_code_var,
            )
            self.entry_whatsapp = styled_entry(phone_frame)
            self.entry_whatsapp.configure(textvariable=self.whatsapp_number_var)
            self.entry_telephone.bind(
                "<KeyRelease>",
                lambda e: self.whatsapp_number_var.set(self.entry_telephone.get()),
            )

        title_row = tk.Frame(clients_panel, bg=PANEL_BG_COLOR)
        title_row.pack(fill="x", pady=(0, 6))
        title_col = tk.Frame(title_row, bg=PANEL_BG_COLOR)
        title_col.pack(side="left", padx=(0, 8))
        _field_type_client(title_col)

        nom_col = tk.Frame(title_row, bg=PANEL_BG_COLOR)
        nom_col.pack(side="left", fill="x", expand=True, padx=(8, 8))
        _field_nom(nom_col)

        prenom_col = tk.Frame(title_row, bg=PANEL_BG_COLOR)
        prenom_col.pack(side="left", fill="x", expand=True, padx=(8, 6))
        _field_prenom(prenom_col)

        info_btn_col = tk.Frame(title_row, bg=PANEL_BG_COLOR)
        info_btn_col.pack(side="left", padx=(0, 0))
        info_canvas = tk.Canvas(
            info_btn_col,
            width=22,
            height=22,
            bg=PANEL_BG_COLOR,
            highlightthickness=0,
            cursor="hand2",
        )
        info_canvas.create_oval(1, 1, 21, 21, fill=BUTTON_BLUE, outline=BUTTON_BLUE)
        info_canvas.create_text(
            11, 11, text="i", fill="white", font=("Poppins", 10, "bold")
        )
        info_canvas.pack(pady=(4, 0))
        info_canvas.bind(
            "<Button-1>",
            lambda e: __import__("tkinter.messagebox", fromlist=["showinfo"]).showinfo(
                "Information client",
                "Renseignez ici les informations personnelles du client.\n"
                "Les champs marqués * sont obligatoires.",
            ),
        )

        # Email + Mobile alignés avec le champ Nom (spacer = largeur de title_col)
        email_mobile_row = tk.Frame(clients_panel, bg=PANEL_BG_COLOR)
        email_mobile_row.pack(fill="x")
        spacer = tk.Frame(email_mobile_row, bg=PANEL_BG_COLOR, width=106)
        spacer.pack_propagate(False)
        spacer.pack(side="left")
        email_mobile_content = tk.Frame(email_mobile_row, bg=PANEL_BG_COLOR)
        email_mobile_content.pack(side="left", fill="x", expand=True)
        inline_row(email_mobile_content, "Adresse e-mail *", _field_email, pady=(0, 3))
        inline_row(email_mobile_content, "Mobile/Whatsapp *", _field_mobile, pady=(0, 6))

        self.entry_prenom.bind(
            "<KeyRelease>", lambda e: self._update_rooming_identity()
        )

        def _field_ref(parent):
            self.entry_ref_client = styled_entry(parent, readonly=True)
            self._apply_placeholder(self.entry_ref_client, "Auto-généré")
            self.entry_ref_client.pack(fill="x")

        def _field_dossier(parent):
            self.entry_numero_dossier = styled_entry(parent, readonly=True)
            self._apply_placeholder(self.entry_numero_dossier, "Auto-généré")
            self.entry_numero_dossier.pack(fill="x")

        two_column_row(
            clients_panel,
            "Référence client",
            _field_ref,
            "Numéro de dossier",
            _field_dossier,
        )

        # ── Statut dossier ───────────────────────────────────────────────
        STATUTS = ["En cours", "Accepté", "En circuit", "Annulé"]
        STATUT_COLORS = {
            "En cours":   "#00BCD4",
            "Accepté":    "#4CAF50",
            "En circuit": "#FF9800",
            "Annulé":     "#F44336",
        }

        statut_row = tk.Frame(clients_panel, bg=PANEL_BG_COLOR)
        statut_row.pack(fill="x", pady=(4, 0))
        tk.Label(
            statut_row, text="Statut", font=LABEL_FONT,
            fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(side="left", padx=(0, 8))

        self._statut_var = tk.StringVar(value="En cours")
        self._statut_btns = {}

        for s in STATUTS:
            btn = tk.Label(
                statut_row, text=s, font=("Poppins", 9, "bold"),
                bg=PANEL_BG_COLOR, fg=TEXT_COLOR,
                padx=10, pady=3, cursor="hand2",
                relief="flat",
                highlightbackground=STATUT_COLORS[s], highlightthickness=1,
            )
            btn.pack(side="left", padx=(0, 6))
            btn.bind("<Button-1>", lambda e, v=s: self._apply_statut(v))
            self._statut_btns[s] = btn

        self._apply_statut("En cours")

        # ── Contenu onglet Compléments ──────────────────────────────────
        COMPAGNIES = [
            "Air Madagascar", "Air France", "Corsair", "Ethiopian Airlines",
            "Kenya Airways", "Réunion Air Transport", "Air Austral",
            "Turkish Airlines", "Emirates", "Autre",
        ]

        def _field_heure_arrivee(parent):
            self.entry_heure_arrivee = styled_entry(parent, width=8)
            self._apply_placeholder(self.entry_heure_arrivee, "hh:mm")
            self.entry_heure_arrivee.pack(fill="x")

        def _field_heure_depart(parent):
            self.entry_heure_depart = styled_entry(parent, width=8)
            self._apply_placeholder(self.entry_heure_depart, "hh:mm")
            self.entry_heure_depart.pack(fill="x")

        def _field_compagnie(parent):
            if ctk:
                self.combo_compagnie = ctk.CTkComboBox(
                    parent, values=COMPAGNIES, state="readonly",
                    font=ENTRY_FONT, corner_radius=8,
                    fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                    button_color="#C9DDE3", button_hover_color=BUTTON_BLUE,
                    border_color="#C9DDE3", border_width=1,
                    dropdown_fg_color=INPUT_BG_COLOR, dropdown_text_color=TEXT_COLOR,
                    dropdown_hover_color="#E8F4F8", height=34,
                )
            else:
                self.combo_compagnie = ttk.Combobox(
                    parent, values=COMPAGNIES, state="readonly", font=ENTRY_FONT,
                )
            self.combo_compagnie.pack(fill="x")

        def _field_aeroport(parent):
            self.entry_aeroport = styled_entry(parent)
            self._apply_placeholder(self.entry_aeroport, "Ex: Ivato International")
            self.entry_aeroport.pack(fill="x")

        def _field_ext_ref(parent):
            self.entry_ext_ref = styled_entry(parent)
            self.entry_ext_ref.pack(fill="x")

        tk.Label(
            complements_panel, text="Infos vol",
            font=("Poppins", 10, "bold"), fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(anchor="w", pady=(0, 8))

        two_column_row(complements_panel,
            "Heure d'arrivée", _field_heure_arrivee,
            "Heure de départ", _field_heure_depart,
        )
        inline_row(complements_panel, "Compagnie aérienne", _field_compagnie, pady=(0, 6))
        inline_row(complements_panel, "Aéroport international", _field_aeroport, pady=(0, 6))
        inline_row(complements_panel, "Réf. externe (Ext. Ref)", _field_ext_ref, pady=(0, 6))

        # ===== CARD: SEJOUR =====
        stay_card = create_card(right_col, title="Séjour")

        def _sl(parent, text):
            return tk.Label(parent, text=text, font=LABEL_FONT,
                            fg=TEXT_COLOR, bg=PANEL_BG_COLOR)

        _INPUT_W = 130   # largeur commune Arrivée / Départ / Saison
        _COMBO_W = 220   # largeur Hébergements / Restaurations

        def _stay_combo(parent, values, fixed_width=None):
            """Combobox arrondie, même style que les autres inputs."""
            if ctk:
                kw = dict(
                    values=values, state="readonly",
                    font=ENTRY_FONT, corner_radius=8,
                    fg_color=INPUT_BG_COLOR, text_color=TEXT_COLOR,
                    button_color="#C9DDE3", button_hover_color=BUTTON_BLUE,
                    border_color="#C9DDE3", border_width=1,
                    dropdown_fg_color=INPUT_BG_COLOR,
                    dropdown_text_color=TEXT_COLOR,
                    dropdown_hover_color="#E8F4F8",
                    height=34,
                )
                if fixed_width:
                    kw["width"] = fixed_width
                cb = ctk.CTkComboBox(parent, **kw
                )
            else:
                cb = ttk.Combobox(parent, values=values, state="readonly", font=ENTRY_FONT)
            return cb

        _CHIP_H = 34

        def _date_chip(parent, entry_attr):
            """Champ date cliquable : zone arrondie affichant la date + icône calendrier."""
            date_var = tk.StringVar()
            hidden = tk.Entry(
                parent, textvariable=date_var, width=0,
                relief="flat", bd=0, highlightthickness=0,
                state="readonly", readonlybackground=PANEL_BG_COLOR,
            )
            setattr(self, entry_attr, hidden)

            chip = tk.Canvas(parent, width=_INPUT_W, height=_CHIP_H, bg=PANEL_BG_COLOR,
                             highlightthickness=0, cursor="hand2")
            chip.pack(side="right")

            def _draw(*_):
                chip.delete("all")
                w = chip.winfo_width()
                if w < 10:
                    return
                r = 8
                x1, y1, x2, y2 = 0, 1, w - 1, _CHIP_H - 2
                pts = [x1+r,y1, x2-r,y1, x2,y1, x2,y1+r,
                       x2,y2-r, x2,y2, x2-r,y2, x1+r,y2,
                       x1,y2, x1,y2-r, x1,y1+r, x1,y1]
                chip.create_polygon(pts, smooth=True,
                                    fill=INPUT_BG_COLOR, outline="#C9DDE3", width=1)
                val = date_var.get()
                chip.create_text(
                    12, _CHIP_H // 2,
                    text=val or "Choisir une date", anchor="w",
                    font=ENTRY_FONT,
                    fill=TEXT_COLOR if val else MUTED_TEXT_COLOR,
                )
                # Icône calendrier simple
                chip.create_text(w - 16, _CHIP_H // 2,
                                 text="📅", anchor="center",
                                 font=("Poppins", 11))

            date_var.trace_add("write", _draw)
            chip.bind("<Configure>", _draw)
            chip.bind("<Button-1>",
                      lambda e: self._open_calendar(getattr(self, entry_attr)))
            chip.after(100, _draw)
            return chip

        # ── Grille 3 colonnes alignées (tk.grid) ─────────────────────────
        g = tk.Frame(stay_card, bg=PANEL_BG_COLOR)
        g.pack(fill="x", pady=(0, 8))
        g.columnconfigure(0, weight=0)
        g.columnconfigure(1, weight=1)
        g.columnconfigure(2, weight=1)

        # Ligne 0 : Durée | Arrivée | Hébergements
        r0c0 = tk.Frame(g, bg=PANEL_BG_COLOR)
        r0c0.grid(row=0, column=0, sticky="w", padx=(0, 18), pady=(0, 5))
        _sl(r0c0, "Durée").pack(side="left", padx=(0, 6))
        self.entry_duree_sejour = styled_entry(r0c0, readonly=True, width=4)
        if ctk and hasattr(self.entry_duree_sejour, "configure"):
            self.entry_duree_sejour.configure(width=60)
        self.entry_duree_sejour.pack(side="left")

        r0c1 = tk.Frame(g, bg=PANEL_BG_COLOR)
        r0c1.grid(row=0, column=1, sticky="ew", padx=(0, 18), pady=(0, 5))
        tk.Label(r0c1, text="Arrivée", font=LABEL_FONT, fg=TEXT_COLOR,
                 bg=PANEL_BG_COLOR, width=8, anchor="w").pack(side="left", padx=(0, 6))
        _date_chip(r0c1, "entry_date_arrivee")

        r0c2 = tk.Frame(g, bg=PANEL_BG_COLOR)
        r0c2.grid(row=0, column=2, sticky="ew", pady=(0, 5))
        tk.Label(r0c2, text="Hébergements", font=LABEL_FONT, fg=TEXT_COLOR,
                 bg=PANEL_BG_COLOR, width=14, anchor="w").pack(side="left", padx=(0, 6))
        self.combo_TypeHebergement = _stay_combo(r0c2, TYPE_HEBERGEMENTS, fixed_width=_COMBO_W)
        self.combo_TypeHebergement.pack(side="left")

        # Ligne 1 : vide | Départ | Restaurations
        r1c1 = tk.Frame(g, bg=PANEL_BG_COLOR)
        r1c1.grid(row=1, column=1, sticky="ew", padx=(0, 18), pady=(0, 5))
        tk.Label(r1c1, text="Départ", font=LABEL_FONT, fg=TEXT_COLOR,
                 bg=PANEL_BG_COLOR, width=8, anchor="w").pack(side="left", padx=(0, 6))
        _date_chip(r1c1, "entry_date_depart")

        r1c2 = tk.Frame(g, bg=PANEL_BG_COLOR)
        r1c2.grid(row=1, column=2, sticky="ew", pady=(0, 5))
        tk.Label(r1c2, text="Restaurations", font=LABEL_FONT, fg=TEXT_COLOR,
                 bg=PANEL_BG_COLOR, width=14, anchor="w").pack(side="left", padx=(0, 6))
        self.combo_restauration = _stay_combo(r1c2, RESTAURATIONS, fixed_width=_COMBO_W)
        self.combo_restauration.pack(side="left")

        # Ligne 2 : vide | Saison | vide
        r2c1 = tk.Frame(g, bg=PANEL_BG_COLOR)
        r2c1.grid(row=2, column=1, sticky="ew", padx=(0, 18), pady=(0, 5))
        tk.Label(r2c1, text="Saison", font=LABEL_FONT, fg=TEXT_COLOR,
                 bg=PANEL_BG_COLOR, width=8, anchor="w").pack(side="left", padx=(12, 6))
        self.combo_periode = _stay_combo(r2c1, PERIODES, fixed_width=_INPUT_W)
        self.combo_periode.pack(side="right")

        # combo_forfait conservé invisible pour la compatibilité des données
        self.combo_forfait = ttk.Combobox(stay_card, values=FORFAITS, state="readonly")

        # ── Accompagnement + Location de voiture (côte à côte) ───────────
        bottom = tk.Frame(stay_card, bg=PANEL_BG_COLOR)
        bottom.pack(fill="x", pady=(4, 0))

        accompagnement_frame = tk.Frame(bottom, bg=PANEL_BG_COLOR)
        accompagnement_frame.pack(side="left", fill="x", expand=True, padx=(0, 10))
        _sl(accompagnement_frame, "Accompagnement").pack(anchor="w")
        self.var_accompagnement_guide = tk.BooleanVar()
        self.var_accompagnement_chauffeur = tk.BooleanVar()

        def _on_guide_checked():
            if self.var_accompagnement_guide.get():
                self.var_accompagnement_chauffeur.set(False)

        def _on_chauffeur_checked():
            if self.var_accompagnement_chauffeur.get():
                self.var_accompagnement_guide.set(False)

        tk.Checkbutton(
            accompagnement_frame, text="Avec guide accompagnateur",
            variable=self.var_accompagnement_guide,
            command=_on_guide_checked,
            fg=TEXT_COLOR, bg=PANEL_BG_COLOR, selectcolor=BUTTON_GREEN, font=LABEL_FONT,
        ).pack(anchor="w")
        tk.Checkbutton(
            accompagnement_frame, text="Avec Chauffeur-guide",
            variable=self.var_accompagnement_chauffeur,
            command=_on_chauffeur_checked,
            fg=TEXT_COLOR, bg=PANEL_BG_COLOR, selectcolor=BUTTON_GREEN, font=LABEL_FONT,
        ).pack(anchor="w")

        location_frame = tk.Frame(bottom, bg=PANEL_BG_COLOR)
        location_frame.pack(side="left", fill="x", expand=True)
        _sl(location_frame, "Location de voiture").pack(anchor="w")
        self.var_location_voiture = tk.StringVar(value="")
        tk.Radiobutton(
            location_frame, text="Avec carburant",
            variable=self.var_location_voiture, value="Avec carburant",
            fg=TEXT_COLOR, bg=PANEL_BG_COLOR, selectcolor=BUTTON_GREEN, font=LABEL_FONT,
        ).pack(anchor="w")
        tk.Radiobutton(
            location_frame, text="Sans carburant",
            variable=self.var_location_voiture, value="Sans carburant",
            fg=TEXT_COLOR, bg=PANEL_BG_COLOR, selectcolor=BUTTON_GREEN, font=LABEL_FONT,
        ).pack(anchor="w")

        # ── Égaliser la hauteur des cartes Clients et Séjour ─────────────
        def _sync_card_heights():
            try:
                info_frame = info_card.master
                stay_frame = stay_card.master
                info_frame.update_idletasks()
                stay_frame.update_idletasks()
                h = max(info_frame.winfo_reqheight(), stay_frame.winfo_reqheight())
                for frm in (info_frame, stay_frame):
                    frm.configure(height=h)
                    frm.pack_propagate(False)
            except Exception:
                pass
        stay_card.after(300, _sync_card_heights)

        # ===== CARD: ROOMING LIST =====
        rooming_card = create_card(
            left_col,
            title="Rooming list",
            show_controls=True,
            on_add=self._add_rooming_row,
            on_remove=self._remove_rooming_row,
            expand=True,
        )

        # -- Participants section --
        participants_frame = tk.Frame(rooming_card, bg=PANEL_BG_COLOR)
        participants_frame.pack(fill="x", pady=(0, 6))

        self.var_enfant = tk.BooleanVar()
        self.check_enfant_widget = tk.Checkbutton(
            participants_frame,
            text="Voyage avec enfant",
            variable=self.var_enfant,
            command=self._toggle_enfant,
            fg=TEXT_COLOR,
            bg=PANEL_BG_COLOR,
            selectcolor=BUTTON_GREEN,
            font=("Poppins", 9),
        )
        self.check_enfant_widget.pack(anchor="w", pady=(0, 4))

        spinboxes_frame = tk.Frame(participants_frame, bg=PANEL_BG_COLOR)
        spinboxes_frame.pack(fill="x")

        def _spinbox_col(parent, label, attr_name):
            col = tk.Frame(parent, bg=PANEL_BG_COLOR)
            col.pack(side="left", fill="x", expand=True, padx=(0, 8))
            tk.Label(
                col,
                text=label,
                font=("Poppins", 9),
                fg=MUTED_TEXT_COLOR,
                bg=PANEL_BG_COLOR,
            ).pack(anchor="w")
            sb = tk.Spinbox(
                col,
                from_=0, to=999, width=5,
                font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                relief="flat", bd=0,
                highlightthickness=1, highlightbackground="#C9DDE3",
                highlightcolor=BUTTON_BLUE,
                buttonbackground=INPUT_BG_COLOR,
                insertbackground=TEXT_COLOR,
            )
            sb.pack(fill="x")
            sb.bind("<KeyRelease>", lambda e: self._update_participants_total())
            sb.bind("<<Increment>>", lambda e: self._update_participants_total())
            sb.bind("<<Decrement>>", lambda e: self._update_participants_total())
            setattr(self, attr_name, sb)
            return sb

        _spinbox_col(spinboxes_frame, "Adultes (+12 ans)", "entry_adultes")
        _spinbox_col(spinboxes_frame, "Enfants (2 à 12 ans)", "entry_enfants_2_12")
        _spinbox_col(spinboxes_frame, "Bébés (< 2 ans)", "entry_bebes_0_2")

        # Hidden compat entry (not packed)
        self.entry_total_participants = tk.Entry(participants_frame)

        # -- Column headers --
        _ROOM_HDR = [
            ("Date",        95),
            ("Nom & Prénom", 0),   # fill
            ("Nb pax",      50),
            ("Chambre",    110),
            ("Nombre",      55),
        ]
        hdr_bg = "#D0E8ED"
        room_hdr = tk.Frame(rooming_card, bg=hdr_bg)
        room_hdr.pack(fill="x", pady=(6, 1))
        for col_name, col_w in _ROOM_HDR:
            hdr_cell = tk.Frame(room_hdr, bg=hdr_bg)
            if col_w:
                hdr_cell.configure(width=col_w, height=28)
                hdr_cell.pack_propagate(False)
                hdr_cell.pack(side="left")
            else:
                hdr_cell.pack(side="left", fill="x", expand=True)
            tk.Label(
                hdr_cell, text=col_name,
                font=("Poppins", 9, "bold"), fg=TEXT_COLOR, bg=hdr_bg,
                padx=4, pady=3,
            ).pack(anchor="w")

        # -- Scrollable canvas area --
        room_scroll_outer = tk.Frame(rooming_card, bg=PANEL_BG_COLOR)
        room_scroll_outer.pack(fill="both", expand=True, pady=(0, 4))

        self._room_canvas = tk.Canvas(
            room_scroll_outer, bg=PANEL_BG_COLOR, highlightthickness=0, height=120
        )
        room_vsb = ttk.Scrollbar(room_scroll_outer, orient="vertical", command=self._room_canvas.yview)
        self._room_canvas.configure(yscrollcommand=room_vsb.set)
        room_vsb.pack(side="right", fill="y")
        self._room_canvas.pack(side="left", fill="both", expand=True)

        self._room_inner = tk.Frame(self._room_canvas, bg=PANEL_BG_COLOR)
        self._room_win = self._room_canvas.create_window(
            (0, 0), window=self._room_inner, anchor="nw"
        )

        def _room_inner_cfg(e):
            self._room_canvas.configure(scrollregion=self._room_canvas.bbox("all"))

        def _room_canvas_cfg(e):
            self._room_canvas.itemconfig(self._room_win, width=e.width)

        self._room_inner.bind("<Configure>", _room_inner_cfg)
        self._room_canvas.bind("<Configure>", _room_canvas_cfg)

        self._rooming_widget_rows = []

        # Legacy compat attributes (not displayed)
        self.rooming_vars = {}
        self.rooming_entries = {}
        self.rooming_date_entries = []
        self.rooming_name_entries = []
        self.rooming_pax_entries = []
        self.rooming_type_widgets = {}

        # Hidden compat entries (not packed)
        self.entry_rooming_date = tk.Entry(rooming_card)
        self.entry_rooming_nom = tk.Entry(rooming_card)
        self.entry_nombre_participants = tk.Entry(rooming_card)

        # -- Summary label --
        self.rooming_summary_label = muted_label(rooming_card, "Aucune ligne")
        self.rooming_summary_label.pack(anchor="w", pady=(2, 0))

        # Room type (hidden from UI; still auto-managed from rooming list)
        self.combo_TypeChambre = ttk.Combobox(
            stay_card, values=TYPE_CHAMBRES, state="readonly", width=37
        )

        # ===== CARD: ITINERAIRES =====
        # Widgets cachés pour compatibilité données (non affichés)
        self.combo_ville_depart = ttk.Combobox(right_col, values=self.city_options, state="normal", width=20)
        self.combo_ville_arrivee = ttk.Combobox(right_col, values=self.city_options, state="normal", width=20)
        self.entry_itineraire_date = styled_entry(right_col, readonly=True, width=12)
        self.entry_itineraire_distance = styled_entry(right_col, width=10)
        self.entry_itineraire_hebergement = styled_entry(right_col, width=18)

        itineraire_card = create_card(
            right_col,
            title=None,
            tabs=[("Itinéraires", True), ("Voiture", False), ("Guide national", False)],
            show_controls=True,
            on_add=self._add_itinerary_row,
            on_remove=self._remove_selected_itinerary_rows,
            expand=True,
        )

        # ── Sélection du circuit ──────────────────────────────────────────
        circuit_row = tk.Frame(itineraire_card, bg=PANEL_BG_COLOR)
        circuit_row.pack(fill="x", pady=(0, 6))

        tk.Label(
            circuit_row, text="Circuit :",
            font=("Poppins", 9, "bold"), fg=TEXT_COLOR, bg=PANEL_BG_COLOR,
        ).pack(side="left", padx=(0, 6))

        self.combo_circuit = ttk.Combobox(
            circuit_row, values=self.circuit_options, state="readonly", width=28,
        )
        self.combo_circuit.pack(side="left", padx=(0, 8))
        self.combo_circuit.bind(
            "<<ComboboxSelected>>", lambda e: self._on_circuit_selected(apply_route=True)
        )

        apply_btn = tk.Button(
            circuit_row, text="Appliquer",
            font=("Poppins", 8, "bold"), fg="white", bg=BUTTON_BLUE,
            relief="flat", bd=0, padx=8, pady=2, cursor="hand2",
            command=lambda: self._on_circuit_selected(apply_route=True),
        )
        apply_btn.pack(side="left", padx=(0, 12))

        self.circuit_info_label = muted_label(circuit_row, "")
        self.circuit_info_label.config(width=1, anchor="w")
        self.circuit_info_label.pack(side="left", fill="x", expand=True)

        # ── En-têtes de colonnes ─────────────────────────────────────────
        _ITIN = [
            ("Date",              105),
            ("Ville de départ",   145),
            ("Ville d'arrivée",   145),
            ("Distance (en Km)",   90),
            ("Hébergement",         0),  # fill
        ]
        hdr_bg = "#D0E8ED"
        hdr = tk.Frame(itineraire_card, bg=hdr_bg)
        hdr.pack(fill="x", pady=(0, 1))
        for col_name, col_w in _ITIN:
            cell = tk.Frame(hdr, bg=hdr_bg)
            if col_w:
                cell.configure(width=col_w, height=28)
                cell.pack_propagate(False)
                cell.pack(side="left")
            else:
                cell.pack(side="left", fill="x", expand=True)
            tk.Label(
                cell, text=col_name,
                font=("Poppins", 9, "bold"), fg=TEXT_COLOR, bg=hdr_bg,
                padx=4, pady=3,
            ).pack(anchor="w")

        # ── Zone scrollable ──────────────────────────────────────────────
        scroll_outer = tk.Frame(itineraire_card, bg=PANEL_BG_COLOR)
        scroll_outer.pack(fill="both", expand=True, pady=(0, 6))

        self._itin_canvas = tk.Canvas(
            scroll_outer, bg=PANEL_BG_COLOR, highlightthickness=0, height=220
        )
        itin_vsb = ttk.Scrollbar(scroll_outer, orient="vertical", command=self._itin_canvas.yview)
        self._itin_canvas.configure(yscrollcommand=itin_vsb.set)
        itin_vsb.pack(side="right", fill="y")
        self._itin_canvas.pack(side="left", fill="both", expand=True)

        self._itin_inner = tk.Frame(self._itin_canvas, bg=PANEL_BG_COLOR)
        self._itin_win = self._itin_canvas.create_window(
            (0, 0), window=self._itin_inner, anchor="nw"
        )

        def _itin_inner_cfg(e):
            self._itin_canvas.configure(scrollregion=self._itin_canvas.bbox("all"))

        def _itin_canvas_cfg(e):
            self._itin_canvas.itemconfig(self._itin_win, width=e.width)

        self._itin_inner.bind("<Configure>", _itin_inner_cfg)
        self._itin_canvas.bind("<Configure>", _itin_canvas_cfg)

        def _on_mousewheel(e):
            self._itin_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        self._itin_canvas.bind("<MouseWheel>", _on_mousewheel)
        self._itin_inner.bind("<MouseWheel>", _on_mousewheel)

        self.itinerary_count_label = muted_label(itineraire_card, "0 itinéraire sélectionné")
        self.itinerary_count_label.pack(anchor="w", pady=(0, 6))

        # Type d'hôtel à la ville d'arrivée (hidden from UI)
        self.combo_type_hotel_arrivee = ttk.Combobox(
            itineraire_card, values=HOTEL_ARRIVAL_TYPES, state="readonly", width=37
        )
        # Not packed on purpose: field removed from visible form.

        def _build_comment_block(parent, title):
            wrapper = tk.Frame(parent, bg=CARD_BG_COLOR)
            wrapper.pack(fill="x", pady=(6, 10))
            tk.Label(
                wrapper,
                text=title,
                font=("Poppins", 10, "bold"),
                fg=TEXT_COLOR,
                bg=CARD_BG_COLOR,
            ).pack(anchor="w", pady=(0, 4))
            hint = tk.Frame(wrapper, bg=CARD_BG_COLOR)
            hint.pack(fill="x", pady=(0, 4))
            icon = tk.Canvas(
                hint,
                width=16,
                height=16,
                bg=CARD_BG_COLOR,
                highlightthickness=0,
            )
            icon.create_rectangle(1, 1, 15, 15, fill=BUTTON_GREEN, outline=BUTTON_GREEN)
            icon.create_text(
                8,
                8,
                text="+",
                fill="white",
                font=("Poppins", 9, "bold"),
            )
            icon.pack(side="left")
            tk.Label(
                hint,
                text="Pour ajouter un nouveau commentaire, cliquez le bouton commentaire",
                font=("Poppins", 8),
                fg=MUTED_TEXT_COLOR,
                bg=CARD_BG_COLOR,
            ).pack(side="left", padx=(6, 0))
            text = tk.Text(
                wrapper,
                height=6,
                font=ENTRY_FONT,
                bg="#FFF6E0",
                fg=TEXT_COLOR,
                wrap="word",
                relief="flat",
                highlightthickness=1,
                highlightbackground="#E6C87A",
                highlightcolor="#E6C87A",
                bd=0,
                padx=10,
                pady=8,
            )
            text.pack(fill="x")
            return text

        self.comment_client_text = _build_comment_block(
            left_col, "Commentaires client"
        )
        self.comment_internal_text = _build_comment_block(
            right_col, "Commentaires internes"
        )

        # Buttons
        btn_frame = tk.Frame(self.main_frame, bg=CARD_BG_COLOR)
        btn_frame.pack(fill="x", pady=(6, 14))
        actions = tk.Frame(btn_frame, bg=CARD_BG_COLOR)
        actions.pack(side="right", padx=16)

        def refresh_form():
            self._reset_form()
            if self.client_to_edit:
                self._populate_fields()

        if self.client_to_edit:
            action_button(
                actions, "Actualiser", variant="info", command=refresh_form
            ).pack(side="left", padx=4)
            action_button(actions, "Modifier", command=self._validate).pack(
                side="left", padx=4
            )
            action_button(
                actions, "Abandonner", variant="danger", command=self._cancel
            ).pack(side="left", padx=4)
        else:
            action_button(
                actions, "Actualiser", variant="info", command=refresh_form
            ).pack(side="left", padx=4)
            action_button(actions, "Valider", command=self._validate).pack(
                side="left", padx=4
            )
            action_button(
                actions, "Abandonner", variant="danger", command=self._cancel
            ).pack(side="left", padx=4)

        # Populate fields if editing
        if self.client_to_edit:
            self._populate_fields()
        else:
            self._toggle_enfant()
            self._update_rooming_summary()
            self._update_itinerary_count()

        self.combo_type_client.focus_set()
        self.parent.bind("<Control-s>", lambda e: self._validate())
        self.parent.bind("<Control-S>", lambda e: self._validate())
        if self.client_to_edit:
            self.parent.bind("<Escape>", lambda e: self._cancel())

    def _populate_fields(self):
        """Populate form fields with client data"""
        if not self.client_to_edit:
            return

        # Split telephone into code and number
        telephone = self.client_to_edit.get("telephone", "")
        code_pays = ""
        numero = ""
        if telephone.startswith("+"):
            for code in PHONE_CODES:
                if telephone.startswith(code):
                    code_pays = code
                    numero = telephone[len(code) :]
                    break
        else:
            numero = telephone

        # Populate basic fields (readonly → normal → insert → readonly)
        for entry, key in (
            (self.entry_ref_client, "ref_client"),
            (self.entry_numero_dossier, "numero_dossier"),
        ):
            try:
                entry.configure(state="normal")
            except Exception:
                pass
            entry.delete(0, tk.END)
            entry.insert(0, self.client_to_edit.get(key, ""))
            try:
                entry.configure(state="readonly")
            except Exception:
                pass
        self._clear_placeholder(self.entry_prenom)
        self._clear_placeholder(self.entry_nom)
        self.entry_prenom.insert(0, self.client_to_edit.get("prenom", ""))
        self.entry_nom.insert(0, self.client_to_edit.get("nom", ""))
        self.entry_date_arrivee.configure(state="normal")
        self.entry_date_arrivee.insert(0, self.client_to_edit.get("date_arrivee", ""))
        self.entry_date_arrivee.configure(state="readonly")
        self.entry_date_depart.configure(state="normal")
        self.entry_date_depart.insert(0, self.client_to_edit.get("date_depart", ""))
        self.entry_date_depart.configure(state="readonly")
        self._sync_rooming_date()
        self._update_rooming_identity()

        if hasattr(self, "entry_total_participants"):
            self.entry_total_participants.configure(state="normal")
            self.entry_total_participants.insert(
                0, self.client_to_edit.get("nombre_participants", "")
            )
            self.entry_total_participants.configure(state="readonly")
        self.entry_adultes.insert(0, self.client_to_edit.get("nombre_adultes", ""))
        self.entry_enfants_2_12.insert(
            0, self.client_to_edit.get("nombre_enfants_2_12", "")
        )
        self.entry_bebes_0_2.insert(0, self.client_to_edit.get("nombre_bebes_0_2", ""))

        self.entry_email.insert(0, self.client_to_edit.get("email", ""))
        self.entry_code_pays.set(code_pays or DEFAULT_PHONE_CODE)
        self.entry_telephone.insert(0, numero)

        # WhatsApp
        telephone_whatsapp = self.client_to_edit.get("telephone_whatsapp", "")
        code_whatsapp = ""
        numero_whatsapp = ""
        if telephone_whatsapp.startswith("+"):
            for code in PHONE_CODES:
                if telephone_whatsapp.startswith(code):
                    code_whatsapp = code
                    numero_whatsapp = telephone_whatsapp[len(code) :]
                    break
        else:
            numero_whatsapp = telephone_whatsapp

        if hasattr(self, "whatsapp_code_var"):
            self.whatsapp_code_var.set(code_whatsapp or code_pays or DEFAULT_PHONE_CODE)
        if hasattr(self, "whatsapp_number_var"):
            self.whatsapp_number_var.set(numero_whatsapp or numero)

        # Statut dossier
        saved_statut = self.client_to_edit.get("statut", "En cours") or "En cours"
        self._apply_statut(saved_statut)

        # Compléments (infos vol)
        self.entry_heure_arrivee.delete(0, tk.END)
        self.entry_heure_arrivee.insert(0, self.client_to_edit.get("heure_arrivee", ""))
        self.entry_heure_depart.delete(0, tk.END)
        self.entry_heure_depart.insert(0, self.client_to_edit.get("heure_depart", ""))
        self.combo_compagnie.set(self.client_to_edit.get("compagnie", ""))
        self.entry_aeroport.delete(0, tk.END)
        self.entry_aeroport.insert(0, self.client_to_edit.get("aeroport", ""))
        self.entry_ext_ref.delete(0, tk.END)
        self.entry_ext_ref.insert(0, self.client_to_edit.get("ext_ref", ""))

        # Accompagnement / Location voiture
        guide_value = str(self.client_to_edit.get("accompagnement_guide", "")).strip().lower()
        chauffeur_value = str(
            self.client_to_edit.get("accompagnement_chauffeur", "")
        ).strip().lower()
        self.var_accompagnement_guide.set(guide_value in {"oui", "true", "1", "yes"})
        self.var_accompagnement_chauffeur.set(
            chauffeur_value in {"oui", "true", "1", "yes"}
        )
        self.var_location_voiture.set(self.client_to_edit.get("location_voiture", ""))

        # Clear and rebuild rooming rows from saved data
        for rw in getattr(self, "_rooming_widget_rows", []):
            rw["frame"].destroy()
        if hasattr(self, "_rooming_widget_rows"):
            self._rooming_widget_rows.clear()
        room_labels = {
            "sgl": "SGL (Single)", "dbl": "DBL (Double)",
            "twn": "TWN (Twin)", "tpl": "TPL (Triple)", "fml": "FML (Familiale)",
        }
        for key, label in room_labels.items():
            count_str = str(self.client_to_edit.get(f"{key}_count", "") or "").strip()
            if count_str and count_str != "0":
                self._add_rooming_widget_row({
                    "date": self.client_to_edit.get("date_arrivee", ""),
                    "nom": f"{self.client_to_edit.get('nom', '')} {self.client_to_edit.get('prenom', '')}".strip(),
                    "nb_pax": self.client_to_edit.get("adultes", ""),
                    "chambre": label,
                    "nombre": count_str,
                })
        self._update_rooming_summary()
        # Combo boxes
        type_client = self.client_to_edit.get("type_client", "").strip()
        self.combo_type_client.set(type_client if type_client in ("Mr", "Mme", "CIE") else "Mr")
        self.combo_periode.set(self.client_to_edit.get("periode", ""))
        self.combo_restauration.set(self.client_to_edit.get("restauration", ""))
        self.combo_TypeHebergement.set(self.client_to_edit.get("hebergement", ""))
        self.combo_TypeChambre.set(self.client_to_edit.get("chambre", ""))
        self.combo_forfait.set(self.client_to_edit.get("forfait", ""))
        self.combo_circuit.set(self.client_to_edit.get("circuit", ""))
        self._on_circuit_selected(apply_route=False)
        self.combo_ville_depart.set(self.client_to_edit.get("ville_depart", ""))
        self.combo_type_hotel_arrivee.set(
            self.client_to_edit.get("type_hotel_arrivee", "")
        )
        self._set_itinerary_rows(
            self.client_to_edit.get("itineraire_detail")
            or self.client_to_edit.get("ville_arrivee", ""),
            depart_value=self.client_to_edit.get("ville_depart", ""),
        )
        self._update_type_chambre_from_rooming(clear_if_empty=False)

        # Handle children
        enfant = self.client_to_edit.get("enfant", "")
        if enfant.lower() == "oui":
            self.var_enfant.set(True)
            self._toggle_enfant()
        else:
            self.var_enfant.set(False)
            self._toggle_enfant()

        self._update_participants_total()

    def _open_calendar(self, entry_widget):
        """Open calendar dialog for date selection"""
        cal_dialog = CalendarDialog(self.parent, "Choisir une date")
        self.parent.wait_window(cal_dialog)

        if cal_dialog.selected_date:
            entry_widget.configure(state="normal")
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, cal_dialog.selected_date.strftime("%d/%m/%Y"))
            entry_widget.configure(state="readonly")
            # Auto-calculate duration if both dates are set
            if entry_widget in (self.entry_date_arrivee, self.entry_date_depart):
                self._calculate_duration()
                self._sync_rooming_date()

    def _calculate_duration(self):
        """Calculate stay duration based on arrival and departure dates"""
        try:
            date_arr = self.entry_date_arrivee.get()
            date_dep = self.entry_date_depart.get()

            if date_arr and date_dep:
                arrival = datetime.strptime(date_arr, "%d/%m/%Y")
                departure = datetime.strptime(date_dep, "%d/%m/%Y")
                duration = (departure - arrival).days

                self.entry_duree_sejour.configure(state="normal")
                self.entry_duree_sejour.delete(0, tk.END)
                self.entry_duree_sejour.insert(0, f"{duration} jours")
                self.entry_duree_sejour.configure(state="readonly")
        except:
            pass

    def _sync_rooming_date(self):
        """Mirror arrival date to rooming list date field."""
        if not hasattr(self, "_rooming_widget_rows"):
            return
        date_arr = self.entry_date_arrivee.get()
        for rw in self._rooming_widget_rows:
            rw["date_var"].set(date_arr)

    def _update_rooming_identity(self):
        """Sync rooming list name field with client name."""
        if not hasattr(self, "_rooming_widget_rows"):
            return
        prenom = self._get_entry_value(self.entry_prenom).strip()
        nom = self._get_entry_value(self.entry_nom).strip()
        full_name = f"{nom} {prenom}".strip()
        for rw in self._rooming_widget_rows:
            rw["nom_var"].set(full_name)

    def _toggle_enfant(self):
        """Show/hide child age field"""
        if self.var_enfant.get():
            self.entry_enfants_2_12.configure(state="normal")
            self.entry_bebes_0_2.configure(state="normal")
        else:
            self.entry_enfants_2_12.configure(state="normal")
            self.entry_enfants_2_12.delete(0, tk.END)
            self.entry_enfants_2_12.insert(0, "0")
            self.entry_enfants_2_12.configure(state="disabled")
            self.entry_bebes_0_2.configure(state="normal")
            self.entry_bebes_0_2.delete(0, tk.END)
            self.entry_bebes_0_2.insert(0, "0")
            self.entry_bebes_0_2.configure(state="disabled")

        self._update_participants_total()

    def _update_participants_total(self):
        """Auto-update total participants from adults + children + babies."""

        def _to_int(value):
            try:
                return int(str(value).strip())
            except Exception:
                return 0

        adults = _to_int(self.entry_adultes.get())
        enfants = _to_int(self.entry_enfants_2_12.get())
        bebes = _to_int(self.entry_bebes_0_2.get())
        total = adults + enfants + bebes
        if hasattr(self, "entry_total_participants"):
            try:
                self.entry_total_participants.configure(state="normal")
                self.entry_total_participants.delete(0, tk.END)
                self.entry_total_participants.insert(0, str(total))
            except Exception:
                pass

    def _on_room_toggle(self, room_key):
        """Enable/disable room entry field when checkbox is toggled"""
        pass

    def _on_room_count_change(self):
        """Refresh rooming derived UI when quantities change."""
        pass

    def _add_rooming_widget_row(self, data=None):
        """Create one editable row in the rooming list canvas area."""
        data = data or {}
        COL_W = [95, 0, 50, 110, 55]
        BORDER = "#C9DDE3"
        row = tk.Frame(self._room_inner, bg=BORDER)
        row.pack(fill="x", pady=(0, 1))

        def _cell(w):
            f = tk.Frame(row, bg=BORDER)
            if w:
                f.configure(width=w, height=32)
                f.pack_propagate(False)
                f.pack(side="left", padx=(0, 1))
            else:
                f.pack(side="left", fill="x", expand=True)
            return f

        def _entry(parent, textvariable, readonly=False):
            e = tk.Entry(parent, font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                         textvariable=textvariable, relief="flat", bd=0,
                         insertbackground=TEXT_COLOR)
            if readonly:
                e.configure(state="readonly", readonlybackground=INPUT_BG_COLOR)
            return e

        date_var = tk.StringVar(value=data.get("date", ""))
        _entry(_cell(COL_W[0]), date_var, readonly=True).pack(fill="both", expand=True, padx=1, pady=1, ipady=3)

        nom_var = tk.StringVar(value=data.get("nom", ""))
        _entry(_cell(COL_W[1]), nom_var).pack(fill="both", expand=True, padx=1, pady=1, ipady=3)

        pax_var = tk.StringVar(value=data.get("nb_pax", ""))
        _entry(_cell(COL_W[2]), pax_var).pack(fill="both", expand=True, padx=1, pady=1, ipady=3)

        ROOM_OPTS = ["SGL (Single)", "DBL (Double)", "TWN (Twin)", "TPL (Triple)", "FML (Familiale)"]
        chambre_var = tk.StringVar(value=data.get("chambre", ""))
        ttk.Combobox(_cell(COL_W[3]), values=ROOM_OPTS, textvariable=chambre_var,
                     state="readonly", font=ENTRY_FONT).pack(fill="both", expand=True, padx=1, pady=1, ipady=3)

        nombre_var = tk.StringVar(value=data.get("nombre", ""))
        _entry(_cell(COL_W[4]), nombre_var).pack(fill="both", expand=True, padx=1, pady=1, ipady=3)

        self._rooming_widget_rows.append({
            "frame": row, "date_var": date_var, "nom_var": nom_var,
            "pax_var": pax_var, "chambre_var": chambre_var, "nombre_var": nombre_var,
        })
        self._room_inner.update_idletasks()
        self._room_canvas.configure(scrollregion=self._room_canvas.bbox("all"))
        self._update_rooming_summary()

    def _add_rooming_row(self):
        date_arr = self.entry_date_arrivee.get() if hasattr(self, "entry_date_arrivee") else ""
        prenom = self._get_entry_value(self.entry_prenom).strip() if hasattr(self, "entry_prenom") else ""
        nom = self._get_entry_value(self.entry_nom).strip() if hasattr(self, "entry_nom") else ""
        self._add_rooming_widget_row({"date": date_arr, "nom": f"{nom} {prenom}".strip()})

    def _remove_rooming_row(self):
        if self._rooming_widget_rows:
            self._rooming_widget_rows.pop()["frame"].destroy()
            self._room_canvas.configure(scrollregion=self._room_canvas.bbox("all"))
            self._update_rooming_summary()

    def _update_rooming_summary(self):
        """Display a compact summary for selected rooming."""
        n = len(self._rooming_widget_rows) if hasattr(self, "_rooming_widget_rows") else 0
        self.rooming_summary_label.config(
            text="Aucune ligne" if n == 0 else f"{n} ligne(s)"
        )

    def _update_type_chambre_from_rooming(self, clear_if_empty=True):
        """Auto-fill room type based on rooming list quantities."""
        room_key_map = {
            "SGL": "Single", "DBL": "Double/twin", "TWN": "Double/twin",
            "TPL": "Triple", "FML": "Familliale",
        }
        counts = {}
        for rw in getattr(self, "_rooming_widget_rows", []):
            chambre = rw["chambre_var"].get()
            for k, v in room_key_map.items():
                if k in chambre.upper():
                    try:
                        n = int(rw["nombre_var"].get().strip() or "0")
                    except Exception:
                        n = 0
                    counts[v] = counts.get(v, 0) + n
                    break
        if counts:
            priority = ["Single", "Double/twin", "Triple", "Familliale"]
            best = sorted(counts.items(), key=lambda x: (-x[1], priority.index(x[0]) if x[0] in priority else 99))[0][0]
            self.combo_TypeChambre.set(best)
        elif clear_if_empty:
            self.combo_TypeChambre.set("")

    def _add_itinerary_widget_row(self, data=None):
        """Create one editable row in the itinerary canvas area."""
        data = data or {}
        COL_W = [105, 145, 145, 90, 0]
        BORDER = "#C9DDE3"

        row = tk.Frame(self._itin_inner, bg=BORDER)
        row.pack(fill="x", pady=(0, 1))

        def _cell(w):
            f = tk.Frame(row, bg=BORDER)
            if w:
                f.configure(width=w, height=32)
                f.pack_propagate(False)
                f.pack(side="left", padx=(0, 1))
            else:
                f.pack(side="left", fill="x", expand=True)
            return f

        def _entry(parent, textvariable, readonly=False):
            e = tk.Entry(parent, font=ENTRY_FONT, bg=INPUT_BG_COLOR, fg=TEXT_COLOR,
                         textvariable=textvariable, relief="flat", bd=0,
                         insertbackground=TEXT_COLOR)
            if readonly:
                e.configure(state="readonly", readonlybackground=INPUT_BG_COLOR)
            return e

        # Date
        date_var = tk.StringVar(value=data.get("date", ""))
        date_cell = _cell(COL_W[0])
        date_e = _entry(date_cell, date_var, readonly=True)
        date_e.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=1, ipady=3)
        self._make_calendar_badge(date_cell, date_e, "15").pack(side="left", padx=(2, 1), pady=1)

        # Ville de départ
        depart_var = tk.StringVar(value=data.get("depart", ""))
        dep_cell = _cell(COL_W[1])
        dep_cb = ttk.Combobox(dep_cell, values=self.city_options, textvariable=depart_var,
                              state="normal", font=ENTRY_FONT)
        dep_cb.pack(fill="both", expand=True, padx=1, pady=1, ipady=3)

        # Ville d'arrivée
        arrivee_var = tk.StringVar(value=data.get("arrivee", ""))
        arr_cell = _cell(COL_W[2])
        arr_cb = ttk.Combobox(arr_cell, values=self.city_options, textvariable=arrivee_var,
                              state="normal", font=ENTRY_FONT)
        arr_cb.pack(fill="both", expand=True, padx=1, pady=1, ipady=3)

        # Distance
        dist_var = tk.StringVar(value=data.get("distance", ""))
        dist_cell = _cell(COL_W[3])
        dist_e = _entry(dist_cell, dist_var)
        dist_e.pack(fill="both", expand=True, padx=1, pady=1, ipady=3)

        # Hébergement (fill)
        heb_var = tk.StringVar(value=data.get("hebergement", ""))
        heb_cell = _cell(COL_W[4])
        heb_e = _entry(heb_cell, heb_var)
        heb_e.pack(fill="both", expand=True, padx=1, pady=1, ipady=3)

        # Bind mousewheel to child widgets
        for w in (date_e, dep_cb, arr_cb, dist_e, heb_e, row, date_cell, dep_cell, arr_cell, dist_cell, heb_cell):
            w.bind("<MouseWheel>", lambda e: self._itin_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        self._itin_widget_rows.append({
            "frame": row,
            "date_var": date_var,
            "depart_var": depart_var,
            "arrivee_var": arrivee_var,
            "distance_var": dist_var,
            "heb_var": heb_var,
        })

        self._itin_inner.update_idletasks()
        self._itin_canvas.configure(scrollregion=self._itin_canvas.bbox("all"))
        self._update_itinerary_count()

    def _add_itinerary_row(self):
        """Add an empty editable row to the itinerary table."""
        self._add_itinerary_widget_row()

    def _remove_selected_itinerary_rows(self):
        """Remove the last itinerary row."""
        if not self._itin_widget_rows:
            return
        last = self._itin_widget_rows.pop()
        last["frame"].destroy()
        self._itin_inner.update_idletasks()
        self._itin_canvas.configure(scrollregion=self._itin_canvas.bbox("all"))
        self._update_itinerary_count()

    def _get_itinerary_rows(self):
        """Return itinerary rows read from the editable widgets."""
        rows = []
        for r in self._itin_widget_rows:
            rows.append({
                "date": r["date_var"].get().strip(),
                "depart": r["depart_var"].get().strip(),
                "arrivee": r["arrivee_var"].get().strip(),
                "distance": r["distance_var"].get().strip(),
                "hebergement": r["heb_var"].get().strip(),
            })
        return rows

    def _serialize_itinerary_rows(self, rows):
        """Serialize itinerary rows into a compact string."""
        lines = []
        for row in rows:
            lines.append(
                " | ".join(
                    [
                        row.get("date", ""),
                        row.get("depart", ""),
                        row.get("arrivee", ""),
                        row.get("distance", ""),
                        row.get("hebergement", ""),
                    ]
                )
            )
        return "\n".join(lines)

    def _refresh_itinerary_table(self):
        """Rebuild editable widget rows from self.itinerary_rows."""
        if self._itin_canvas is None:
            return
        for r in self._itin_widget_rows:
            r["frame"].destroy()
        self._itin_widget_rows = []
        for row_data in self.itinerary_rows:
            self._add_itinerary_widget_row(row_data)
        self._update_itinerary_count()

    def _set_itinerary_rows(self, rows_value, depart_value=""):
        """Set itinerary rows from a string or list."""
        self.itinerary_rows = []
        if not rows_value:
            self._refresh_itinerary_table()
            return

        if isinstance(rows_value, list):
            rows = rows_value
        else:
            rows = []
            raw = str(rows_value)
            if "|" in raw:
                for line in raw.splitlines():
                    parts = [p.strip() for p in line.split("|")]
                    while len(parts) < 5:
                        parts.append("")
                    if any(parts):
                        rows.append(
                            {
                                "date": parts[0],
                                "depart": parts[1],
                                "arrivee": parts[2],
                                "distance": parts[3],
                                "hebergement": parts[4],
                            }
                        )
            else:
                cities = [
                    c.strip() for c in re.split(r"[;,>\n]+", raw) if c.strip()
                ]
                previous = depart_value.strip()
                for city in cities:
                    rows.append(
                        {
                            "date": "",
                            "depart": previous,
                            "arrivee": city,
                            "distance": "",
                            "hebergement": "",
                        }
                    )
                    previous = city

        self.itinerary_rows.extend(rows)
        self._refresh_itinerary_table()

    def _parse_circuit_itinerary(self, itinerary_value):
        """Parse itinerary text into ordered city names."""
        if not itinerary_value:
            return []
        raw_parts = re.split(r"[;,>\n]+", str(itinerary_value))
        return [part.strip() for part in raw_parts if part and part.strip()]

    def _on_circuit_selected(self, _event=None, apply_route=True):
        """Update UI from selected circuit metadata."""
        circuit_name = self.combo_circuit.get().strip()
        circuit = self.circuit_map.get(circuit_name, {})
        if not circuit:
            self.circuit_info_label.config(text="")
            return

        info_lines = []
        if circuit.get("duree"):
            info_lines.append(f"Durée: {circuit['duree']}")
        if circuit.get("condition_physique"):
            info_lines.append(f"Condition: {circuit['condition_physique']}")
        if circuit.get("type_voiture"):
            info_lines.append(f"Véhicule: {circuit['type_voiture']}")
        if circuit.get("activite"):
            info_lines.append(f"Activité: {circuit['activite']}")
        self.circuit_info_label.config(text="  ·  ".join(info_lines))

        if not apply_route:
            return

        cities = self._parse_circuit_itinerary(circuit.get("itineraire", ""))
        if not cities:
            return

        self.itinerary_rows = [
            {"date": "", "depart": cities[i - 1], "arrivee": cities[i], "distance": "", "hebergement": ""}
            for i in range(1, len(cities))
        ]
        self._refresh_itinerary_table()

    def _update_itinerary_count(self):
        """Update itinerary counter label."""
        count = len(self._itin_widget_rows)
        suffix = "itinéraire sélectionné" if count == 1 else "itinéraires sélectionnés"
        self.itinerary_count_label.config(text=f"{count} {suffix}")

    def _validate(self):
        """Validate and save client data"""
        selected_circuit = self.circuit_map.get(self.combo_circuit.get().strip(), {})
        existing_circuit = self.client_to_edit or {}
        itinerary_rows = self._get_itinerary_rows()
        ville_depart = (
            itinerary_rows[0].get("depart", "").strip()
            if itinerary_rows
            else self.combo_ville_depart.get().strip()
        )
        ville_arrivee = ", ".join(
            [row.get("arrivee", "").strip() for row in itinerary_rows if row.get("arrivee")]
        )
        itineraire_detail = self._serialize_itinerary_rows(itinerary_rows)
        type_client = self.combo_type_client.get().strip()
        mobile_phone = self.entry_code_pays.get() + self.entry_telephone.get()
        whatsapp_number = (
            self.whatsapp_number_var.get().strip()
            if hasattr(self, "whatsapp_number_var")
            else ""
        )
        whatsapp_code = (
            self.whatsapp_code_var.get().strip()
            if hasattr(self, "whatsapp_code_var")
            else ""
        )
        whatsapp_phone = (
            f"{whatsapp_code}{whatsapp_number}" if whatsapp_number else mobile_phone
        )

        # Collect rooming rows data
        _room_key_map = {"SGL": "sgl", "DBL": "dbl", "TWN": "twn", "TPL": "tpl", "FML": "fml"}
        _room_counts = {"sgl": 0, "dbl": 0, "twn": 0, "tpl": 0, "fml": 0}
        for rw in getattr(self, "_rooming_widget_rows", []):
            chambre = rw["chambre_var"].get()
            for k, v in _room_key_map.items():
                if k in chambre.upper():
                    try:
                        _room_counts[v] += int(rw["nombre_var"].get().strip() or "0")
                    except Exception:
                        pass
                    break

        # Auto-generate ref/dossier for new clients if not already set
        ref_client_val = self._get_entry_value(self.entry_ref_client).strip()
        numero_dossier_val = self._get_entry_value(self.entry_numero_dossier).strip()
        if not self.client_to_edit and (not ref_client_val or not numero_dossier_val):
            ref_auto, doss_auto = self._generate_client_refs()
            if not ref_client_val:
                ref_client_val = ref_auto
            if not numero_dossier_val:
                numero_dossier_val = doss_auto

        form_data = {
            "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "date_jour": self.entry_date_jour.get(),
            "ref_client": ref_client_val,
            "numero_dossier": numero_dossier_val,
            "type_client": type_client,
            "prenom": self._get_entry_value(self.entry_prenom).strip(),
            "nom": self._get_entry_value(self.entry_nom).strip(),
            "date_arrivee": self.entry_date_arrivee.get(),
            "date_depart": self.entry_date_depart.get(),
            "duree_sejour": self.entry_duree_sejour.get(),
            "nombre_participants": self.entry_total_participants.get().strip(),
            "nombre_adultes": self.entry_adultes.get().strip(),
            "nombre_enfants_2_12": self.entry_enfants_2_12.get().strip(),
            "nombre_bebes_0_2": self.entry_bebes_0_2.get().strip(),
            "telephone": mobile_phone,
            "telephone_whatsapp": whatsapp_phone,
            "email": self.entry_email.get().strip(),
            "periode": self.combo_periode.get(),
            "restauration": self.combo_restauration.get(),
            "hebergement": self.combo_TypeHebergement.get(),
            "chambre": self.combo_TypeChambre.get(),
            "accompagnement_guide": "Oui"
            if self.var_accompagnement_guide.get()
            else "Non",
            "accompagnement_chauffeur": "Oui"
            if self.var_accompagnement_chauffeur.get()
            else "Non",
            "location_voiture": self.var_location_voiture.get(),
            "enfant": "Oui" if self.var_enfant.get() else "Non",
            "age_enfant": "",
            "statut": self._statut_var.get(),
            "heure_arrivee": self._get_entry_value(self.entry_heure_arrivee).strip(),
            "heure_depart": self._get_entry_value(self.entry_heure_depart).strip(),
            "compagnie": self.combo_compagnie.get().strip(),
            "aeroport": self._get_entry_value(self.entry_aeroport).strip(),
            "ext_ref": self._get_entry_value(self.entry_ext_ref).strip(),
            "forfait": self.combo_forfait.get(),
            "circuit": self.combo_circuit.get(),
            "type_circuit": selected_circuit.get("type_circuit")
            or existing_circuit.get("type_circuit", ""),
            "id_circuit": selected_circuit.get("id_circuit")
            or existing_circuit.get("id_circuit", ""),
            "itineraire_circuit": selected_circuit.get("itineraire")
            or existing_circuit.get("itineraire_circuit", ""),
            "activite_circuit": selected_circuit.get("activite")
            or existing_circuit.get("activite_circuit", ""),
            "duree_circuit": selected_circuit.get("duree")
            or existing_circuit.get("duree_circuit", ""),
            "condition_physique_circuit": selected_circuit.get("condition_physique")
            or existing_circuit.get("condition_physique_circuit", ""),
            "type_voiture_circuit": selected_circuit.get("type_voiture")
            or existing_circuit.get("type_voiture_circuit", ""),
            "hotels_defaut_villes_circuit": selected_circuit.get("hotels_defaut_villes")
            or existing_circuit.get("hotels_defaut_villes_circuit", ""),
            "prestations_incluses_circuit": selected_circuit.get("prestations_incluses")
            or existing_circuit.get("prestations_incluses_circuit", ""),
            "transports_associes_circuit": selected_circuit.get("transports_associes")
            or existing_circuit.get("transports_associes_circuit", ""),
            "ville_depart": ville_depart,
            "ville_arrivee": ville_arrivee,
            "itineraire_detail": itineraire_detail,
            "type_hotel_arrivee": self.combo_type_hotel_arrivee.get(),
            "sgl_count": str(_room_counts["sgl"]) if _room_counts["sgl"] else "",
            "dbl_count": str(_room_counts["dbl"]) if _room_counts["dbl"] else "",
            "twn_count": str(_room_counts["twn"]) if _room_counts["twn"] else "",
            "tpl_count": str(_room_counts["tpl"]) if _room_counts["tpl"] else "",
            "fml_count": str(_room_counts["fml"]) if _room_counts["fml"] else "",
        }

        # Create client data object
        client = ClientData.from_form_data(form_data)

        # Validate
        errors = client.validate()

        # Additional validations
        if not validate_email(client.email):
            errors.append("Email invalide")
        if not validate_phone_number(
            self.entry_code_pays.get(), self.entry_telephone.get()
        ):
            errors.append("Numéro téléphone invalide")

        if errors:
            messagebox.showerror("❌ Erreur", "\n".join(errors))
            return

        # Save or update to Excel
        try:
            if self.client_to_edit:
                # Update existing client
                from utils.excel_handler import update_client_in_excel

                success = update_client_in_excel(
                    self.client_to_edit["row_number"], client.to_dict()
                )
                if success:
                    messagebox.showinfo(
                        "✅ SUCCÈS", f"Client {client.nom} modifié avec succès !"
                    )
                    logger.info(f"Client updated: {client.ref_client} - {client.nom}")
                    if self.on_save_callback:
                        self.on_save_callback()
                else:
                    error_msg = (
                        "Erreur lors de la modification du client. Voir les logs."
                    )
                    messagebox.showerror("❌ Erreur Excel", error_msg)
                    logger.error(f"Failed to update client: {client.ref_client}")
            else:
                # Save new client
                row = save_client_to_excel(client.to_dict())
                if row > 0:
                    messagebox.showinfo(
                        "✅ SUCCÈS", f"Client {client.nom} sauvé ligne Excel {row} !"
                    )
                    logger.info(
                        f"New client saved: {client.ref_client} - {client.nom} at row {row}"
                    )
                    self._reset_form()
                    if self.on_save_callback:
                        self.on_save_callback()
                else:
                    error_msg = "Erreur lors de la sauvegarde du client. Voir les logs."
                    messagebox.showerror("❌ Erreur Excel", error_msg)
                    logger.error(f"Failed to save new client: {client.ref_client}")
        except Exception as e:
            error_msg = (
                f"Erreur Excel: {str(e)}\n\nDétails dans les logs de l'application."
            )
            messagebox.showerror("❌ Erreur", error_msg)
            logger.error(f"Exception during client save/update: {e}", exc_info=True)

    _STATUT_COLORS = {
        "En cours":   "#00BCD4",
        "Accepté":    "#4CAF50",
        "En circuit": "#FF9800",
        "Annulé":     "#F44336",
    }

    def _generate_client_refs(self):
        """Auto-generate ref_client and numero_dossier based on the last used sequence."""
        try:
            from utils.excel_handler import load_all_clients

            existing = load_all_clients()
        except Exception:
            existing = []

        yymm = datetime.now().strftime("%y%m")
        pattern = re.compile(rf"^LHM-[RD]{re.escape(yymm)}(\d{{3}})$")
        max_sequence = 0
        for client in existing:
            for key in ("ref_client", "numero_dossier"):
                value = str(client.get(key, "")).strip().upper()
                match = pattern.match(value)
                if match:
                    max_sequence = max(max_sequence, int(match.group(1)))

        next_sequence = max_sequence + 1
        ref = f"LHM-R{yymm}{next_sequence:03d}"
        dossier = f"LHM-D{yymm}{next_sequence:03d}"
        return ref, dossier

    def _apply_statut(self, val):
        """Highlight the correct statut button and update the StringVar."""
        if not hasattr(self, "_statut_var") or not hasattr(self, "_statut_btns"):
            return
        valid = ("En cours", "Accepté", "En circuit", "Annulé")
        val = val if val in valid else "En cours"
        self._statut_var.set(val)
        for s, btn in self._statut_btns.items():
            active = s == val
            btn.configure(
                bg=self._STATUT_COLORS[s] if active else "#F0F0F0",
                fg="white" if active else "#333333",
            )

    def _cancel(self):
        """Cancel editing — ask for confirmation first."""
        confirmed = messagebox.askyesno(
            "Confirmation",
            "Voulez-vous vraiment abandonner ?\nToutes les modifications non enregistrées seront perdues.",
            icon="warning",
        )
        if not confirmed:
            return
        if self.on_save_callback:
            self.on_save_callback()
        elif not self.client_to_edit:
            self._reset_form()

    def _reset_form(self):
        """Reset all form fields"""
        for entry in (self.entry_ref_client, self.entry_numero_dossier):
            try:
                entry.configure(state="normal")
            except Exception:
                pass
            entry.delete(0, tk.END)
            self._restore_placeholder(entry)
            try:
                entry.configure(state="readonly")
            except Exception:
                pass
        self.entry_prenom.delete(0, tk.END)
        self.entry_nom.delete(0, tk.END)
        self._restore_placeholder(self.entry_prenom)
        self._restore_placeholder(self.entry_nom)
        self.entry_date_arrivee.configure(state="normal")
        self.entry_date_arrivee.delete(0, tk.END)
        self.entry_date_arrivee.configure(state="readonly")
        self.entry_date_depart.configure(state="normal")
        self.entry_date_depart.delete(0, tk.END)
        self.entry_date_depart.configure(state="readonly")
        self.entry_duree_sejour.configure(state="normal")
        self.entry_duree_sejour.delete(0, tk.END)
        self.entry_duree_sejour.configure(state="readonly")
        self.entry_telephone.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_code_pays.set(DEFAULT_PHONE_CODE)
        if hasattr(self, "whatsapp_code_var"):
            self.whatsapp_code_var.set(DEFAULT_PHONE_CODE)
        if hasattr(self, "whatsapp_number_var"):
            self.whatsapp_number_var.set("")
        if hasattr(self, "comment_client_text"):
            self.comment_client_text.delete("1.0", tk.END)
        if hasattr(self, "comment_internal_text"):
            self.comment_internal_text.delete("1.0", tk.END)

        # Reset rooming rows
        for rw in getattr(self, "_rooming_widget_rows", []):
            rw["frame"].destroy()
        if hasattr(self, "_rooming_widget_rows"):
            self._rooming_widget_rows.clear()
        self._update_rooming_summary()
        # Reset participant spinboxes
        for attr in ("entry_adultes", "entry_enfants_2_12", "entry_bebes_0_2"):
            if hasattr(self, attr):
                w = getattr(self, attr)
                try:
                    w.configure(state="normal")
                    w.delete(0, tk.END)
                    w.insert(0, "0")
                except Exception:
                    pass
        self.var_enfant.set(False)
        self._toggle_enfant()

        self.combo_type_client.set("Mr")
        self.combo_periode.set("")
        self.combo_restauration.set("")
        self.combo_TypeHebergement.set("")
        self.combo_TypeChambre.set("")
        self.combo_forfait.set("")
        self.combo_circuit.set("")
        self.circuit_info_label.config(text="")
        self.combo_ville_depart.set("")
        self.combo_ville_arrivee.set("")
        self.entry_itineraire_date.configure(state="normal")
        self.entry_itineraire_date.delete(0, tk.END)
        self.entry_itineraire_date.configure(state="readonly")
        self.entry_itineraire_distance.delete(0, tk.END)
        self.entry_itineraire_hebergement.delete(0, tk.END)
        self.itinerary_rows = []
        self._refresh_itinerary_table()
        self.combo_type_hotel_arrivee.set("")
        self.var_accompagnement_guide.set(False)
        self.var_accompagnement_chauffeur.set(False)
        self.var_location_voiture.set("")
