"""
Client quotation history GUI component
"""

import datetime
import os
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from config import (
    BUTTON_BLUE,
    BUTTON_FONT,
    BUTTON_GREEN,
    BUTTON_RED,
    DEVIS_FOLDER,
    INPUT_BG_COLOR,
    LABEL_FONT,
    MAIN_BG_COLOR,
    TEXT_COLOR,
    TITLE_FONT,
)
from utils.logger import logger


class ClientQuotationHistory:
    """Display history of generated client quotations"""

    def __init__(self, parent):
        self.parent = parent
        self._create_interface()
        self._refresh_list()

    def _create_interface(self):
        for widget in self.parent.winfo_children():
            widget.destroy()

        title = tk.Label(
            self.parent,
            text="HISTORIQUE DES DEVIS CLIENTS",
            font=TITLE_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
        )
        title.pack(pady=(20, 10))

        main_frame = tk.Frame(self.parent, bg=MAIN_BG_COLOR)
        main_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        style = ttk.Style()
        style.configure(
            "Treeview",
            background=INPUT_BG_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=INPUT_BG_COLOR,
        )
        style.map("Treeview", background=[("selected", BUTTON_GREEN)])

        history_frame = tk.LabelFrame(
            main_frame,
            text="Devis clients",
            font=LABEL_FONT,
            fg=TEXT_COLOR,
            bg=MAIN_BG_COLOR,
            padx=10,
            pady=10,
        )
        history_frame.pack(fill="both", expand=True, pady=(0, 10))

        columns = ("filename", "date", "size")
        self.history_tree = ttk.Treeview(
            history_frame, columns=columns, height=12, show="tree headings"
        )
        self.history_tree.heading("#0", text="Nom du fichier")
        self.history_tree.heading("date", text="Date")
        self.history_tree.heading("size", text="Taille")

        self.history_tree.column("#0", width=300)
        self.history_tree.column("date", width=140)
        self.history_tree.column("size", width=90)

        self.history_tree.pack(fill="both", expand=True, side="left")

        scrollbar = ttk.Scrollbar(
            history_frame, orient="vertical", command=self.history_tree.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.history_tree.configure(yscroll=scrollbar.set)

        buttons_frame = tk.Frame(main_frame, bg=MAIN_BG_COLOR)
        buttons_frame.pack(fill="x", pady=(0, 10))

        tk.Button(
            buttons_frame,
            text="🔄 Rafraîchir",
            command=self._refresh_list,
            bg=BUTTON_BLUE,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8,
        ).pack(side="left", padx=5)

        tk.Button(
            buttons_frame,
            text="📂 Ouvrir devis",
            command=self._open_selected,
            bg=BUTTON_GREEN,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8,
        ).pack(side="left", padx=5)

        tk.Button(
            buttons_frame,
            text="🗑️ Supprimer devis",
            command=self._delete_selected,
            bg=BUTTON_RED,
            fg="white",
            font=BUTTON_FONT,
            padx=15,
            pady=8,
        ).pack(side="left", padx=5)

    def _refresh_list(self):
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)

        if not os.path.exists(DEVIS_FOLDER):
            logger.warning(f"Devis folder not found: {DEVIS_FOLDER}")
            return

        pdf_files = [
            f
            for f in os.listdir(DEVIS_FOLDER)
            if f.startswith("DEVIS_CLIENT_") and f.endswith(".pdf")
        ]
        pdf_files.sort(reverse=True)

        for filename in pdf_files:
            filepath = os.path.join(DEVIS_FOLDER, filename)
            try:
                file_stat = os.stat(filepath)
                file_size = f"{file_stat.st_size / 1024:.1f} KB"
                file_date = datetime.datetime.fromtimestamp(
                    file_stat.st_mtime
                ).strftime("%d/%m/%Y %H:%M")
                self.history_tree.insert(
                    "",
                    "end",
                    text=filename,
                    values=(filename, file_date, file_size),
                )
            except Exception as e:
                logger.error(f"Error processing client quotation file {filename}: {e}")

    def _get_selected_file(self):
        selection = self.history_tree.selection()
        if not selection:
            return None
        item = self.history_tree.item(selection[0])
        filename = item.get("text") or ""
        if not filename:
            return None
        return os.path.join(DEVIS_FOLDER, filename)

    def _open_selected(self):
        filepath = self._get_selected_file()
        if not filepath:
            messagebox.showwarning("Aucun devis", "Veuillez sélectionner un devis.")
            return
        if not os.path.exists(filepath):
            messagebox.showwarning("Fichier introuvable", f"Le fichier PDF est introuvable :\n{filepath}")
            return
        try:
            if os.name == "nt":
                os.startfile(filepath)
            elif os.name == "posix":
                subprocess.run(["xdg-open", filepath], check=False)
        except Exception as e:
            logger.warning(f"Could not open quotation file: {e}")
            messagebox.showwarning(
                "⚠️ Ouverture impossible",
                "Le fichier existe mais n'a pas pu s'ouvrir automatiquement.",
            )

    def _delete_selected(self):
        filepath = self._get_selected_file()
        if not filepath:
            messagebox.showwarning("Aucun devis", "Veuillez sélectionner un devis.")
            return
        if not messagebox.askyesno(
            "Confirmation", "Supprimer définitivement ce devis ?"
        ):
            return
        try:
            os.remove(filepath)
            self._refresh_list()
        except Exception as e:
            logger.error(f"Failed to delete quotation file: {e}")
            messagebox.showerror("Erreur", "Impossible de supprimer le devis.")
