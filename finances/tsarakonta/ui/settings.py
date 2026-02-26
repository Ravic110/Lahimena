"""
Fenêtre de paramétrage des informations d'en-tête (société)
Sauvegarde/charge les préférences dans EtatFiFolder/settings.json
"""

import json
import os
import tkinter as tk
from tkinter import messagebox, ttk

STATE_FOLDER = os.path.join(os.path.dirname(os.path.dirname(__file__)), "EtatFiFolder")
SETTINGS_FILE = os.path.join(STATE_FOLDER, "settings.json")

DEFAULT = {"nom_societe": "", "adresse": "", "NIF": "", "STAT": "", "RCS": ""}


def load_header_settings():
    os.makedirs(STATE_FOLDER, exist_ok=True)
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    merged = DEFAULT.copy()
                    merged.update(data)
                    return merged
    except Exception:
        pass
    return DEFAULT.copy()


def save_header_settings(values):
    os.makedirs(STATE_FOLDER, exist_ok=True)
    merged = DEFAULT.copy()
    if isinstance(values, dict):
        merged.update(values)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)


def format_header_text(settings):
    data = DEFAULT.copy()
    if isinstance(settings, dict):
        data.update(settings)

    lines = []
    nom = str(data.get("nom_societe", "")).strip()
    if nom:
        lines.append(nom)

    adresse = str(data.get("adresse", "")).strip()
    if adresse:
        lines.append(adresse)

    infos = []
    for key in ["NIF", "STAT", "RCS"]:
        value = str(data.get(key, "")).strip()
        if value:
            infos.append(f"{key}: {value}")

    if infos:
        lines.append(" | ".join(infos))

    return "\n".join(lines)


class SettingsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        os.makedirs(STATE_FOLDER, exist_ok=True)
        self.title("Paramétrage")
        self.geometry("480x220")
        self._load()
        self._build()

    def _load(self):
        self.settings = load_header_settings()

    def _build(self):
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Nom société:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.nom_var = tk.StringVar(value=self.settings.get("nom_societe", ""))
        ttk.Entry(frm, textvariable=self.nom_var, width=50).grid(
            row=0, column=1, pady=4
        )

        ttk.Label(frm, text="Adresse:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.addr_var = tk.StringVar(value=self.settings.get("adresse", ""))
        ttk.Entry(frm, textvariable=self.addr_var, width=50).grid(
            row=1, column=1, pady=4
        )

        ttk.Label(frm, text="NIF:").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.nif_var = tk.StringVar(value=self.settings.get("NIF", ""))
        ttk.Entry(frm, textvariable=self.nif_var, width=30).grid(
            row=2, column=1, sticky=tk.W, pady=4
        )

        ttk.Label(frm, text="STAT:").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.stat_var = tk.StringVar(value=self.settings.get("STAT", ""))
        ttk.Entry(frm, textvariable=self.stat_var, width=30).grid(
            row=3, column=1, sticky=tk.W, pady=4
        )

        ttk.Label(frm, text="RCS:").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.rcs_var = tk.StringVar(value=self.settings.get("RCS", ""))
        ttk.Entry(frm, textvariable=self.rcs_var, width=30).grid(
            row=4, column=1, sticky=tk.W, pady=4
        )

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="Enregistrer", command=self.save).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(btn_frame, text="Annuler", command=self.destroy).pack(
            side=tk.LEFT, padx=6
        )

    def save(self):
        data = {
            "nom_societe": self.nom_var.get(),
            "adresse": self.addr_var.get(),
            "NIF": self.nif_var.get(),
            "STAT": self.stat_var.get(),
            "RCS": self.rcs_var.get(),
        }
        try:
            save_header_settings(data)
            messagebox.showinfo("Paramétrage", "Paramètres enregistrés")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'enregistrer: {e}")
