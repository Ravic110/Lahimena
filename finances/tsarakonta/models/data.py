"""
Gestion des données Excel et du Plan Comptable Général
"""

import os
from tkinter import messagebox

import pandas as pd

from finances.tsarakonta.config import CONFIG


class DataManager:
    """Gestion des données Excel - chargement et sauvegarde"""

    @staticmethod
    def charger_feuille(fichier, feuille):
        """
        Charge une feuille Excel spécifique

        Args:
            fichier: chemin du fichier Excel
            feuille: nom de la feuille

        Returns:
            DataFrame ou None si erreur
        """
        if not os.path.exists(fichier):
            return None
        try:
            xl_file = pd.ExcelFile(fichier)
            if feuille in xl_file.sheet_names:
                return pd.read_excel(fichier, sheet_name=feuille)
            return None
        except Exception as e:
            print(f"Erreur lecture {feuille}: {e}")
            return None

    @staticmethod
    def sauvegarder_df(df, fichier, feuille):
        """
        Sauvegarde un DataFrame dans une feuille Excel

        Args:
            df: DataFrame à sauvegarder
            fichier: chemin du fichier Excel
            feuille: nom de la feuille

        Returns:
            True si succès, False sinon
        """
        try:
            if os.path.exists(fichier):
                with pd.ExcelWriter(
                    fichier,
                    engine="openpyxl",
                    mode="a",
                    if_sheet_exists="replace",
                ) as writer:
                    df.to_excel(writer, sheet_name=feuille, index=False)
            else:
                df.to_excel(fichier, sheet_name=feuille, index=False)
            return True
        except PermissionError:
            messagebox.showerror("Erreur", "FERMEZ EXCEL avant de sauvegarder !")
            return False
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur sauvegarde: {str(e)}")
            return False

    @staticmethod
    def synchroniser_lahimena_vers_journal(fichier, feuille_journal=None):
        """
        Synchronise les feuilles Lahimena compatibles vers la feuille Journal.

        Sources supportées:
          - COTATION_H
          - COTATION_FRAIS_COL (si présente)

        Returns:
            int: nombre de lignes ajoutées
        """
        if feuille_journal is None:
            feuille_journal = CONFIG.get("feuille_journal", "Journal")

        if not os.path.exists(fichier):
            return 0

        try:
            xl = pd.ExcelFile(fichier)
            sheet_names = set(xl.sheet_names)
        except Exception:
            return 0

        journal_cols = CONFIG["colonnes_journal"]
        journal_df = DataManager.charger_feuille(fichier, feuille_journal)
        if journal_df is None:
            journal_df = pd.DataFrame(columns=journal_cols)
        else:
            for col in journal_cols:
                if col not in journal_df.columns:
                    journal_df[col] = 0.0 if col in ("MontantDébit", "MontantCrédit") else ""

        def _normalize_amount(value):
            if pd.isna(value):
                return 0.0
            if isinstance(value, (int, float)):
                return float(value)
            txt = str(value).strip().replace(" ", "").replace("\u00a0", "")
            txt = txt.replace(",", ".")
            try:
                return float(txt)
            except Exception:
                return 0.0

        def _year_from_date(value):
            try:
                dt = pd.to_datetime(value, errors="coerce")
                if pd.isna(dt):
                    return ""
                return str(int(dt.year))
            except Exception:
                return ""

        existing_labels = set(journal_df.get("Libellé", pd.Series(dtype=str)).astype(str))

        new_rows = []

        if "COTATION_H" in sheet_names:
            cot_h = pd.read_excel(fichier, sheet_name="COTATION_H")
            for i, row in cot_h.iterrows():
                excel_row = i + 2
                signature = f"[LHM:H:{excel_row}]"
                if any(signature in lbl for lbl in existing_labels):
                    continue

                amount = _normalize_amount(row.get("Total_Devise", 0))
                if amount <= 0:
                    continue

                date_value = row.get("Date", "")
                client = str(row.get("Nom_Client", "") or "").strip()
                hotel = str(row.get("Hôtel", "") or "").strip()
                city = str(row.get("Ville", "") or "").strip()
                currency = str(row.get("Devise", "Ariary") or "").strip()

                label_core = " / ".join(
                    part for part in [client, hotel, city] if part
                ) or "Cotation hôtel"
                libelle = f"{signature} {label_core} ({currency})"

                new_rows.append(
                    {
                        "Date": date_value,
                        "Libellé": libelle,
                        "DateValeur": date_value,
                        "MontantDébit": amount,
                        "MontantCrédit": amount,
                        "CompteDébit": "411000",
                        "CompteCrédit": "706000",
                        "Année": _year_from_date(date_value),
                    }
                )

        if "COTATION_FRAIS_COL" in sheet_names:
            cot_f = pd.read_excel(fichier, sheet_name="COTATION_FRAIS_COL")
            lowered_headers = {c.lower(): c for c in cot_f.columns}

            def _pick_header(*keywords):
                for lower, original in lowered_headers.items():
                    if all(k in lower for k in keywords):
                        return original
                return None

            amount_col = (
                _pick_header("total")
                or _pick_header("montant")
                or _pick_header("prix")
            )
            date_col = _pick_header("date")
            desg_col = _pick_header("designation") or _pick_header("libelle")
            prest_col = _pick_header("prestataire")
            devis_col = _pick_header("devise")

            for i, row in cot_f.iterrows():
                excel_row = i + 2
                signature = f"[LHM:F:{excel_row}]"
                if any(signature in lbl for lbl in existing_labels):
                    continue

                amount = _normalize_amount(row.get(amount_col, 0) if amount_col else 0)
                if amount <= 0:
                    continue

                date_value = row.get(date_col, "") if date_col else ""
                designation = str(row.get(desg_col, "") if desg_col else "").strip()
                prestataire = str(row.get(prest_col, "") if prest_col else "").strip()
                devise = str(row.get(devis_col, "Ariary") if devis_col else "Ariary").strip()

                label_core = " / ".join(
                    part for part in [designation, prestataire] if part
                ) or "Frais collectif"
                libelle = f"{signature} {label_core} ({devise})"

                new_rows.append(
                    {
                        "Date": date_value,
                        "Libellé": libelle,
                        "DateValeur": date_value,
                        "MontantDébit": amount,
                        "MontantCrédit": amount,
                        "CompteDébit": "625000",
                        "CompteCrédit": "401000",
                        "Année": _year_from_date(date_value),
                    }
                )

        if not new_rows:
            return 0

        append_df = pd.DataFrame(new_rows, columns=journal_cols)
        result = pd.concat([journal_df[journal_cols], append_df], ignore_index=True)
        DataManager.sauvegarder_df(result, fichier, feuille_journal)
        return len(new_rows)


class PCGManager:
    """Gestion du Plan Comptable Général"""

    @staticmethod
    def charger_pcg(fichier=None):
        """
        Charge le Plan Comptable Général depuis Excel

        Args:
            fichier: chemin du fichier Excel contenant le PCG

        Returns:
            tuple: (liste_comptes, liste_numeros, dict_comptes)
        """
        pcg_comptes = []
        pcg_numeros = []
        pcg_dict = {}

        # Prefer a dedicated PCG file if configured, otherwise fall back to the provided file
        pcg_df = None
        pcg_file = CONFIG.get("fichier_pcg")
        if pcg_file and os.path.exists(pcg_file):
            pcg_df = DataManager.charger_feuille(pcg_file, CONFIG["feuille_pcg"])

        if pcg_df is None and fichier:
            # fallback: try the provided file (e.g. EtatFidata.xlsx)
            pcg_df = DataManager.charger_feuille(fichier, CONFIG["feuille_pcg"])

        if pcg_df is not None and len(pcg_df.columns) >= 2:
            col_numero = pcg_df.columns[0]
            col_libelle = pcg_df.columns[1]

            for _, row in pcg_df.iterrows():
                numero = str(row[col_numero]).strip()
                libelle = str(row[col_libelle]).strip()

                if numero and libelle:
                    affichage = f"{numero} - {libelle}"
                    pcg_comptes.append(affichage)
                    pcg_numeros.append(numero)
                    pcg_dict[numero] = libelle

            print(f"✅ PCG chargé: {len(pcg_comptes)} comptes")
        else:
            print("ℹ️ PCG vide - saisie manuelle autorisée")

        return pcg_comptes, pcg_numeros, pcg_dict
