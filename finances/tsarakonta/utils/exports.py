"""Utilitaires d'export pour tableaux Tkinter (Treeview)."""

import json
import os
from tkinter import filedialog, messagebox

import pandas as pd


def treeview_to_dataframe(tree):
    columns = list(tree["columns"])
    rows = [tree.item(item_id, "values") for item_id in tree.get_children()]
    return pd.DataFrame(rows, columns=columns)


def _default_export_dir():
    root_dir = os.path.dirname(os.path.dirname(__file__))
    export_dir = os.path.join(root_dir, "EtatFiFolder", "exports")
    os.makedirs(export_dir, exist_ok=True)
    return export_dir


def _load_header_text():
    root_dir = os.path.dirname(os.path.dirname(__file__))
    settings_file = os.path.join(root_dir, "EtatFiFolder", "settings.json")
    try:
        if not os.path.exists(settings_file):
            return ""
        with open(settings_file, "r", encoding="utf-8") as f:
            settings = json.load(f)
        if not isinstance(settings, dict):
            return ""

        lines = []
        nom = str(settings.get("nom_societe", "")).strip()
        if nom:
            lines.append(nom)

        adresse = str(settings.get("adresse", "")).strip()
        if adresse:
            lines.append(adresse)

        infos = []
        for key in ["NIF", "STAT", "RCS"]:
            value = str(settings.get(key, "")).strip()
            if value:
                infos.append(f"{key}: {value}")
        if infos:
            lines.append(" | ".join(infos))

        return "\n".join(lines)
    except Exception:
        return ""


def export_treeview_to_excel(tree, default_filename, parent=None, header_text=None):
    initial_dir = _default_export_dir()
    path = filedialog.asksaveasfilename(
        parent=parent,
        title="Exporter en Excel",
        defaultextension=".xlsx",
        initialdir=initial_dir,
        initialfile=default_filename,
        filetypes=[("Excel", "*.xlsx")],
    )
    if not path:
        return

    try:
        df = treeview_to_dataframe(tree)
        header = (header_text or "").strip() or _load_header_text()
        if not header:
            df.to_excel(path, index=False)
        else:
            header_lines = [
                line.strip() for line in header.splitlines() if line.strip()
            ]
            startrow = len(header_lines) + 1
            with pd.ExcelWriter(path) as writer:
                sheet_name = "Export"
                df.to_excel(
                    writer, index=False, sheet_name=sheet_name, startrow=startrow
                )

                worksheet = writer.sheets.get(sheet_name)
                if worksheet is not None:
                    for i, line in enumerate(header_lines, start=1):
                        if hasattr(worksheet, "cell"):
                            worksheet.cell(row=i, column=1, value=line)
                        elif hasattr(worksheet, "write"):
                            worksheet.write(i - 1, 0, line)
        messagebox.showinfo(
            "Export réussi", f"Export Excel réussi:\n{path}", parent=parent
        )
    except PermissionError:
        messagebox.showerror(
            "Erreur", "Le fichier est ouvert. Fermez-le puis réessayez.", parent=parent
        )
    except Exception as exc:
        messagebox.showerror(
            "Erreur", f"Impossible d'exporter en Excel:\n{exc}", parent=parent
        )


def export_treeview_to_pdf(
    tree, title, default_filename, parent=None, header_text=None
):
    initial_dir = _default_export_dir()
    path = filedialog.asksaveasfilename(
        parent=parent,
        title="Exporter en PDF",
        defaultextension=".pdf",
        initialdir=initial_dir,
        initialfile=default_filename,
        filetypes=[("PDF", "*.pdf")],
    )
    if not path:
        return

    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas
    except Exception:
        messagebox.showerror(
            "Erreur",
            "Le module 'reportlab' est requis pour l'export PDF.\nInstallez-le avec: pip install reportlab",
            parent=parent,
        )
        return

    try:
        columns = list(tree["columns"])
        rows = [tree.item(item_id, "values") for item_id in tree.get_children()]
        header = (header_text or "").strip() or _load_header_text()
        header_lines = [line.strip() for line in header.splitlines() if line.strip()]

        page_width, page_height = landscape(A4)
        margin = 30
        y_start = page_height - 32
        row_height = 16
        width_available = page_width - (2 * margin)

        col_widths = []
        raw_sum = 0.0
        for col in columns:
            width = float(tree.column(col, "width") or 120)
            col_widths.append(width)
            raw_sum += width

        if raw_sum <= 0:
            col_widths = [width_available / max(1, len(columns))] * len(columns)
        else:
            col_widths = [w * (width_available / raw_sum) for w in col_widths]

        pdf = canvas.Canvas(path, pagesize=landscape(A4))

        def ellipsize(text, max_width, font_name="Helvetica", font_size=9):
            txt = "" if text is None else str(text)
            if pdf.stringWidth(txt, font_name, font_size) <= max_width:
                return txt
            suffix = "..."
            while (
                txt and pdf.stringWidth(txt + suffix, font_name, font_size) > max_width
            ):
                txt = txt[:-1]
            return txt + suffix

        def draw_header(y):
            pdf.setFont("Helvetica-Bold", 12)
            pdf.drawString(margin, y, title)
            y -= 16

            if header_lines:
                pdf.setFont("Helvetica", 9)
                for line in header_lines:
                    pdf.drawString(margin, y, line)
                    y -= 12
                y -= 4

            pdf.setFont("Helvetica-Bold", 9)
            x = margin
            for idx, col in enumerate(columns):
                pdf.drawString(
                    x + 2, y, ellipsize(col, col_widths[idx] - 4, "Helvetica-Bold", 9)
                )
                x += col_widths[idx]
            pdf.line(margin, y - 2, margin + width_available, y - 2)
            return y - 14

        y = draw_header(y_start)
        pdf.setFont("Helvetica", 9)

        for row in rows:
            if y <= margin + row_height:
                pdf.showPage()
                y = draw_header(y_start)
                pdf.setFont("Helvetica", 9)

            x = margin
            for idx, value in enumerate(row):
                pdf.drawString(x + 2, y, ellipsize(value, col_widths[idx] - 4))
                x += col_widths[idx]
            y -= row_height

        pdf.save()
        messagebox.showinfo(
            "Export réussi", f"Export PDF réussi:\n{path}", parent=parent
        )
    except Exception as exc:
        messagebox.showerror(
            "Erreur", f"Impossible d'exporter en PDF:\n{exc}", parent=parent
        )
