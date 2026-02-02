"""
Excel file handling utilities
"""

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

import os
from config import DATA_EXCEL_PATH, CLIENT_SHEET_NAME


def save_client_to_excel(client_data):
    """
    Save client data to Excel file

    Args:
        client_data (dict): Client data dictionary

    Returns:
        int: Row number where data was saved, or -1 if failed
    """
    if not OPENPYXL_AVAILABLE:
        print("Warning: openpyxl not available. Cannot save to Excel.")
        return -1

    # Create file if it doesn't exist
    if not os.path.exists(DATA_EXCEL_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = CLIENT_SHEET_NAME

        # Headers (A1:M1)
        headers = ['Date', 'Réf. Client', 'Nom', 'Téléphone', 'Email', 'Période',
                  'Restauration', 'Hébergement', 'Chambre', 'Enfant', 'Âge Enfant',
                  'Forfait', 'Circuit']

        for i, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=i, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")

        wb.save(DATA_EXCEL_PATH)

    # Open existing file
    wb = load_workbook(DATA_EXCEL_PATH)
    if CLIENT_SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(CLIENT_SHEET_NAME)
    else:
        ws = wb[CLIENT_SHEET_NAME]

    # Find last empty row (column A)
    last_row = 2
    while ws[f'A{last_row}'].value is not None:
        last_row += 1

    # Write data
    ws[f'A{last_row}'] = client_data.get('Timestamp', '')
    ws[f'B{last_row}'] = client_data.get('Ref_Client', '')
    ws[f'C{last_row}'] = client_data.get('Nom', '')
    ws[f'D{last_row}'] = client_data.get('Téléphone', '')
    ws[f'E{last_row}'] = client_data.get('Email', '')
    ws[f'F{last_row}'] = client_data.get('Période', '')
    ws[f'G{last_row}'] = client_data.get('Restauration', '')
    ws[f'H{last_row}'] = client_data.get('Hébergement', '')
    ws[f'I{last_row}'] = client_data.get('Chambre', '')
    ws[f'J{last_row}'] = client_data.get('Enfant', '')
    ws[f'K{last_row}'] = client_data.get('Âge_Enfant', '')
    ws[f'L{last_row}'] = client_data.get('Forfait', '')
    ws[f'M{last_row}'] = client_data.get('Circuit', '')

    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column_letter].width = min(max_length + 2, 25)

    wb.save(DATA_EXCEL_PATH)
    return last_row