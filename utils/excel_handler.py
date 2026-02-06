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
import shutil
from datetime import datetime
from config import CLIENT_EXCEL_PATH, HOTEL_EXCEL_PATH, CLIENT_SHEET_NAME, HOTEL_SHEET_NAME, COTATION_H_SHEET_NAME
from utils.logger import logger
from utils.cache import cached_hotel_data, cached_client_data, invalidate_hotel_cache, invalidate_client_cache

import re


def _parse_num(val):
    """Parse a cell value into int or float, stripping thousand separators and currency text.

    Returns 0 on failure or empty values.
    """
    if val is None or val == "":
        return 0
    if isinstance(val, (int, float)):
        return val
    try:
        s = str(val).strip()
        s = s.replace(',', '').replace(' ', '')
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        if not m:
            return 0
        num_str = m.group(0)
        if '.' in num_str:
            return float(num_str)
        return int(num_str)
    except Exception:
        return 0


def create_backup(filepath):
    """
    Create a backup of Excel file before modification
    
    Args:
        filepath (str): Path to the Excel file to backup
        
    Returns:
        str: Path to backup file, or None if failed
    """
    if not os.path.exists(filepath):
        logger.warning(f"File not found for backup: {filepath}")
        return None
    
    try:
        # Create backups directory
        backup_dir = os.path.join(os.path.dirname(filepath), 'backups')
        os.makedirs(backup_dir, exist_ok=True)
        
        # Create backup filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.basename(filepath)
        backup_path = os.path.join(backup_dir, f"{filename}.{timestamp}.bak")
        
        # Copy file
        shutil.copy2(filepath, backup_path)
        logger.info(f"Backup created: {backup_path}")
        
        return backup_path
    except Exception as e:
        logger.error(f"Failed to create backup for {filepath}: {e}", exc_info=True)
        return None


def save_client_to_excel(client_data):
    """
    Save client data to Excel file

    Args:
        client_data (dict): Client data dictionary

    Returns:
        int: Row number where data was saved, or -1 if failed
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning( "openpyxl not available. Cannot save to Excel.")
        return -1

    # Create file if it doesn't exist
    if not os.path.exists(CLIENT_EXCEL_PATH):
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

        wb.save(CLIENT_EXCEL_PATH)

    # Open existing file
    wb = load_workbook(CLIENT_EXCEL_PATH)
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

    wb.save(CLIENT_EXCEL_PATH)
    
    # Invalidate cache after modification
    invalidate_client_cache()
    
    return last_row


@cached_client_data(ttl_seconds=3600)  # Cache for 1 hour
def load_all_clients():
    """
    Load all client data from Excel file
    Results are cached for 1 hour

    Returns:
        list: List of client dictionaries
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning("openpyxl not available. Cannot load from Excel.")
        return []

    if not os.path.exists(CLIENT_EXCEL_PATH):
        return []

    wb = load_workbook(CLIENT_EXCEL_PATH)
    if CLIENT_SHEET_NAME not in wb.sheetnames:
        return []

    ws = wb[CLIENT_SHEET_NAME]

    clients = []
    # Start from row 2 (skip headers)
    for row in range(2, ws.max_row + 1):
        if ws[f'A{row}'].value is None:
            continue

        client = {
            'row_number': row,
            'timestamp': ws[f'A{row}'].value or '',
            'ref_client': ws[f'B{row}'].value or '',
            'nom': ws[f'C{row}'].value or '',
            'telephone': ws[f'D{row}'].value or '',
            'email': ws[f'E{row}'].value or '',
            'periode': ws[f'F{row}'].value or '',
            'restauration': ws[f'G{row}'].value or '',
            'hebergement': ws[f'H{row}'].value or '',
            'chambre': ws[f'I{row}'].value or '',
            'enfant': ws[f'J{row}'].value or '',
            'age_enfant': ws[f'K{row}'].value or '',
            'forfait': ws[f'L{row}'].value or '',
            'circuit': ws[f'M{row}'].value or ''
        }
        clients.append(client)

    return clients


def update_client_in_excel(row_number, client_data):
    """
    Update client data in Excel file

    Args:
        row_number (int): Row number to update
        client_data (dict): Updated client data dictionary

    Returns:
        bool: True if successful, False otherwise
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning("openpyxl not available. Cannot update Excel.")
        return False

    # Create backup before modifying Excel file
    create_backup(CLIENT_EXCEL_PATH)

    if not os.path.exists(CLIENT_EXCEL_PATH):
        return False

    wb = load_workbook(CLIENT_EXCEL_PATH)
    if CLIENT_SHEET_NAME not in wb.sheetnames:
        return False

    ws = wb[CLIENT_SHEET_NAME]

    # Update data
    ws[f'A{row_number}'] = client_data.get('Timestamp', '')
    ws[f'B{row_number}'] = client_data.get('Ref_Client', '')
    ws[f'C{row_number}'] = client_data.get('Nom', '')
    ws[f'D{row_number}'] = client_data.get('Téléphone', '')
    ws[f'E{row_number}'] = client_data.get('Email', '')
    ws[f'F{row_number}'] = client_data.get('Période', '')
    ws[f'G{row_number}'] = client_data.get('Restauration', '')
    ws[f'H{row_number}'] = client_data.get('Hébergement', '')
    ws[f'I{row_number}'] = client_data.get('Chambre', '')
    ws[f'J{row_number}'] = client_data.get('Enfant', '')
    ws[f'K{row_number}'] = client_data.get('Âge_Enfant', '')
    ws[f'L{row_number}'] = client_data.get('Forfait', '')
    ws[f'M{row_number}'] = client_data.get('Circuit', '')

    wb.save(CLIENT_EXCEL_PATH)
    return True


def delete_client_from_excel(row_number):
    """
    Delete client data from Excel file

    Args:
        row_number (int): Row number to delete

    Returns:
        bool: True if successful, False otherwise
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning("openpyxl not available. Cannot delete from Excel.")
        return False

    if not os.path.exists(CLIENT_EXCEL_PATH):
        return False

    wb = load_workbook(CLIENT_EXCEL_PATH)
    if CLIENT_SHEET_NAME not in wb.sheetnames:
        return False

    ws = wb[CLIENT_SHEET_NAME]

    # Delete row
    ws.delete_rows(row_number)

    wb.save(CLIENT_EXCEL_PATH)
    return True


@cached_hotel_data(ttl_seconds=86400)  # Cache for 24 hours
def load_all_hotels(client_type=None):
    """
    Load all hotel data from Excel file
    Results are cached for 24 hours

    Args:
        client_type (str): Filter by client type ('TO' or 'PBC'), if None load all

    Returns:
        list: List of hotel dictionaries
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning("openpyxl not available. Cannot load from Excel.")
        return []

    if not os.path.exists(HOTEL_EXCEL_PATH):
        return []

    wb = load_workbook(HOTEL_EXCEL_PATH)
    if HOTEL_SHEET_NAME not in wb.sheetnames:
        return []

    ws = wb[HOTEL_SHEET_NAME]

    hotels = []
    # Start from row 2 (skip headers)
    for row in range(2, ws.max_row + 1):
        if ws[f'A{row}'].value is None:
            continue

        hotel = {
            'row_number': row,
            'id': f"{ws[f'A{row}'].value}_{ws[f'B{row}'].value}",  # Combine Ville + HTL as ID
            'nom': ws[f'B{row}'].value or '',  # HTL
            'lieu': ws[f'A{row}'].value or '',  # Ville
            'type_hebergement': 'Hôtel',  # Default type
            'categorie': ws[f'C{row}'].value or '',  # CATÉGORIE
            'type_client': ws[f'N{row}'].value or 'TO',  # TYPE_CLIENT
            'chambre_single': _parse_num(ws[f'E{row}'].value),  # SPL
            'chambre_double': _parse_num(ws[f'F{row}'].value),  # DBL
            'chambre_familiale': _parse_num(ws[f'H{row}'].value),  # FML
            'lit_supp': _parse_num(ws[f'I{row}'].value),  # SUPP
            'day_use': 0,  # Not available in source
            'vignette': 0,  # Not available in source
            'taxe_sejour': 0,  # Not available in source
            'petit_dejeuner': _parse_num(ws[f'K{row}'].value),  # PDJ
            'dejeuner': _parse_num(ws[f'L{row}'].value),  # DJ
            'diner': _parse_num(ws[f'M{row}'].value),  # DR
            'description': f"Unité: {ws[f'D{row}'].value or ''}, Suite: {ws[f'J{row}'].value or ''}",  # UNITÉ + SUITE
            'contact': '',  # Not available in source
            'email': ''  # Not available in source
        }

        # Filter by client type if specified
        if client_type and hotel['type_client'] != client_type:
            continue

        hotels.append(hotel)

    return hotels


def save_hotel_to_excel(hotel_data):
    """
    Save hotel data to Excel file

    Args:
        hotel_data (dict): Hotel data dictionary

    Returns:
        int: Row number where data was saved, or -1 if failed
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning("openpyxl not available. Cannot save to Excel.")
        return -1

    # Create backup before modifying Excel file
    create_backup(HOTEL_EXCEL_PATH)
    # Create file if it doesn't exist
    if not os.path.exists(HOTEL_EXCEL_PATH):
        wb = Workbook()
        wb.save(HOTEL_EXCEL_PATH)

    # Open existing file
    wb = load_workbook(HOTEL_EXCEL_PATH)

    if HOTEL_SHEET_NAME not in wb.sheetnames:
        # If the sheet doesn't exist, create it with existing format headers
        ws = wb.create_sheet(HOTEL_SHEET_NAME)
        
        # Headers in existing format
        headers = ['Ville', 'HTL', 'CATÉGORIE', 'UNITÉ', 'SPL', 'DBL', 'TWINS', 'FML', 'SUPP', 'SUITE', 'PDJ', 'DJ', 'DR']
        
        for i, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=i, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        
        wb.save(HOTEL_EXCEL_PATH)

    # Open existing file again
    wb = load_workbook(HOTEL_EXCEL_PATH)
    ws = wb[HOTEL_SHEET_NAME]

    # Find last empty row (column A)
    last_row = ws.max_row + 1

    # Write data in existing format
    ws[f'A{last_row}'] = hotel_data.get('Lieu', '')  # Ville
    ws[f'B{last_row}'] = hotel_data.get('Nom', '')  # HTL
    ws[f'C{last_row}'] = hotel_data.get('Catégorie', '')  # CATÉGORIE
    ws[f'D{last_row}'] = '$'  # UNITÉ (default)
    ws[f'E{last_row}'] = hotel_data.get('Chambre_Single', 0)  # SPL
    ws[f'F{last_row}'] = hotel_data.get('Chambre_Double', 0)  # DBL
    ws[f'G{last_row}'] = hotel_data.get('Chambre_Double', 0)  # TWINS (same as DBL)
    ws[f'H{last_row}'] = hotel_data.get('Chambre_Familiale', 0)  # FML
    ws[f'I{last_row}'] = hotel_data.get('Lit_Supp', 0)  # SUPP
    ws[f'J{last_row}'] = ''  # SUITE
    ws[f'K{last_row}'] = hotel_data.get('Petit_Déjeuner', 0)  # PDJ
    ws[f'L{last_row}'] = hotel_data.get('Déjeuner', 0)  # DJ
    ws[f'M{last_row}'] = hotel_data.get('Dîner', 0)  # DR

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

    wb.save(HOTEL_EXCEL_PATH)
    return last_row


def update_hotel_in_excel(row_number, hotel_data):
    """
    Update hotel data in Excel file

    Args:
        row_number (int): Row number to update
        hotel_data (dict): Updated hotel data dictionary

    Returns:
        bool: True if successful, False otherwise
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning("openpyxl not available. Cannot update Excel.")
        return False

    # Create backup before modifying Excel file
    create_backup(HOTEL_EXCEL_PATH)

    if not os.path.exists(HOTEL_EXCEL_PATH):
        return False

    wb = load_workbook(HOTEL_EXCEL_PATH)
    if HOTEL_SHEET_NAME not in wb.sheetnames:
        return False

    ws = wb[HOTEL_SHEET_NAME]

    # Update data in existing format
    ws[f'A{row_number}'] = hotel_data.get('Lieu', '')  # Ville
    ws[f'B{row_number}'] = hotel_data.get('Nom', '')  # HTL
    ws[f'C{row_number}'] = hotel_data.get('Catégorie', '')  # CATÉGORIE
    ws[f'D{row_number}'] = '$'  # UNITÉ
    ws[f'E{row_number}'] = hotel_data.get('Chambre_Single', 0)  # SPL
    ws[f'F{row_number}'] = hotel_data.get('Chambre_Double', 0)  # DBL
    ws[f'G{row_number}'] = hotel_data.get('Chambre_Double', 0)  # TWINS
    ws[f'H{row_number}'] = hotel_data.get('Chambre_Familiale', 0)  # FML
    ws[f'I{row_number}'] = hotel_data.get('Lit_Supp', 0)  # SUPP
    ws[f'J{row_number}'] = ''  # SUITE
    ws[f'K{row_number}'] = hotel_data.get('Petit_Déjeuner', 0)  # PDJ
    ws[f'L{row_number}'] = hotel_data.get('Déjeuner', 0)  # DJ
    ws[f'M{row_number}'] = hotel_data.get('Dîner', 0)  # DR

    wb.save(HOTEL_EXCEL_PATH)
    return True


def delete_hotel_from_excel(row_number):
    """
    Delete hotel data from Excel file

    Args:
        row_number (int): Row number to delete

    Returns:
        bool: True if successful, False otherwise
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning( "openpyxl not available. Cannot delete from Excel.")
        return False

    if not os.path.exists(HOTEL_EXCEL_PATH):
        return False

    wb = load_workbook(HOTEL_EXCEL_PATH)
    if HOTEL_SHEET_NAME not in wb.sheetnames:
        return False

    ws = wb[HOTEL_SHEET_NAME]

    # Delete row
    ws.delete_rows(row_number)

    wb.save(HOTEL_EXCEL_PATH)
    return True


def save_hotel_quotation_to_excel(quotation_data):
    """
    Save hotel quotation to COTATION_H sheet in data.xlsx
    
    Args:
        quotation_data (dict): Quotation data with keys:
            - client_id (str): ID/Ref of the client
            - client_name (str): Name of the client
            - hotel_name (str): Name of the hotel
            - city (str): City where the hotel is located
            - total_price (float): Total price of the quotation
            - currency (str): Currency of the price
            - nights (int): Number of nights
            - adults (int): Number of adults
            - children (int): Number of children
            - room_type (str): Type of room
            - meal_plan (str): Meal plan selected
            - period (str): Period/season
            - quote_date (str): Date of the quotation
            
    Returns:
        int: Row number where data was saved, or -1 if failed
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning("openpyxl not available. Cannot save quotation to Excel.")
        return -1
    
    try:
        # Create or load the file
        if not os.path.exists(CLIENT_EXCEL_PATH):
            wb = Workbook()
            ws = wb.active
            ws.title = COTATION_H_SHEET_NAME
        else:
            wb = load_workbook(CLIENT_EXCEL_PATH)
            if COTATION_H_SHEET_NAME not in wb.sheetnames:
                ws = wb.create_sheet(COTATION_H_SHEET_NAME)
            else:
                ws = wb[COTATION_H_SHEET_NAME]
        
        # Add headers if this is the first row
        if ws.max_row == 1 or ws[f'A1'].value is None:
            headers = [
                'Date', 'ID_Client', 'Nom_Client', 'Hôtel', 'Ville', 
                'Nuits', 'Type_Chambre', 'Adultes', 'Enfants', 
                'Plan_Repas', 'Période', 'Total_Devise', 'Devise'
            ]
            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=col)
                cell.value = header
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # Find next row
        next_row = ws.max_row + 1
        
        # Save data
        ws[f'A{next_row}'] = quotation_data.get('quote_date', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        ws[f'B{next_row}'] = quotation_data.get('client_id', '')
        ws[f'C{next_row}'] = quotation_data.get('client_name', '')
        ws[f'D{next_row}'] = quotation_data.get('hotel_name', '')
        ws[f'E{next_row}'] = quotation_data.get('city', '')
        ws[f'F{next_row}'] = quotation_data.get('nights', 0)
        ws[f'G{next_row}'] = quotation_data.get('room_type', '')
        ws[f'H{next_row}'] = quotation_data.get('adults', 0)
        ws[f'I{next_row}'] = quotation_data.get('children', 0)
        ws[f'J{next_row}'] = quotation_data.get('meal_plan', '')
        ws[f'K{next_row}'] = quotation_data.get('period', '')
        ws[f'L{next_row}'] = _parse_num(quotation_data.get('total_price', 0))
        ws[f'M{next_row}'] = quotation_data.get('currency', 'Ariary')
        
        # Adjust column widths
        ws.column_dimensions['A'].width = 16
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 8
        ws.column_dimensions['G'].width = 14
        ws.column_dimensions['H'].width = 8
        ws.column_dimensions['I'].width = 8
        ws.column_dimensions['J'].width = 18
        ws.column_dimensions['K'].width = 12
        ws.column_dimensions['L'].width = 14
        ws.column_dimensions['M'].width = 10
        
        wb.save(CLIENT_EXCEL_PATH)
        logger.info(f"Quotation saved to row {next_row} in {COTATION_H_SHEET_NAME}")
        return next_row
        
    except Exception as e:
        logger.error(f"Failed to save quotation to Excel: {e}", exc_info=True)
        return -1


def load_all_hotel_quotations():
    """
    Load all hotel quotations from COTATION_H sheet
    
    Returns:
        list: List of quotation dictionaries
    """
    if not OPENPYXL_AVAILABLE:
        logger.warning("openpyxl not available. Cannot load from Excel.")
        return []
    
    if not os.path.exists(CLIENT_EXCEL_PATH):
        return []
    
    try:
        wb = load_workbook(CLIENT_EXCEL_PATH)
        if COTATION_H_SHEET_NAME not in wb.sheetnames:
            return []
        
        ws = wb[COTATION_H_SHEET_NAME]
        
        quotations = []
        # Start from row 2 (skip headers)
        for row in range(2, ws.max_row + 1):
            if ws[f'A{row}'].value is None:
                continue
            
            quotation = {
                'row_number': row,
                'quote_date': ws[f'A{row}'].value or '',
                'client_id': ws[f'B{row}'].value or '',
                'client_name': ws[f'C{row}'].value or '',
                'hotel_name': ws[f'D{row}'].value or '',
                'city': ws[f'E{row}'].value or '',
                'nights': _parse_num(ws[f'F{row}'].value),
                'room_type': ws[f'G{row}'].value or '',
                'adults': _parse_num(ws[f'H{row}'].value),
                'children': _parse_num(ws[f'I{row}'].value),
                'meal_plan': ws[f'J{row}'].value or '',
                'period': ws[f'K{row}'].value or '',
                'total_price': _parse_num(ws[f'L{row}'].value),
                'currency': ws[f'M{row}'].value or 'Ariary'
            }
            quotations.append(quotation)
        
        return quotations
        
    except Exception as e:
        logger.error(f"Failed to load quotations from Excel: {e}", exc_info=True)
        return []


def get_quotations_grouped_by_client():
    """
    Get all hotel quotations grouped by client with subtotals
    
    Returns:
        dict: Dictionary with client_id as key and list of quotations as value,
              plus 'total' key with grand total per client
    """
    quotations = load_all_hotel_quotations()
    grouped = {}
    
    for quotation in quotations:
        client_id = quotation['client_id']
        if client_id not in grouped:
            grouped[client_id] = {
                'client_name': quotation['client_name'],
                'quotations': [],
                'total': 0,
                'currency': quotation['currency']
            }
        
        grouped[client_id]['quotations'].append(quotation)
        grouped[client_id]['total'] += quotation['total_price']
    
    return grouped


def get_quotations_by_city():
    """
    Get all hotel quotations grouped by city with subtotals
    
    Returns:
        dict: Dictionary with city as key and list of quotations as value,
              plus 'total' key with grand total per city
    """
    quotations = load_all_hotel_quotations()
    grouped = {}
    
    for quotation in quotations:
        city = quotation['city']
        if city not in grouped:
            grouped[city] = {
                'quotations': [],
                'total': 0,
                'currency': quotation['currency']
            }
        
        grouped[city]['quotations'].append(quotation)
        grouped[city]['total'] += quotation['total_price']
    
    return grouped
