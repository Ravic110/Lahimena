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


def _ensure_headers(ws, headers, header_style=None):
    header_map = {}
    max_col = ws.max_column if ws.max_column and ws.max_column > 0 else 0
    for col in range(1, max_col + 1):
        value = ws.cell(row=1, column=col).value
        if value:
            header_map[str(value).strip()] = col

    if max_col == 1 and not header_map and ws.cell(row=1, column=1).value is None:
        max_col = 0

    next_col = max_col + 1 if max_col else 1
    for header in headers:
        if header not in header_map:
            cell = ws.cell(row=1, column=next_col, value=header)
            if header_style:
                if header_style.get("font"):
                    cell.font = header_style["font"]
                if header_style.get("fill"):
                    cell.fill = header_style["fill"]
                if header_style.get("alignment"):
                    cell.alignment = header_style["alignment"]
            header_map[header] = next_col
            next_col += 1

    return header_map


def _first_available(data, keys, default=""):
    for key in keys:
        if key in data and data.get(key) not in (None, ""):
            return data.get(key)
    return default


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
    client_headers = [
        "Date", "Réf. Client", "Type Client", "Prénom", "Nom",
        "Date Arrivée", "Date Départ", "Durée Séjour",
        "Nombre Participants", "Nombre Adultes",
        "Enfants 2-12", "Bébés 0-2",
        "Téléphone", "Téléphone WhatsApp", "Email",
        "Période", "Restauration", "Hébergement", "Chambre",
        "Enfant", "Âge Enfant", "Forfait", "Circuit",
        "SGL", "DBL", "TWN", "TPL", "FML"
    ]
    client_header_style = {
        "font": Font(bold=True, color="FFFFFF"),
        "fill": PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid"),
        "alignment": Alignment(horizontal="center")
    }

    if not os.path.exists(CLIENT_EXCEL_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = CLIENT_SHEET_NAME

        _ensure_headers(ws, client_headers, client_header_style)

        wb.save(CLIENT_EXCEL_PATH)

    # Open existing file
    wb = load_workbook(CLIENT_EXCEL_PATH)
    if CLIENT_SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(CLIENT_SHEET_NAME)
    else:
        ws = wb[CLIENT_SHEET_NAME]

    header_map = _ensure_headers(ws, client_headers, client_header_style)

    # Find last empty row (column A)
    last_row = 2
    while ws[f'A{last_row}'].value is not None:
        last_row += 1

    # Write data by headers for compatibility and extensibility
    field_map = {
        "Date": ["Timestamp", "Date", "date_jour"],
        "Réf. Client": ["Ref_Client", "ref_client"],
        "Type Client": ["Type_Client", "type_client"],
        "Prénom": ["Prénom", "prenom"],
        "Nom": ["Nom", "nom"],
        "Date Arrivée": ["Date_Arrivée", "date_arrivee"],
        "Date Départ": ["Date_Départ", "date_depart"],
        "Durée Séjour": ["Durée_Séjour", "duree_sejour"],
        "Nombre Participants": ["Nombre_Participants", "nombre_participants"],
        "Nombre Adultes": ["Nombre_Adultes", "nombre_adultes"],
        "Enfants 2-12": ["Enfants_2_12_ans", "nombre_enfants_2_12"],
        "Bébés 0-2": ["Bébés_0_2_ans", "nombre_bebes_0_2"],
        "Téléphone": ["Téléphone", "telephone"],
        "Téléphone WhatsApp": ["Téléphone_WhatsApp", "telephone_whatsapp"],
        "Email": ["Email", "email"],
        "Période": ["Période", "periode"],
        "Restauration": ["Restauration", "restauration"],
        "Hébergement": ["Hébergement", "hebergement"],
        "Chambre": ["Chambre", "chambre"],
        "Enfant": ["Enfant", "enfant"],
        "Âge Enfant": ["Âge_Enfant", "age_enfant"],
        "Forfait": ["Forfait", "forfait"],
        "Circuit": ["Circuit", "circuit"],
        "SGL": ["SGL_Count", "sgl_count"],
        "DBL": ["DBL_Count", "dbl_count"],
        "TWN": ["TWN_Count", "twn_count"],
        "TPL": ["TPL_Count", "tpl_count"],
        "FML": ["FML_Count", "fml_count"]
    }
    for header, keys in field_map.items():
        col = header_map.get(header)
        if col:
            ws.cell(row=last_row, column=col, value=_first_available(client_data, keys, ""))

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

    # Also store extended infos in dedicated sheet
    _save_client_infos_to_excel(client_data)
    
    return last_row


def _save_client_infos_to_excel(client_data):
    """Save extended client infos to the INFOS_CLIENTS sheet"""
    if not OPENPYXL_AVAILABLE:
        logger.warning("openpyxl not available. Cannot save client infos.")
        return False

    infos_headers = [
        "Date", "Réf. Client", "Numéro Dossier", "Type Client", "Prénom", "Nom",
        "Date Arrivée", "Date Départ", "Durée Séjour",
        "Nombre Participants", "Nombre Adultes",
        "Enfants 2-12", "Bébés 0-2",
        "Téléphone", "Téléphone WhatsApp", "Email",
        "Période", "Restauration", "Hébergement", "Chambre",
        "Enfant", "Âge Enfant", "Forfait", "Circuit",
        "Type Circuit", "Ville Départ", "Ville Arrivée", "Type Hôtel Arrivée",
        "SGL", "DBL", "TWN", "TPL", "FML"
    ]
    infos_header_style = {
        "font": Font(bold=True, color="FFFFFF"),
        "fill": PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid"),
        "alignment": Alignment(horizontal="center")
    }

    if not os.path.exists(CLIENT_EXCEL_PATH):
        wb = Workbook()
        wb.save(CLIENT_EXCEL_PATH)

    wb = load_workbook(CLIENT_EXCEL_PATH)
    if CLIENT_INFOS_SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(CLIENT_INFOS_SHEET_NAME)
    else:
        ws = wb[CLIENT_INFOS_SHEET_NAME]

    header_map = _ensure_headers(ws, infos_headers, infos_header_style)

    # Find existing row by ref client
    ref_client = _first_available(client_data, ["Ref_Client", "ref_client"], "")
    target_row = None
    if ref_client:
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=header_map["Réf. Client"]).value == ref_client:
                target_row = row
                break

    if target_row is None:
        target_row = 2
        while ws[f'A{target_row}'].value is not None:
            target_row += 1

    field_map = {
        "Date": ["Timestamp", "Date", "date_jour"],
        "Réf. Client": ["Ref_Client", "ref_client"],
        "Numéro Dossier": ["Numero_Dossier", "numero_dossier"],
        "Type Client": ["Type_Client", "type_client"],
        "Prénom": ["Prénom", "prenom"],
        "Nom": ["Nom", "nom"],
        "Date Arrivée": ["Date_Arrivée", "date_arrivee"],
        "Date Départ": ["Date_Départ", "date_depart"],
        "Durée Séjour": ["Durée_Séjour", "duree_sejour"],
        "Nombre Participants": ["Nombre_Participants", "nombre_participants"],
        "Nombre Adultes": ["Nombre_Adultes", "nombre_adultes"],
        "Enfants 2-12": ["Enfants_2_12_ans", "nombre_enfants_2_12"],
        "Bébés 0-2": ["Bébés_0_2_ans", "nombre_bebes_0_2"],
        "Téléphone": ["Téléphone", "telephone"],
        "Téléphone WhatsApp": ["Téléphone_WhatsApp", "telephone_whatsapp"],
        "Email": ["Email", "email"],
        "Période": ["Période", "periode"],
        "Restauration": ["Restauration", "restauration"],
        "Hébergement": ["Hébergement", "hebergement"],
        "Chambre": ["Chambre", "chambre"],
        "Enfant": ["Enfant", "enfant"],
        "Âge Enfant": ["Âge_Enfant", "age_enfant"],
        "Forfait": ["Forfait", "forfait"],
        "Circuit": ["Circuit", "circuit"],
        "Type Circuit": ["Type_Circuit", "type_circuit"],
        "Ville Départ": ["Ville_Depart", "ville_depart"],
        "Ville Arrivée": ["Ville_Arrivee", "ville_arrivee"],
        "Type Hôtel Arrivée": ["Type_Hotel_Arrivee", "type_hotel_arrivee"],
        "SGL": ["SGL_Count", "sgl_count"],
        "DBL": ["DBL_Count", "dbl_count"],
        "TWN": ["TWN_Count", "twn_count"],
        "TPL": ["TPL_Count", "tpl_count"],
        "FML": ["FML_Count", "fml_count"]
    }
    for header, keys in field_map.items():
        col = header_map.get(header)
        if col:
            ws.cell(row=target_row, column=col, value=_first_available(client_data, keys, ""))

    wb.save(CLIENT_EXCEL_PATH)
    return True


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
    header_map = _ensure_headers(ws, [])

    clients = []
    infos_map = _load_client_infos_map()
    # Start from row 2 (skip headers)
    for row in range(2, ws.max_row + 1):
        if ws[f'A{row}'].value is None:
            continue

        def _cell(header, fallback_cell=None):
            col = header_map.get(header)
            if col:
                return ws.cell(row=row, column=col).value
            if fallback_cell:
                return ws[fallback_cell].value
            return None

        client = {
            'row_number': row,
            'timestamp': _cell("Date", f'A{row}') or '',
            'ref_client': _cell("Réf. Client", f'B{row}') or '',
            'type_client': _cell("Type Client") or '',
            'prenom': _cell("Prénom") or '',
            'nom': _cell("Nom", f'C{row}') or '',
            'date_arrivee': _cell("Date Arrivée") or '',
            'date_depart': _cell("Date Départ") or '',
            'duree_sejour': _cell("Durée Séjour") or '',
            'nombre_participants': _cell("Nombre Participants") or '',
            'nombre_adultes': _cell("Nombre Adultes") or '',
            'nombre_enfants_2_12': _cell("Enfants 2-12") or '',
            'nombre_bebes_0_2': _cell("Bébés 0-2") or '',
            'telephone': _cell("Téléphone", f'D{row}') or '',
            'telephone_whatsapp': _cell("Téléphone WhatsApp") or '',
            'email': _cell("Email", f'E{row}') or '',
            'periode': _cell("Période", f'F{row}') or '',
            'restauration': _cell("Restauration", f'G{row}') or '',
            'hebergement': _cell("Hébergement", f'H{row}') or '',
            'chambre': _cell("Chambre", f'I{row}') or '',
            'enfant': _cell("Enfant", f'J{row}') or '',
            'age_enfant': _cell("Âge Enfant", f'K{row}') or '',
            'forfait': _cell("Forfait", f'L{row}') or '',
            'circuit': _cell("Circuit", f'M{row}') or '',
            'sgl_count': _cell("SGL") or '',
            'dbl_count': _cell("DBL") or '',
            'twn_count': _cell("TWN") or '',
            'tpl_count': _cell("TPL") or '',
            'fml_count': _cell("FML") or ''
        }
        ref_key = client.get('ref_client')
        if ref_key and ref_key in infos_map:
            client.update(infos_map[ref_key])
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
    client_headers = [
        "Date", "Réf. Client", "Type Client", "Prénom", "Nom",
        "Date Arrivée", "Date Départ", "Durée Séjour",
        "Nombre Participants", "Nombre Adultes",
        "Enfants 2-12", "Bébés 0-2",
        "Téléphone", "Téléphone WhatsApp", "Email",
        "Période", "Restauration", "Hébergement", "Chambre",
        "Enfant", "Âge Enfant", "Forfait", "Circuit",
        "SGL", "DBL", "TWN", "TPL", "FML"
    ]
    client_header_style = {
        "font": Font(bold=True, color="FFFFFF"),
        "fill": PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid"),
        "alignment": Alignment(horizontal="center")
    }
    header_map = _ensure_headers(ws, client_headers, client_header_style)

    # Update data
    field_map = {
        "Date": ["Timestamp", "Date", "date_jour"],
        "Réf. Client": ["Ref_Client", "ref_client"],
        "Type Client": ["Type_Client", "type_client"],
        "Prénom": ["Prénom", "prenom"],
        "Nom": ["Nom", "nom"],
        "Date Arrivée": ["Date_Arrivée", "date_arrivee"],
        "Date Départ": ["Date_Départ", "date_depart"],
        "Durée Séjour": ["Durée_Séjour", "duree_sejour"],
        "Nombre Participants": ["Nombre_Participants", "nombre_participants"],
        "Nombre Adultes": ["Nombre_Adultes", "nombre_adultes"],
        "Enfants 2-12": ["Enfants_2_12_ans", "nombre_enfants_2_12"],
        "Bébés 0-2": ["Bébés_0_2_ans", "nombre_bebes_0_2"],
        "Téléphone": ["Téléphone", "telephone"],
        "Téléphone WhatsApp": ["Téléphone_WhatsApp", "telephone_whatsapp"],
        "Email": ["Email", "email"],
        "Période": ["Période", "periode"],
        "Restauration": ["Restauration", "restauration"],
        "Hébergement": ["Hébergement", "hebergement"],
        "Chambre": ["Chambre", "chambre"],
        "Enfant": ["Enfant", "enfant"],
        "Âge Enfant": ["Âge_Enfant", "age_enfant"],
        "Forfait": ["Forfait", "forfait"],
        "Circuit": ["Circuit", "circuit"],
        "SGL": ["SGL_Count", "sgl_count"],
        "DBL": ["DBL_Count", "dbl_count"],
        "TWN": ["TWN_Count", "twn_count"],
        "TPL": ["TPL_Count", "tpl_count"],
        "FML": ["FML_Count", "fml_count"]
    }
    for header, keys in field_map.items():
        col = header_map.get(header)
        if col:
            ws.cell(row=row_number, column=col, value=_first_available(client_data, keys, ""))

    wb.save(CLIENT_EXCEL_PATH)
    invalidate_client_cache()
    _save_client_infos_to_excel(client_data)
    return True


def _load_client_infos_map():
    """Load extended client infos from INFOS_CLIENTS sheet by ref client"""
    if not OPENPYXL_AVAILABLE:
        return {}
    if not os.path.exists(CLIENT_EXCEL_PATH):
        return {}

    wb = load_workbook(CLIENT_EXCEL_PATH, data_only=True)
    if CLIENT_INFOS_SHEET_NAME not in wb.sheetnames:
        return {}

    ws = wb[CLIENT_INFOS_SHEET_NAME]
    header_map = _ensure_headers(ws, [])
    infos_map = {}

    def _cell(row, header, fallback_cell=None):
        col = header_map.get(header)
        if col:
            return ws.cell(row=row, column=col).value
        if fallback_cell:
            return ws[fallback_cell].value
        return None

    for row in range(2, ws.max_row + 1):
        if ws[f'A{row}'].value is None:
            continue
        ref_client = _cell(row, "Réf. Client", f'B{row}') or ''
        if not ref_client:
            continue
        infos_map[ref_client] = {
            'numero_dossier': _cell(row, "Numéro Dossier") or '',
            'type_circuit': _cell(row, "Type Circuit") or '',
            'ville_depart': _cell(row, "Ville Départ") or '',
            'ville_arrivee': _cell(row, "Ville Arrivée") or '',
            'type_hotel_arrivee': _cell(row, "Type Hôtel Arrivée") or ''
        }

    return infos_map


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
    invalidate_client_cache()
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
    header_map = _ensure_headers(ws, [])

    hotels = []
    # Start from row 2 (skip headers)
    for row in range(2, ws.max_row + 1):
        if ws[f'A{row}'].value is None:
            continue

        def _cell(header, fallback_cell=None):
            col = header_map.get(header)
            if col:
                return ws.cell(row=row, column=col).value
            if fallback_cell:
                return ws[fallback_cell].value
            return None

        hotel = {
            'row_number': row,
            'id': _cell("ID") or f"{_cell('Ville', f'A{row}')}_{_cell('HTL', f'B{row}')}",
            'nom': _cell("HTL", f'B{row}') or '',
            'lieu': _cell("Ville", f'A{row}') or '',
            'type_hebergement': _cell("TYPE_HEBERGEMENT") or 'Hôtel',
            'categorie': _cell("CATÉGORIE", f'C{row}') or '',
            'type_client': _cell("TYPE_CLIENT", f'N{row}') or 'TO',
            'chambre_single': _parse_num(_cell("SPL", f'E{row}')),
            'chambre_double': _parse_num(_cell("DBL", f'F{row}')),
            'chambre_familiale': _parse_num(_cell("FML", f'H{row}')),
            'lit_supp': _parse_num(_cell("SUPP", f'I{row}')),
            'day_use': _parse_num(_cell("DAY_USE")) if _cell("DAY_USE") is not None else 0,
            'vignette': _parse_num(_cell("VIGNETTE")) if _cell("VIGNETTE") is not None else 0,
            'taxe_sejour': _parse_num(_cell("TAXE_SEJOUR")) if _cell("TAXE_SEJOUR") is not None else 0,
            'petit_dejeuner': _parse_num(_cell("PDJ", f'K{row}')),
            'dejeuner': _parse_num(_cell("DJ", f'L{row}')),
            'diner': _parse_num(_cell("DR", f'M{row}')),
            'description': _cell("DESCRIPTION") or f"Unité: {_cell('UNITÉ', f'D{row}') or ''}, Suite: {_cell('SUITE', f'J{row}') or ''}",
            'contact': _cell("CONTACT") or '',
            'email': _cell("EMAIL") or ''
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
    hotel_headers = [
        "Ville", "HTL", "CATÉGORIE", "UNITÉ", "SPL", "DBL", "TWINS",
        "FML", "SUPP", "SUITE", "PDJ", "DJ", "DR",
        "ID", "TYPE_HEBERGEMENT", "TYPE_CLIENT", "CONTACT", "EMAIL",
        "DESCRIPTION", "DAY_USE", "VIGNETTE", "TAXE_SEJOUR"
    ]
    hotel_header_style = {
        "font": Font(bold=True, color="FFFFFF"),
        "fill": PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid"),
        "alignment": Alignment(horizontal="center")
    }

    if not os.path.exists(HOTEL_EXCEL_PATH):
        wb = Workbook()
        wb.save(HOTEL_EXCEL_PATH)

    # Open existing file
    wb = load_workbook(HOTEL_EXCEL_PATH)

    if HOTEL_SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(HOTEL_SHEET_NAME)
        _ensure_headers(ws, hotel_headers, hotel_header_style)
        wb.save(HOTEL_EXCEL_PATH)

    # Open existing file again
    wb = load_workbook(HOTEL_EXCEL_PATH)
    ws = wb[HOTEL_SHEET_NAME]
    header_map = _ensure_headers(ws, hotel_headers, hotel_header_style)

    # Find last empty row (column A)
    last_row = ws.max_row + 1

    field_map = {
        "Ville": ["Lieu", "lieu"],
        "HTL": ["Nom", "nom"],
        "CATÉGORIE": ["Catégorie", "categorie"],
        "UNITÉ": ["Unité", "unite"],
        "SPL": ["Chambre_Single", "chambre_single"],
        "DBL": ["Chambre_Double", "chambre_double"],
        "TWINS": ["Chambre_Double", "chambre_double"],
        "FML": ["Chambre_Familiale", "chambre_familiale"],
        "SUPP": ["Lit_Supp", "lit_supp"],
        "SUITE": ["Suite", "suite"],
        "PDJ": ["Petit_Déjeuner", "petit_dejeuner"],
        "DJ": ["Déjeuner", "dejeuner"],
        "DR": ["Dîner", "diner"],
        "ID": ["ID", "id"],
        "TYPE_HEBERGEMENT": ["Type_Hébergement", "type_hebergement"],
        "TYPE_CLIENT": ["Type_Client", "type_client"],
        "CONTACT": ["Contact", "contact"],
        "EMAIL": ["Email", "email"],
        "DESCRIPTION": ["Description", "description"],
        "DAY_USE": ["Day_Use", "day_use"],
        "VIGNETTE": ["Vignette", "vignette"],
        "TAXE_SEJOUR": ["Taxe_Séjour", "taxe_sejour"]
    }
    for header, keys in field_map.items():
        col = header_map.get(header)
        if col:
            value = _first_available(hotel_data, keys, "")
            if header == "UNITÉ" and value == "":
                value = "$"
            ws.cell(row=last_row, column=col, value=value)

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
    invalidate_hotel_cache()
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
    hotel_headers = [
        "Ville", "HTL", "CATÉGORIE", "UNITÉ", "SPL", "DBL", "TWINS",
        "FML", "SUPP", "SUITE", "PDJ", "DJ", "DR",
        "ID", "TYPE_HEBERGEMENT", "TYPE_CLIENT", "CONTACT", "EMAIL",
        "DESCRIPTION", "DAY_USE", "VIGNETTE", "TAXE_SEJOUR"
    ]
    hotel_header_style = {
        "font": Font(bold=True, color="FFFFFF"),
        "fill": PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid"),
        "alignment": Alignment(horizontal="center")
    }
    header_map = _ensure_headers(ws, hotel_headers, hotel_header_style)

    field_map = {
        "Ville": ["Lieu", "lieu"],
        "HTL": ["Nom", "nom"],
        "CATÉGORIE": ["Catégorie", "categorie"],
        "UNITÉ": ["Unité", "unite"],
        "SPL": ["Chambre_Single", "chambre_single"],
        "DBL": ["Chambre_Double", "chambre_double"],
        "TWINS": ["Chambre_Double", "chambre_double"],
        "FML": ["Chambre_Familiale", "chambre_familiale"],
        "SUPP": ["Lit_Supp", "lit_supp"],
        "SUITE": ["Suite", "suite"],
        "PDJ": ["Petit_Déjeuner", "petit_dejeuner"],
        "DJ": ["Déjeuner", "dejeuner"],
        "DR": ["Dîner", "diner"],
        "ID": ["ID", "id"],
        "TYPE_HEBERGEMENT": ["Type_Hébergement", "type_hebergement"],
        "TYPE_CLIENT": ["Type_Client", "type_client"],
        "CONTACT": ["Contact", "contact"],
        "EMAIL": ["Email", "email"],
        "DESCRIPTION": ["Description", "description"],
        "DAY_USE": ["Day_Use", "day_use"],
        "VIGNETTE": ["Vignette", "vignette"],
        "TAXE_SEJOUR": ["Taxe_Séjour", "taxe_sejour"]
    }
    for header, keys in field_map.items():
        col = header_map.get(header)
        if col:
            value = _first_available(hotel_data, keys, "")
            if header == "UNITÉ" and value == "":
                value = "$"
            ws.cell(row=row_number, column=col, value=value)

    wb.save(HOTEL_EXCEL_PATH)
    invalidate_hotel_cache()
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
    invalidate_hotel_cache()
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
