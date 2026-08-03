"""Manipulation de feuilles Excel : en-tetes, colonnes, conversion de valeurs.

Fonctions pures extraites de utils.excel_handler. Elles ne connaissent ni les
chemins de fichiers, ni les entites metier : elles operent sur une feuille
openpyxl deja ouverte, ou sur des valeurs brutes.

Cette absence d'etat les rend directement testables, ce qui n'etait pas le cas
tant qu'elles vivaient au milieu de 7000 lignes d'acces disque.
"""

import re
from datetime import datetime, time, timedelta


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
        s = s.replace(",", "").replace(" ", "")
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        if not m:
            return 0
        num_str = m.group(0)
        if "." in num_str:
            return float(num_str)
        return int(num_str)
    except Exception:
        return 0


def _parse_duration_hours(val):
    """Parse duration values into hours.

    Supports numeric hours, Excel time/timedelta values, and strings like
    "3h30", "03:30", "210 min", or "2.5".
    """
    if val is None or val == "":
        return 0.0

    if isinstance(val, timedelta):
        return max(0.0, float(val.total_seconds()) / 3600)

    if isinstance(val, time):
        return val.hour + (val.minute / 60.0) + (val.second / 3600.0)

    if isinstance(val, datetime):
        return val.hour + (val.minute / 60.0) + (val.second / 3600.0)

    if isinstance(val, (int, float)):
        hours = float(val)
        return hours if hours > 0 else 0.0

    raw = str(val).strip().lower().replace(",", ".")
    if not raw:
        return 0.0

    compact = re.sub(r"\s+", "", raw)

    match_h = re.fullmatch(r"(\d+(?:\.\d+)?)h(?:(\d+(?:\.\d+)?)(?:mn|min|m)?)?", compact)
    if match_h:
        hours = float(match_h.group(1))
        minutes = float(match_h.group(2)) if match_h.group(2) else 0.0
        return max(0.0, hours + (minutes / 60.0))

    match_clock = re.fullmatch(r"(\d{1,3}):(\d{1,2})(?::(\d{1,2}))?", compact)
    if match_clock:
        hh = int(match_clock.group(1))
        mm = int(match_clock.group(2))
        ss = int(match_clock.group(3) or 0)
        return max(0.0, hh + (mm / 60.0) + (ss / 3600.0))

    match_min = re.fullmatch(r"(\d+(?:\.\d+)?)(?:mn|min|m)", compact)
    if match_min:
        minutes = float(match_min.group(1))
        return max(0.0, minutes / 60.0)

    try:
        parsed = float(compact)
        return parsed if parsed > 0 else 0.0
    except Exception:
        return 0.0


def _normalize_header_key(value):
    """Normalize header labels for resilient matching."""
    if value is None:
        return ""
    normalized = str(value).strip().lower()
    replacements = str.maketrans(
        {
            "é": "e",
            "è": "e",
            "ê": "e",
            "ë": "e",
            "à": "a",
            "â": "a",
            "ä": "a",
            "î": "i",
            "ï": "i",
            "ô": "o",
            "ö": "o",
            "ù": "u",
            "û": "u",
            "ü": "u",
            "ç": "c",
            "'": " ",
            "_": " ",
            "-": " ",
        }
    )
    normalized = normalized.translate(replacements)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def _find_header_column(header_map, *aliases):
    """Find a header column using exact or normalized aliases."""
    for alias in aliases:
        if alias in header_map:
            return header_map[alias]

    normalized_map = {_normalize_header_key(k): v for k, v in header_map.items()}
    for alias in aliases:
        col = normalized_map.get(_normalize_header_key(alias))
        if col:
            return col
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


def _get_header_map(ws, header_row=1):
    header_map = {}
    max_col = ws.max_column if ws.max_column and ws.max_column > 0 else 0
    for col in range(1, max_col + 1):
        value = ws.cell(row=header_row, column=col).value
        if value is not None and str(value).strip() != "":
            header_map[str(value).strip()] = col
    return header_map


def _iter_grouped_columns(ws, group_row=1, header_row=2):
    columns = []
    last_group = ""
    max_col = ws.max_column if ws.max_column and ws.max_column > 0 else 0
    for col in range(1, max_col + 1):
        group_val = ws.cell(row=group_row, column=col).value
        if group_val is not None and str(group_val).strip() != "":
            last_group = str(group_val).strip()
        header_val = ws.cell(row=header_row, column=col).value
        if header_val is None or str(header_val).strip() == "":
            continue
        columns.append((last_group, str(header_val).strip(), col))
    return columns


def _first_available(data, keys, default=""):
    for key in keys:
        if key in data and data.get(key) not in (None, ""):
            return data.get(key)
    return default
