"""Tables de reference du classeur data-hotel.xlsx.

Six entites y cohabitent : circuits, frais collectifs, transport, avion,
visites/excursions et distances (KM_MADA). Elles partagent le meme classeur et
la meme forme d'API (`get_*_db_headers`, `load_*_db_rows`, `save_*_db_row`,
`update_*_db_row`, `delete_*_db_row`), mais pas leur mecanique interne.

Aucune abstraction unificatrice n'est tentee ici, et c'est delibere : la mesure
a montre sept axes de divergence entre ces six entites (resolution du nom de
feuille, decouverte des en-tetes, adressage des colonnes, ligne de depart des
donnees, point d'insertion, coercition, invalidation de cache), chacune
retenant une combinaison unique. Un socle parametre sur sept axes aurait coute
plus a lire que la repetition qu'il supprime.

Ce qui est mutualise, ce sont les primitives reellement identiques :
`open_workbook` pour le cycle ouverture/sauvegarde/fermeture, et les fonctions
de `sheet` pour la lecture des lignes et le placement des ecritures. Chaque
entite garde sa specificite visible, en quelques lignes.

Ecart assume avec le code d'origine : `km_mada` n'attrapait que
`(OSError, ValueError, KeyError)` la ou les cinq autres attrapaient `Exception`.
Une `TypeError` y remontait donc jusqu'a l'interface. L'uniformisation via
`sentinel_on_error` supprime cette asymetrie, tres probablement involontaire.
"""

import config
from utils.storage.sheet import (
    _ensure_headers,
    _get_header_map,
    _normalize_header_key,
    _parse_num,
    index_header_map,
    next_empty_row,
    read_data_rows,
    resolve_sheet_name,
    write_row,
)
from utils.storage.workbook import (
    SheetMissing,
    open_workbook,
    sentinel_on_error,
)

CIRCUIT_SHEET = "Circuits"

CIRCUIT_DEFAULT_HEADERS = [
    "ID circuit",
    "Nom du circuit",
    "itinéraire",
    "Villes parcourues",
    "Activité",
    "Durée",
    "condition physique",
    "Type de voiture",
    "Hôtels défaut par ville",
    "Prestations incluses",
    "Transports associés",
]

COLLECTIVE_HEADERS = [
    "FORFAIT",
    "PRESTATAIRES",
    "DESIGNATION",
    "MONTANT",
    "ID circuit",
]

AVION_DEFAULT_HEADERS = [
    "Ville de départ",
    "Ville d'arrivée",
    "Tarif Adultes",
    "Tarif Enfants",
]

VISITE_DEFAULT_HEADERS = ["PRESTATIONS", "DESIGNATION", "Tarif par pax"]


def _path():
    """Chemin du classeur, resolu a l'appel pour rester substituable en test."""
    return config.HOTEL_EXCEL_PATH


def _read_first_row_labels(ws):
    """Libelles non vides de la premiere ligne, dans l'ordre des colonnes."""
    labels = []
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value
        if value is None:
            continue
        label = str(value).strip()
        if label:
            labels.append(label)
    return labels


# --------------------------------------------------------------------------
# Circuits — en-tetes garantis par _ensure_headers, insertion par balayage
# --------------------------------------------------------------------------


@sentinel_on_error([], label="Lecture des en-tetes circuits")
def get_circuit_db_headers():
    """Libelles de colonnes de la feuille Circuits.

    Ecrit les en-tetes par defaut s'ils manquent : la lecture a donc un effet
    de bord sur le fichier, comportement conserve de l'origine.
    """
    with open_workbook(_path(), CIRCUIT_SHEET, write=True) as ws:
        return list(_ensure_headers(ws, CIRCUIT_DEFAULT_HEADERS).keys())


@sentinel_on_error([], label="Lecture des circuits")
def load_circuit_db_rows():
    """Lignes brutes de la feuille Circuits."""
    with open_workbook(_path(), CIRCUIT_SHEET) as ws:
        header_map = _get_header_map(ws, 1)
        return read_data_rows(ws, header_map, 2) if header_map else []


@sentinel_on_error(-1, -2, label="Ecriture d'un circuit")
def save_circuit_db_row(row_data):
    """Ajoute une ligne dans Circuits et renvoie son numero."""
    with open_workbook(_path(), CIRCUIT_SHEET, create=True, write=True) as ws:
        header_map = _ensure_headers(ws, CIRCUIT_DEFAULT_HEADERS)
        if row_data:
            header_map = _ensure_headers(ws, list(row_data.keys()))
        if not header_map:
            return -1

        row_number = next_empty_row(ws, header_map)
        write_row(ws, header_map, row_number, row_data)
        return row_number


@sentinel_on_error(-1, -2, label="Mise a jour d'un circuit")
def update_circuit_db_row(row_number, row_data):
    """Remplace une ligne de Circuits. Renvoie 0 en succes."""
    with open_workbook(_path(), CIRCUIT_SHEET, write=True) as ws:
        header_map = _get_header_map(ws, 1)
        if not header_map:
            return -1
        write_row(ws, header_map, row_number, row_data)
        return 0


@sentinel_on_error(False, label="Suppression d'un circuit")
def delete_circuit_db_row(row_number):
    """Supprime une ligne de Circuits."""
    with open_workbook(_path(), CIRCUIT_SHEET, write=True) as ws:
        ws.delete_rows(row_number)
        return True


# --------------------------------------------------------------------------
# Frais collectifs — colonnes positionnelles figees, vocabulaire metier
# --------------------------------------------------------------------------


@sentinel_on_error([], label="Lecture des frais collectifs")
def load_collective_expense_db_rows():
    """Lignes de la feuille Frais collectifs, en cles metier.

    Seule entite dont les cles renvoyees different des libelles de colonnes, et
    seule a coercer ses valeurs (`montant` en nombre, le reste en texte nettoye).
    """
    with open_workbook(_path(), config.FRAIS_COLLECTIFS_SHEET_NAME) as ws:
        rows = []
        for row_idx in range(2, ws.max_row + 1):
            cells = [ws.cell(row=row_idx, column=col).value for col in range(1, 6)]
            if all(value in (None, "") for value in cells):
                continue
            forfait, prestataire, designation, montant, id_circuit = cells
            rows.append(
                {
                    "row_number": row_idx,
                    "forfait": str(forfait or "").strip(),
                    "prestataire": str(prestataire or "").strip(),
                    "designation": str(designation or "").strip(),
                    "montant": _parse_num(montant),
                    "id_circuit": (
                        "" if id_circuit is None else str(id_circuit).strip()
                    ),
                }
            )
        return rows


@sentinel_on_error(-1, -2, label="Ecriture d'un frais collectif")
def save_collective_expense_db_row(row_data):
    """Ajoute une ligne dans Frais collectifs et renvoie son numero.

    Insere a `max_row + 1`, sans chercher a combler les lignes vides.
    """
    path = _path()
    with open_workbook(
        path, config.FRAIS_COLLECTIFS_SHEET_NAME, create=True, write=True
    ) as ws:
        if ws.max_row == 1 and ws.cell(row=1, column=1).value in (None, ""):
            for col_idx, header in enumerate(COLLECTIVE_HEADERS, start=1):
                ws.cell(row=1, column=col_idx, value=header)

        row_number = ws.max_row + 1
        ws.cell(row=row_number, column=1, value=row_data.get("forfait", ""))
        ws.cell(row=row_number, column=2, value=row_data.get("prestataire", ""))
        ws.cell(row=row_number, column=3, value=row_data.get("designation", ""))
        ws.cell(row=row_number, column=4, value=_parse_num(row_data.get("montant", 0)))
        ws.cell(row=row_number, column=5, value=row_data.get("id_circuit", ""))
        return row_number


@sentinel_on_error(-1, -2, label="Mise a jour d'un frais collectif")
def update_collective_expense_db_row(row_number, row_data):
    """Remplace une ligne de Frais collectifs. Renvoie 0 en succes.

    `id_circuit` n'est ecrit que s'il est fourni, contrairement aux autres
    colonnes qui sont vidées quand elles manquent.
    """
    with open_workbook(
        _path(), config.FRAIS_COLLECTIFS_SHEET_NAME, write=True
    ) as ws:
        ws.cell(row=row_number, column=1, value=row_data.get("forfait", ""))
        ws.cell(row=row_number, column=2, value=row_data.get("prestataire", ""))
        ws.cell(row=row_number, column=3, value=row_data.get("designation", ""))
        ws.cell(row=row_number, column=4, value=_parse_num(row_data.get("montant", 0)))
        if "id_circuit" in row_data:
            ws.cell(row=row_number, column=5, value=row_data.get("id_circuit", ""))
        return 0


@sentinel_on_error(False, label="Suppression d'un frais collectif")
def delete_collective_expense_db_row(row_number):
    """Supprime une ligne de Frais collectifs."""
    with open_workbook(
        _path(), config.FRAIS_COLLECTIFS_SHEET_NAME, write=True
    ) as ws:
        ws.delete_rows(row_number)
        return True
