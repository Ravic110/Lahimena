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
from utils.storage.km_cache import _invalidate_km_mada_cache
from utils.storage.sheet import (
    _ensure_headers,
    _find_header_column,
    _get_header_map,
    _parse_num,
    index_header_map,
    next_empty_row,
    read_data_rows,
    resolve_sheet_name,
    write_row,
)
from utils.storage.workbook import open_workbook, sentinel_on_error

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
    with open_workbook(_path(), config.FRAIS_COLLECTIFS_SHEET_NAME, write=True) as ws:
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
    with open_workbook(_path(), config.FRAIS_COLLECTIFS_SHEET_NAME, write=True) as ws:
        ws.delete_rows(row_number)
        return True


# --------------------------------------------------------------------------
# Transport et KM_MADA — en-tetes resolus, ligne de donnees variable
# --------------------------------------------------------------------------
#
# Ces deux feuilles sont editees a la main avec des colonnes groupees : une
# ligne de regroupement peut surmonter la ligne d'en-tetes. Les resolveurs
# essaient la ligne 1, puis la ligne 2, en reperant une colonne sentinelle, et
# decalent le debut des donnees en consequence.


def _resolve_transport_source_header_map(ws):
    """En-tetes de la feuille TRANSPORT et premiere ligne de donnees."""
    header_map = _get_header_map(ws, 1)
    prestataire_col = _find_header_column(header_map, "Prestataire")
    type_col = _find_header_column(header_map, "Type de voiture", "Type voiture")
    if prestataire_col and type_col:
        return header_map, 2

    header_map = _get_header_map(ws, 2)
    prestataire_col = _find_header_column(header_map, "Prestataire")
    type_col = _find_header_column(header_map, "Type de voiture", "Type voiture")
    if prestataire_col and type_col:
        return header_map, 3

    return {}, 2


def _resolve_km_mada_header_map(ws):
    """En-tetes de la feuille KM_MADA et premiere ligne de donnees."""
    for header_row, data_row in ((1, 2), (2, 3)):
        header_map = _get_header_map(ws, header_row)
        if _find_header_column(
            header_map, "REPERES", "Reperes", "Repères", "REPERE", "Repere"
        ):
            return header_map, data_row
    return {}, 2


def _write_header_row(ws, headers):
    """Ecrit une ligne d'en-tetes en partant de la colonne 1."""
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=header)


def _save_resolved_row(sheet, resolver, row_data, invalidate=False):
    """Ajoute une ligne dans une feuille a en-tetes resolus.

    Si aucun en-tete n'est reconnu, les cles de `row_data` en tiennent lieu et
    sont ecrites en ligne 1.
    """
    with open_workbook(_path(), sheet, create=True, write=True) as ws:
        header_map, data_start_row = resolver(ws)
        if not header_map:
            headers = list(row_data.keys()) if row_data else []
            if not headers:
                return -1
            _write_header_row(ws, headers)
            header_map = _get_header_map(ws, 1)
            data_start_row = 2

        row_number = next_empty_row(ws, header_map, data_start_row)
        write_row(ws, header_map, row_number, row_data)
        if invalidate:
            _invalidate_km_mada_cache()
        return row_number


def _update_resolved_row(sheet, resolver, row_number, row_data, invalidate=False):
    """Remplace une ligne dans une feuille a en-tetes resolus."""
    with open_workbook(_path(), sheet, write=True) as ws:
        header_map, _ = resolver(ws)
        if not header_map:
            return -1
        write_row(ws, header_map, row_number, row_data)
        if invalidate:
            _invalidate_km_mada_cache()
        return 0


def _load_resolved_rows(sheet, resolver):
    """Lit les lignes d'une feuille a en-tetes resolus."""
    with open_workbook(_path(), sheet) as ws:
        header_map, data_start_row = resolver(ws)
        return read_data_rows(ws, header_map, data_start_row) if header_map else []


def _resolved_headers(sheet, resolver):
    """Libelles de colonnes d'une feuille a en-tetes resolus."""
    with open_workbook(_path(), sheet) as ws:
        header_map, _ = resolver(ws)
        return list(header_map.keys()) if header_map else []


def _delete_row(sheet, row_number, invalidate=False):
    """Supprime une ligne d'une feuille du classeur."""
    with open_workbook(_path(), sheet, write=True) as ws:
        ws.delete_rows(row_number)
        if invalidate:
            _invalidate_km_mada_cache()
        return True


@sentinel_on_error([], label="Lecture des en-tetes transport")
def get_transport_db_headers():
    """Libelles de colonnes de la feuille TRANSPORT."""
    return _resolved_headers(
        config.TRANSPORT_SOURCE_SHEET_NAME, _resolve_transport_source_header_map
    )


@sentinel_on_error([], label="Lecture du transport")
def load_transport_db_rows():
    """Lignes brutes de la feuille TRANSPORT."""
    return _load_resolved_rows(
        config.TRANSPORT_SOURCE_SHEET_NAME, _resolve_transport_source_header_map
    )


@sentinel_on_error(-1, -2, label="Ecriture d'un transport")
def save_transport_db_row(row_data):
    """Ajoute une ligne dans TRANSPORT et renvoie son numero.

    Invalide le cache des distances : les tarifs de transport alimentent les
    calculs kilometriques. KM_MADA, lui, ne le fait pas sur ecriture.
    """
    return _save_resolved_row(
        config.TRANSPORT_SOURCE_SHEET_NAME,
        _resolve_transport_source_header_map,
        row_data,
        invalidate=True,
    )


@sentinel_on_error(-1, -2, label="Mise a jour d'un transport")
def update_transport_db_row(row_number, row_data):
    """Remplace une ligne de TRANSPORT. Renvoie 0 en succes."""
    return _update_resolved_row(
        config.TRANSPORT_SOURCE_SHEET_NAME,
        _resolve_transport_source_header_map,
        row_number,
        row_data,
        invalidate=True,
    )


@sentinel_on_error(False, label="Suppression d'un transport")
def delete_transport_db_row(row_number):
    """Supprime une ligne de TRANSPORT, sans toucher au cache."""
    return _delete_row(config.TRANSPORT_SOURCE_SHEET_NAME, row_number)


@sentinel_on_error([], label="Lecture des en-tetes KM_MADA")
def get_km_mada_db_headers():
    """Libelles de colonnes de la feuille KM_MADA."""
    return _resolved_headers(config.KM_MADA_SHEET_NAME, _resolve_km_mada_header_map)


@sentinel_on_error([], label="Lecture des distances KM_MADA")
def load_km_mada_db_rows():
    """Lignes brutes de la feuille KM_MADA."""
    return _load_resolved_rows(config.KM_MADA_SHEET_NAME, _resolve_km_mada_header_map)


@sentinel_on_error(-1, -2, label="Ecriture d'une distance KM_MADA")
def save_km_mada_db_row(row_data):
    """Ajoute une ligne dans KM_MADA et renvoie son numero."""
    return _save_resolved_row(
        config.KM_MADA_SHEET_NAME, _resolve_km_mada_header_map, row_data
    )


@sentinel_on_error(-1, -2, label="Mise a jour d'une distance KM_MADA")
def update_km_mada_db_row(row_number, row_data):
    """Remplace une ligne de KM_MADA. Renvoie 0 en succes."""
    return _update_resolved_row(
        config.KM_MADA_SHEET_NAME, _resolve_km_mada_header_map, row_number, row_data
    )


@sentinel_on_error(False, label="Suppression d'une distance KM_MADA")
def delete_km_mada_db_row(row_number):
    """Supprime une ligne de KM_MADA et vide le cache des distances."""
    return _delete_row(config.KM_MADA_SHEET_NAME, row_number, invalidate=True)


# --------------------------------------------------------------------------
# Avion et visites/excursions — colonnes adressees par position
# --------------------------------------------------------------------------
#
# Asymetrie conservee de l'origine : la lecture retrouve la feuille meme si sa
# casse ou ses accents different, alors que l'ecriture exige le nom exact.


def _labels_of(sheet):
    """Libelles de la premiere ligne, feuille retrouvee de facon souple."""
    with open_workbook(_path()) as wb:
        actual = resolve_sheet_name(wb, sheet)
        return _read_first_row_labels(wb[actual]) if actual else []


def _load_positional_rows(sheet, headers):
    """Lit une feuille dont les colonnes sont adressees par position."""
    if not headers:
        return []
    with open_workbook(_path()) as wb:
        actual = resolve_sheet_name(wb, sheet)
        if not actual:
            return []
        return read_data_rows(wb[actual], index_header_map(headers), 2)


def _save_positional_row(sheet, row_data, default_headers):
    """Ajoute une ligne a `max_row + 1`, sans combler les trous."""
    with open_workbook(_path(), sheet, create=True, write=True) as ws:
        headers = _read_first_row_labels(ws)
        if not headers:
            headers = [key for key in row_data.keys() if key != "row_number"]
            if not headers:
                headers = list(default_headers)
            _write_header_row(ws, headers)

        row_number = ws.max_row + 1
        write_row(ws, index_header_map(headers), row_number, row_data)
        return row_number


def _update_positional_row(sheet, headers, row_number, row_data):
    """Remplace une ligne adressee par position."""
    with open_workbook(_path(), sheet, write=True) as ws:
        write_row(ws, index_header_map(headers), row_number, row_data)
        return 0


@sentinel_on_error([], label="Lecture des en-tetes avion")
def get_avion_db_headers():
    """Libelles de colonnes de la feuille avion."""
    return _labels_of(config.AVION_SOURCE_SHEET_NAME)


@sentinel_on_error([], label="Lecture des vols")
def load_avion_db_rows():
    """Lignes brutes de la feuille avion."""
    return _load_positional_rows(config.AVION_SOURCE_SHEET_NAME, get_avion_db_headers())


@sentinel_on_error(-1, -2, label="Ecriture d'un vol")
def save_avion_db_row(row_data):
    """Ajoute une ligne dans la feuille avion et renvoie son numero."""
    return _save_positional_row(
        config.AVION_SOURCE_SHEET_NAME, row_data, AVION_DEFAULT_HEADERS
    )


@sentinel_on_error(-1, -2, label="Mise a jour d'un vol")
def update_avion_db_row(row_number, row_data):
    """Remplace une ligne de la feuille avion. Renvoie 0 en succes."""
    return _update_positional_row(
        config.AVION_SOURCE_SHEET_NAME, get_avion_db_headers(), row_number, row_data
    )


@sentinel_on_error(False, label="Suppression d'un vol")
def delete_avion_db_row(row_number):
    """Supprime une ligne de la feuille avion."""
    return _delete_row(config.AVION_SOURCE_SHEET_NAME, row_number)


@sentinel_on_error([], label="Lecture des en-tetes visites")
def get_visite_excursion_db_headers():
    """Libelles de colonnes de la feuille Visite_excursion."""
    return _labels_of(config.VISITE_EXCURSION_SOURCE_SHEET_NAME)


@sentinel_on_error([], label="Lecture des visites")
def load_visite_excursion_db_rows():
    """Lignes brutes de la feuille Visite_excursion."""
    return _load_positional_rows(
        config.VISITE_EXCURSION_SOURCE_SHEET_NAME, get_visite_excursion_db_headers()
    )


@sentinel_on_error(-1, -2, label="Ecriture d'une visite")
def save_visite_excursion_db_row(row_data):
    """Ajoute une ligne dans Visite_excursion et renvoie son numero."""
    return _save_positional_row(
        config.VISITE_EXCURSION_SOURCE_SHEET_NAME, row_data, VISITE_DEFAULT_HEADERS
    )


@sentinel_on_error(-1, -2, label="Mise a jour d'une visite")
def update_visite_excursion_db_row(row_number, row_data):
    """Remplace une ligne de Visite_excursion. Renvoie 0 en succes."""
    return _update_positional_row(
        config.VISITE_EXCURSION_SOURCE_SHEET_NAME,
        get_visite_excursion_db_headers(),
        row_number,
        row_data,
    )


@sentinel_on_error(False, label="Suppression d'une visite")
def delete_visite_excursion_db_row(row_number):
    """Supprime une ligne de Visite_excursion."""
    return _delete_row(config.VISITE_EXCURSION_SOURCE_SHEET_NAME, row_number)
