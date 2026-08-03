"""Ouverture, sauvegarde et fermeture des classeurs Excel.

Avant ce module, 78 fonctions de utils.excel_handler repetaient le meme
prologue et le meme epilogue : verification d'openpyxl, verification du
fichier, verification de la feuille, `wb = None`, `try`, `wb.save(...)`,
`except PermissionError`, `except Exception` + log, puis un `finally` qui
fermait le classeur en avalant l'erreur. `open_workbook` porte tout cela une
fois.

Le module expose aussi une hierarchie d'exceptions. Les fonctions publiques
historiques, elles, doivent continuer a renvoyer leurs sentinelles (-1, -2,
False, []) parce que l'interface les teste directement : `sentinel_on_error`
fait cette traduction, au seul endroit ou elle a lieu.
"""

import os
import shutil
from contextlib import contextmanager
from datetime import datetime
from functools import wraps

try:
    from openpyxl import Workbook, load_workbook

    OPENPYXL_AVAILABLE = True
except ImportError:  # pragma: no cover - depend de l'environnement
    OPENPYXL_AVAILABLE = False

from utils.logger import logger


class StorageError(Exception):
    """Base des erreurs d'acces au stockage Excel."""


class StorageUnavailable(StorageError):
    """openpyxl n'est pas installe."""


class WorkbookMissing(StorageError):
    """Le fichier classeur n'existe pas."""


class SheetMissing(StorageError):
    """La feuille demandee n'existe pas dans le classeur."""


class WorkbookLocked(StorageError):
    """Le fichier est verrouille, typiquement ouvert dans Excel."""


def create_backup(filepath):
    """Copie horodatee du classeur dans un sous-dossier `backups`.

    Args:
        filepath (str): chemin du classeur a sauvegarder.

    Returns:
        str | None: chemin de la copie, ou None en cas d'echec.
    """
    if not os.path.exists(filepath):
        logger.warning(f"File not found for backup: {filepath}")
        return None

    try:
        backup_dir = os.path.join(os.path.dirname(filepath), "backups")
        os.makedirs(backup_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.basename(filepath)
        backup_path = os.path.join(backup_dir, f"{filename}.{timestamp}.bak")

        shutil.copy2(filepath, backup_path)
        logger.info(f"Backup created: {backup_path}")

        return backup_path
    except Exception as e:
        logger.error(f"Failed to create backup for {filepath}: {e}", exc_info=True)
        return None


@contextmanager
def open_workbook(
    path,
    sheet=None,
    *,
    create=False,
    write=False,
    backup=False,
    data_only=False,
):
    """Ouvre un classeur et cede la feuille demandee.

    Args:
        path (str): chemin du classeur.
        sheet (str | None): feuille a ceder. Si None, cede le classeur entier.
        create (bool): cree le fichier et/ou la feuille s'ils manquent.
        write (bool): sauvegarde le classeur a la sortie du bloc.
        backup (bool): copie de securite avant sauvegarde. Sans effet si le
            fichier n'existe pas encore.
        data_only (bool): lit les valeurs calculees plutot que les formules.

    Yields:
        La feuille demandee, ou le classeur si `sheet` vaut None.

    Raises:
        StorageUnavailable: openpyxl absent.
        WorkbookMissing: fichier absent et `create` faux.
        SheetMissing: feuille absente et `create` faux.
        WorkbookLocked: fichier verrouille par une autre application.
    """
    if not OPENPYXL_AVAILABLE:
        raise StorageUnavailable(f"openpyxl indisponible, acces impossible a {path}")

    exists = os.path.exists(path)
    if not exists and not create:
        raise WorkbookMissing(path)

    wb = None
    try:
        if exists:
            try:
                wb = load_workbook(path, data_only=data_only)
            except PermissionError as exc:
                raise WorkbookLocked(path) from exc
        else:
            wb = Workbook()
            if sheet is not None:
                wb.active.title = sheet

        if sheet is None:
            target = wb
        elif sheet in wb.sheetnames:
            target = wb[sheet]
        elif create:
            target = wb.create_sheet(sheet)
        else:
            raise SheetMissing(f"{sheet} absente de {path}")

        yield target

        if write:
            if backup and exists:
                create_backup(path)
            try:
                wb.save(path)
            except PermissionError as exc:
                raise WorkbookLocked(path) from exc
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def sentinel_on_error(missing, locked=None, *, label=""):
    """Traduit les exceptions du stockage vers les sentinelles historiques.

    L'interface teste directement les valeurs de retour (-1, -2, False, []),
    donc la signature des fonctions publiques ne peut pas changer. Ce decorateur
    concentre la conversion au lieu de la repeter dans chaque fonction.

    Reproduit fidelement le comportement d'origine : une cause attendue
    (openpyxl absent, fichier ou feuille manquants) est silencieuse, alors
    qu'une erreur inattendue est journalisee avec sa trace.

    Args:
        missing: valeur renvoyee pour une cause attendue ou une erreur.
        locked: valeur renvoyee si le fichier est verrouille. Par defaut,
            identique a `missing`.
        label (str): libelle utilise dans le message de log.
    """
    locked_value = missing if locked is None else locked

    def decorate(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except WorkbookLocked:
                return locked_value
            except (StorageUnavailable, WorkbookMissing, SheetMissing):
                return missing
            except Exception as exc:
                logger.error(
                    f"{label or func.__name__} a echoue : {exc}", exc_info=True
                )
                return missing

        return wrapper

    return decorate
