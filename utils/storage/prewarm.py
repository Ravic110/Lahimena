"""Prechauffage des caches de reference au demarrage.

Les tables de reference sont mises en cache, mais le premier appel les paie
plein tarif : ~1,7 s pour le catalogue hotels, ~1,5 s pour l'appel reseau des
taux de change. Ce cout tombait sur le premier ecran de cotation ouvert, qui
mettait donc plus de 4 secondes a apparaitre.

Ces lectures n'ont aucune raison d'attendre l'ouverture de l'ecran : elles
peuvent se faire pendant que l'utilisateur saisit ses identifiants. Le
prechauffage tourne dans un fil demon, en dehors du fil de l'interface, et
n'est jamais attendu -- s'il n'a pas fini, l'ecran se comporte comme avant.

Aucune synchronisation n'est necessaire : les caches sont des dictionnaires et
le pire cas d'une course est que le travail soit fait deux fois, jamais qu'une
donnee soit corrompue.
"""

import threading
import time

from utils.logger import logger


def _chargeurs():
    """Les lectures a prechauffer, de la plus couteuse a la moins couteuse.

    Importees ici et non au chargement du module : `utils.excel_handler` est
    lourd, et le prechauffage ne doit pas rallonger le demarrage a froid.
    """
    from utils import excel_handler as eh
    from utils.validators import get_exchange_rates

    return [
        ("catalogue hotels", eh.load_all_hotels),
        ("taux de change", get_exchange_rates),
        ("tarifs transport", eh.get_transport_prestataires),
        ("distances KM_MADA", eh.get_km_mada_reperes),
        ("circuits", eh.load_circuit_catalog),
        ("frais collectifs", eh.load_collective_expenses_data),
        ("tarifs avion", eh.load_avion_source_data),
        ("visites et excursions", eh.load_visite_excursion_data),
        ("parametres", eh.load_all_parametrages),
    ]


def prechauffer(bloquant=False):
    """Remplit les caches de reference en arriere-plan.

    Args:
        bloquant (bool): attend la fin. Reserve aux tests ; l'application
            n'attend jamais, sous peine de figer sa fenetre de connexion.

    Returns:
        threading.Thread | None: le fil lance, ou None en mode bloquant.
    """
    if bloquant:
        _executer()
        return None

    fil = threading.Thread(target=_executer, name="prechauffage", daemon=True)
    fil.start()
    return fil


# Pause entre deux lectures. L'analyse d'un classeur par openpyxl est du Python
# pur : elle retient le GIL et concurrence la boucle d'evenements Tk. Sans cette
# respiration, la fenetre de connexion devient saccadee pendant le prechauffage.
PAUSE_ENTRE_LECTURES_S = 0.05


def _executer():
    for libelle, charger in _chargeurs():
        try:
            charger()
        except Exception as exc:
            # Un prechauffage est une optimisation : son echec ne doit jamais
            # empecher l'application de demarrer. L'ecran refera la lecture.
            logger.warning(f"Prechauffage de {libelle} sans effet : {exc}")
        time.sleep(PAUSE_ENTRE_LECTURES_S)
    logger.info("Caches de reference prechauffes")
