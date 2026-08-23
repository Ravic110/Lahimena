"""Cache des tables de reference, invalide par la date du classeur.

Les tables de reference -- circuits, tarifs transport, tarifs avion,
visites/excursions -- changent rarement mais sont relues sans cesse :
l'interface les interroge depuis les rappels de listes deroulantes. Chaque
lecture reparsait le classeur entier, soit ~500 ms sur les donnees reelles.
Choisir un prestataire gelait donc l'ecran une demi-seconde.

L'invalidation repose sur `st_mtime_ns` et la taille du fichier, et non sur
des appels manuels apres chaque `save_`. Une ecriture change la date, donc le
cache se perime de lui-meme -- y compris quand le classeur est modifie depuis
Excel, en dehors de l'application. Il n'y a aucune invalidation a oublier.

Un fichier absent n'est jamais mis en cache : sinon un classeur cree apres le
premier appel resterait invisible jusqu'au redemarrage.
"""

import os
from functools import wraps

# {nom_qualifie: (empreinte_fichier, valeur)}
_CACHE: dict = {}


def _empreinte(chemin):
    """Signature du fichier, ou None s'il n'existe pas.

    `st_mtime_ns` plutot que `st_mtime` : a la seconde pres, deux ecritures
    rapprochees passeraient pour une seule et le cache servirait des donnees
    perimees. La taille sert de garde-fou supplementaire sur les systemes de
    fichiers a faible resolution temporelle.
    """
    try:
        infos = os.stat(chemin)
    except OSError:
        return None
    return (chemin, infos.st_mtime_ns, infos.st_size)


def _protege(valeur, copier):
    """Copie de surface d'une liste retenue, pour qu'un appelant ne la modifie pas."""
    if copier and isinstance(valeur, list):
        return list(valeur)
    return valeur


def mtime_cached(chemin_fn, copier=True):
    """Retient le resultat tant que le classeur designe n'a pas change.

    Args:
        chemin_fn (callable): renvoie le chemin du classeur. Appele a chaque
            invocation, et non fige a la decoration : les tests redirigent les
            classeurs vers des fichiers temporaires.
        copier (bool): renvoie une copie de surface des listes plutot que
            l'objet retenu. Sans cela, un appelant qui trierait ou completerait
            la liste recue corromprait le cache pour tous les suivants. Le cout
            est de l'ordre de la microseconde face aux ~500 ms economisees.
    """

    def decorate(func):
        cle = f"{func.__module__}.{func.__qualname__}"

        @wraps(func)
        def wrapper(*args, **kwargs):
            # Les appels parametres ne sont pas mis en cache : ce decorateur
            # vise les chargeurs de table entiere, sans argument.
            if args or kwargs:
                return func(*args, **kwargs)

            empreinte = _empreinte(chemin_fn())
            if empreinte is not None:
                connue = _CACHE.get(cle)
                if connue is not None and connue[0] == empreinte:
                    return _protege(connue[1], copier)

            valeur = func()

            if empreinte is not None:
                _CACHE[cle] = (empreinte, valeur)
            else:
                _CACHE.pop(cle, None)

            return _protege(valeur, copier)

        return wrapper

    return decorate


def vider_les_caches_de_source():
    """Oublie tout ce qui est retenu. Utile aux tests et apres une migration."""
    _CACHE.clear()
