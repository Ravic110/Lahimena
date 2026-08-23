"""Traduction des codes de retour du stockage en messages utilisateur.

La couche de donnees renvoie des sentinelles heteregenes, testees directement
par l'interface :

    save_*    n° de ligne | -1 erreur | -2 fichier verrouille
    update_*  0           | -1        | -2
    delete_*  True        | False     | False

Ces comparaisons sont recopiees sur une cinquantaine de sites dans `gui/`, avec
des libelles qui divergent. Pire, trois sauvegardes implicites -- celles qui
persistent un document au moment de l'afficher -- ignoraient purement leur
valeur de retour : l'echec passait inapercu et l'utilisateur travaillait sur un
document qu'il croyait enregistre.

Ce module isole la traduction. Elle est pure et testable sans tkinter ;
`signaler_echec` n'est que la fine couche d'affichage par-dessus.
"""

from tkinter import messagebox

# Sentinelles renvoyees par utils.excel_handler.
ECHEC = -1
FICHIER_VERROUILLE = -2


def message_d_echec(resultat, fichier="data.xlsx"):
    """Message a afficher pour un code de retour, ou None si l'operation a reussi.

    Args:
        resultat: valeur renvoyee par une fonction save_/update_/delete_.
        fichier (str): classeur concerne, cite dans le message de verrouillage.

    Returns:
        str | None: le message, ou None quand il n'y a rien a signaler.
    """
    if resultat is False:
        return "L'enregistrement a echoue."
    if resultat is True or resultat is None:
        return None
    if not isinstance(resultat, int):
        return None
    if resultat == FICHIER_VERROUILLE:
        return (
            f"{fichier} est ouvert dans une autre application.\n"
            "Fermez-le puis reessayez."
        )
    if resultat == ECHEC:
        return "L'enregistrement a echoue. Voir les journaux pour le detail."
    return None


def signaler_echec(titre, resultat, fichier="data.xlsx"):
    """Affiche une alerte si `resultat` denote un echec. Silencieux si succes.

    Returns:
        bool: True si l'operation avait reussi, False si un echec a ete signale.
    """
    message = message_d_echec(resultat, fichier)
    if message is None:
        return True
    messagebox.showerror(titre, message)
    return False
