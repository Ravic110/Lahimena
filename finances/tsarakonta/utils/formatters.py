"""
Fonctions utilitaires de formatage
"""


def format_montant(montant):
    """
    Formate un montant numérique pour l'affichage

    Args:
        montant: valeur numérique

    Returns:
        str: montant formaté avec séparateurs
    """
    return f"{montant:,.2f}".replace(",", " ")


def parse_montant(value_str):
    """
    Convertit une chaîne en montant numérique

    Args:
        value_str: chaîne à convertir (format: 1234,56)

    Returns:
        float: montant parsé

    Raises:
        ValueError: si le format n'est pas valide
    """
    val = value_str.strip().replace(" ", "").replace(",", ".")
    if not val:
        return 0.0
    return float(val)


def extraire_numero_compte(compte_display):
    """
    Extrait le numéro de compte de l'affichage (ex: "401 - Compte client")

    Args:
        compte_display: chaîne au format "numero - libelle"

    Returns:
        str: numéro du compte
    """
    if " - " in compte_display:
        return compte_display.split(" - ")[0]
    return compte_display
