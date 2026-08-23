"""Helpers partages par les ecrans de cotation client.

Les cinq modules `client_*_cotation.py` -- hotel, restauration, transport,
billets d'avion, charges collectives -- recopiaient chacun les memes fonctions
de conversion et de formatage, ainsi que la meme palette de survol. Cinq
exemplaires signifient cinq endroits ou corriger un defaut, et aucun test :
`gui/` est a 0 % de couverture.

Ce module ne rassemble que ce qui etait *reellement* identique. Trois autres
familles se ressemblent sans etre interchangeables et restent donc chez elles :

* `_make_row`      -- cinq versions, une par entite (colonnes differentes) ;
* `_compute_prix_unitaire` -- deux versions (chambres vs repas) ;
* `_parse_cities`  -- deux versions : celle du transport sait retirer les
  prefixes de jour (`J1 - Antananarivo`) et decoupe sur d'autres separateurs
  que celle de l'hotellerie. Les unifier changerait le decoupage des
  itineraires, donc les segments factures.
"""

import re
import unicodedata

# Palette de survol des boutons, identique dans les cinq ecrans.
HOVER_GREEN = "#0A6870"
HOVER_BLUE = "#0B6080"
HOVER_RED = "#A82020"
HOVER_GREY = "#9EA7AA"
HOVER_ORANGE = "#C8860A"


def to_float(s, default: float = 0.0) -> float:
    """Nombre decimal depuis une saisie utilisateur.

    Accepte la virgule decimale francaise. Une saisie vide ou illisible
    retombe sur `default`.
    """
    try:
        return float(str(s).replace(",", ".").strip() or default)
    except (ValueError, TypeError):
        return default


def to_int(s, default: int = 0) -> int:
    """Entier positif depuis une saisie utilisateur.

    Tronque (3.9 donne 3) et ramene les negatifs a 0 : ces valeurs comptent
    des nuits, des pax ou des chambres, jamais des quantites negatives.
    """
    try:
        return max(0, int(to_float(s, default)))
    except (ValueError, TypeError):
        return default


def fmt(value) -> str:
    """Montant formate pour l'affichage : separateur de milliers, 2 decimales."""
    return f"{value:,.2f}"


def normalize(name: str) -> str:
    """Cle de comparaison d'un nom de ville ou d'hotel.

    Replie la casse et les accents, retire les parentheses et toute
    ponctuation, puis normalise les espaces. Sert a rapprocher deux libelles
    saisis differemment pour la meme realite.
    """
    if not name:
        return ""
    text = unicodedata.normalize("NFKD", str(name).strip().lower())
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()
