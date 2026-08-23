"""Helpers partages par les cinq ecrans de cotation client.

Ces fonctions etaient recopiees dans chacun des cinq modules
`gui/forms/client_*_cotation.py`. Elles n'y avaient aucun test : la couche
`gui/` est a 0 % de couverture. Les regrouper permet de les couvrir une fois.

Attention a ce qui n'est PAS ici : `_make_row`, `_parse_cities` et
`_compute_prix_unitaire` existent aussi en plusieurs exemplaires, mais leurs
versions divergent sur de la logique metier reelle. Les fusionner serait faux.
"""

import pytest

from gui.forms.cotation_commons import fmt, normalize, to_float, to_int


class TestToFloat:
    @pytest.mark.parametrize(
        "entree, attendu",
        [
            ("12.5", 12.5),
            ("12,5", 12.5),  # virgule decimale francaise
            ("  42  ", 42.0),
            (7, 7.0),
            (7.5, 7.5),
        ],
    )
    def test_valeurs_convertibles(self, entree, attendu):
        assert to_float(entree) == attendu

    @pytest.mark.parametrize("entree", ["", "   ", None, "abc", "12,5,3", []])
    def test_valeurs_non_convertibles_donnent_le_defaut(self, entree):
        assert to_float(entree) == 0.0

    def test_defaut_personnalise(self):
        assert to_float("abc", default=9.0) == 9.0

    def test_chaine_vide_donne_le_defaut(self):
        """`float(str(s).strip() or default)` : le defaut passe par float()."""
        assert to_float("", default=3.5) == 3.5

    def test_negatif(self):
        assert to_float("-12,5") == -12.5


class TestToInt:
    @pytest.mark.parametrize(
        "entree, attendu",
        [("3", 3), ("  7 ", 7), (4, 4), ("3.9", 3), ("3,9", 3)],
    )
    def test_valeurs_convertibles(self, entree, attendu):
        assert to_int(entree) == attendu

    @pytest.mark.parametrize("entree", ["", None, "abc"])
    def test_valeurs_non_convertibles_donnent_le_defaut(self, entree):
        assert to_int(entree) == 0

    def test_defaut_personnalise(self):
        assert to_int("abc", default=2) == 2

    def test_troncature_et_non_arrondi(self):
        """Le comportement historique tronque : 3.9 devient 3, pas 4."""
        assert to_int("3.9") == 3


class TestFmt:
    @pytest.mark.parametrize(
        "entree, attendu",
        [
            (0, "0.00"),
            (1234.5, "1,234.50"),
            (1000000, "1,000,000.00"),
            (-42.125, "-42.12"),  # arrondi bancaire de format()
        ],
    )
    def test_formatage(self, entree, attendu):
        assert fmt(entree) == attendu

    def test_separateur_de_milliers_present(self):
        assert fmt(1234567.89) == "1,234,567.89"


class TestNormalize:
    def test_casse_ignoree(self):
        assert normalize("ANTSIRABE") == normalize("antsirabe")

    def test_accents_ignores(self):
        assert normalize("Fianarantsoa") == normalize("Fianarantsoä")

    def test_espaces_replies(self):
        assert normalize("  Ranohira   Isalo ") == normalize("Ranohira Isalo")

    def test_valeur_vide(self):
        assert normalize("") == ""
        assert normalize(None) == ""

    def test_villes_distinctes_le_restent(self):
        assert normalize("Antsirabe") != normalize("Antananarivo")


class TestToIntClampeLesNegatifs:
    """Ces entiers comptent des nuits, des pax, des chambres.

    Le comportement historique ramene un negatif a 0 plutot que de le
    propager dans un calcul de total.
    """

    @pytest.mark.parametrize("entree", ["-3", -3, "-0.5"])
    def test_negatif_devient_zero(self, entree):
        assert to_int(entree) == 0

    def test_zero_reste_zero(self):
        assert to_int("0") == 0
