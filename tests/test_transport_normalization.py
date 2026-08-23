"""
Tests for transport city normalization, KM_MADA lookup robustness,
and distance calculation.

Cases observed in real data:
- "Antananarivo(1 jours)"  → "Antananarivo"
- "Antsirabe(2 jours)"     → "Antsirabe"
- "Ranohira (Isalo)"       → "Ranohira"
- "Tuler"                  → "Toliary"
- "Tulear"                 → "Toliary"
- Distance = abs(KM(arrivee) - KM(depart))
"""

import os
import sys
import unittest
from contextlib import contextmanager
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# Import apres l'ajustement de sys.path ci-dessus.
from utils.excel_handler import (  # noqa: E402
    get_km_mada_km_for_repere,
    get_segment_distance,
    normalize_city_name,
)


@contextmanager
def km_mada_dataset(entries):
    """Installe un jeu KM_MADA en memoire, sans dependre d'un classeur.

    `entries` est une liste de couples (repere, km).

    Neutralise _load_km_mada_rows : appelee sans data-hotel.xlsx sur le disque,
    elle appelle _invalidate_km_mada_cache() et vide donc le cache que le test
    vient d'injecter. Sans ce garde-fou les tests dependent du classeur local
    du poste de developpement et ne verifient plus rien en CI.
    """
    import utils.excel_handler as eh

    rows = [{"repere": r, "km": k, "duree": 0} for r, k in entries]
    lookup = eh._rebuild_km_mada_lookup(rows)
    with patch.object(eh, "_load_km_mada_rows", return_value=rows):
        with patch.dict(
            eh._KM_MADA_CACHE,
            {
                "lookup": lookup,
                "rows": rows,
                "path": None,
                "mtime": None,
                "loaded_at": float("inf"),
            },
        ):
            yield


class TestNormalizeCityName(unittest.TestCase):
    """normalize_city_name doit nettoyer les suffixes et alias métier."""

    def test_strip_duration_suffix_jours(self):
        self.assertEqual(normalize_city_name("Antananarivo(1 jours)"), "Antananarivo")

    def test_strip_duration_suffix_no_space(self):
        self.assertEqual(normalize_city_name("Antsirabe(2 jours)"), "Antsirabe")

    def test_strip_duration_suffix_jour_singular(self):
        self.assertEqual(normalize_city_name("Fianarantsoa(3 jour)"), "Fianarantsoa")

    def test_strip_duration_with_space(self):
        self.assertEqual(normalize_city_name("Morondava (4 jours)"), "Morondava")

    def test_alias_tuler(self):
        self.assertEqual(normalize_city_name("Tuler"), "Toliary")

    def test_alias_tulear(self):
        self.assertEqual(normalize_city_name("Tulear"), "Toliary")

    def test_alias_tulear_mixed_case(self):
        self.assertEqual(normalize_city_name("tulear"), "Toliary")

    def test_alias_ranohira_isalo(self):
        self.assertEqual(normalize_city_name("Ranohira (Isalo)"), "Ranohira")

    def test_alias_ranohira_isalo_uppercase(self):
        self.assertEqual(normalize_city_name("RANOHIRA (ISALO)"), "Ranohira")

    def test_plain_city_unchanged(self):
        self.assertEqual(normalize_city_name("Antsirabe"), "Antsirabe")

    def test_strips_whitespace(self):
        self.assertEqual(normalize_city_name("  Toliary  "), "Toliary")

    def test_empty_returns_empty(self):
        self.assertEqual(normalize_city_name(""), "")

    def test_none_returns_empty(self):
        self.assertEqual(normalize_city_name(None), "")

    def test_combined_suffix_and_alias(self):
        # Duration suffix removed first, then alias applied
        self.assertEqual(normalize_city_name("Tulear(1 jours)"), "Toliary")


class TestKmMadaLookupRobust(unittest.TestCase):
    """get_km_mada_km_for_repere doit gérer les doublons et la normalisation."""

    def _make_rows(self, entries):
        """entries: list of (repere, km)"""
        return [{"repere": r, "km": k, "duree": 0} for r, k in entries]

    def test_lookup_by_normalized_name(self):
        rows = self._make_rows([("ANTSIRABE", 169)])
        from utils.excel_handler import _rebuild_km_mada_lookup

        lookup = _rebuild_km_mada_lookup(rows)
        import utils.excel_handler as eh

        # _load_km_mada_rows() doit etre neutralise : sans classeur sur le
        # disque il appelle _invalidate_km_mada_cache(), qui efface le cache
        # injecte juste avant la lecture. Le test dependait donc du fichier
        # data-hotel.xlsx local et echouait sur un clone frais.
        with patch.object(eh, "_load_km_mada_rows", return_value=rows):
            with patch.dict(
                eh._KM_MADA_CACHE,
                {
                    "lookup": lookup,
                    "rows": rows,
                    "path": None,
                    "mtime": None,
                    "loaded_at": float("inf"),
                },
            ):
                result = get_km_mada_km_for_repere("antsirabe")
        self.assertEqual(result, 169)

    def test_lookup_city_with_duration_suffix(self):
        """City name with duration suffix must resolve via normalize_city_name."""
        with km_mada_dataset([("ANTSIRABE", 169)]):
            result_clean = get_km_mada_km_for_repere("Antsirabe")
            result_dirty = get_km_mada_km_for_repere("Antsirabe(2 jours)")
        self.assertEqual(result_clean, 169)
        self.assertEqual(result_dirty, 169)

    def test_lookup_alias_tuler(self):
        """'Tuler' must resolve to 'Toliary' km."""
        with km_mada_dataset([("TOLIARY", 936)]):
            km_toliary = get_km_mada_km_for_repere("Toliary")
            km_tuler = get_km_mada_km_for_repere("Tuler")
        self.assertEqual(km_toliary, 936)
        self.assertEqual(km_tuler, 936)

    def test_lookup_alias_ranohira_isalo(self):
        """'Ranohira (Isalo)' must resolve to same km as 'Ranohira'."""
        with km_mada_dataset([("RANOHIRA", 690)]):
            km_plain = get_km_mada_km_for_repere("Ranohira")
            km_isalo = get_km_mada_km_for_repere("Ranohira (Isalo)")
        self.assertEqual(km_plain, 690)
        self.assertEqual(km_isalo, 690)

    def test_duplicate_repere_prefers_nonzero(self):
        """When KM_MADA has duplicate repères, prefer km > 0."""
        rows = self._make_rows([("MORONDAVA", 0), ("MORONDAVA", 741)])
        with patch("utils.excel_handler._load_km_mada_rows", return_value=rows):
            with patch.dict(
                "utils.excel_handler._KM_MADA_CACHE",
                {
                    "lookup": {},
                    "rows": rows,
                    "path": None,
                    "mtime": None,
                    "loaded_at": 0.0,
                },
            ):
                from utils.excel_handler import _rebuild_km_mada_lookup

                lookup = _rebuild_km_mada_lookup(rows)
                best = lookup.get("morondava")
                self.assertIsNotNone(best)
                self.assertEqual(_parse_num_local(best.get("km", 0)), 741)

    def test_duplicate_repere_prefers_largest_km(self):
        rows = self._make_rows([("FIANARANTSOA", 200), ("FIANARANTSOA", 298)])
        from utils.excel_handler import _rebuild_km_mada_lookup

        lookup = _rebuild_km_mada_lookup(rows)
        best = lookup.get("fianarantsoa")
        self.assertIsNotNone(best)
        self.assertEqual(_parse_num_local(best.get("km", 0)), 298)


# Jeu de reference partage par les tests de distance. Valeurs de la RN7 telles
# qu'elles figurent dans KM_MADA : le km est cumule depuis Antananarivo.
RN7 = [
    ("ANTANANARIVO", 0),
    ("ANTSIRABE", 169),
    ("FIANARANTSOA", 411),
    ("RANOHIRA", 690),
    ("TOLIARY", 936),
]


class TestSegmentDistance(unittest.TestCase):
    """get_segment_distance doit retourner abs(km_arr - km_dep)."""

    def test_segment_antananarivo_to_antsirabe(self):
        """Antananarivo = 0 km (origine), Antsirabe = 169 km → distance = 169."""
        with km_mada_dataset(RN7):
            self.assertEqual(get_segment_distance("Antananarivo", "Antsirabe"), 169)

    def test_segment_antsirabe_to_fianarantsoa(self):
        with km_mada_dataset(RN7):
            self.assertEqual(get_segment_distance("Antsirabe", "Fianarantsoa"), 242)

    def test_segment_is_symmetric(self):
        """La distance ne depend pas du sens du trajet."""
        with km_mada_dataset(RN7):
            aller = get_segment_distance("Antsirabe", "Fianarantsoa")
            retour = get_segment_distance("Fianarantsoa", "Antsirabe")
        self.assertEqual(aller, 242)
        self.assertEqual(retour, 242)

    def test_segment_with_dirty_names(self):
        """Dirty city names must normalize before lookup."""
        with km_mada_dataset(RN7):
            dist_clean = get_segment_distance("Antsirabe", "Fianarantsoa")
            dist_dirty = get_segment_distance(
                "Antsirabe(2 jours)", "Fianarantsoa(3 jour)"
            )
        self.assertEqual(dist_clean, 242)
        self.assertEqual(dist_dirty, 242)

    def test_segment_with_alias(self):
        with km_mada_dataset(RN7):
            dist_clean = get_segment_distance("Ranohira", "Toliary")
            dist_dirty = get_segment_distance("Ranohira (Isalo)", "Tulear")
        self.assertEqual(dist_clean, 246)
        self.assertEqual(dist_dirty, 246)

    def test_unknown_departure_falls_back_to_arrival_km(self):
        """If depart is unknown, fall back to km(arrivee)."""
        with km_mada_dataset(RN7):
            dist = get_segment_distance("VilleInconnueXYZ", "Antsirabe")
        self.assertEqual(dist, 169)

    def test_both_unknown_returns_zero(self):
        with km_mada_dataset(RN7):
            dist = get_segment_distance("InconnuA", "InconnuB")
        self.assertEqual(dist, 0)


def _parse_num_local(v):
    try:
        return float(str(v).replace(",", ".").strip() or 0)
    except Exception:
        return 0.0


if __name__ == "__main__":
    unittest.main()
