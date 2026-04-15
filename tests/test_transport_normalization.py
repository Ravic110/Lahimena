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

import sys
import os
import unittest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from utils.excel_handler import (
    normalize_city_name,
    get_km_mada_km_for_repere,
    get_segment_distance,
)


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
        with patch.dict(eh._KM_MADA_CACHE, {"lookup": lookup, "rows": rows,
                                             "path": None, "mtime": None, "loaded_at": float("inf")}):
            result = get_km_mada_km_for_repere("antsirabe")
        self.assertEqual(result, 169)

    def test_lookup_city_with_duration_suffix(self):
        """City name with duration suffix must resolve via normalize_city_name."""
        result_clean = get_km_mada_km_for_repere("Antsirabe")
        result_dirty = get_km_mada_km_for_repere("Antsirabe(2 jours)")
        # Both should return the same value (or both 0 if not in BD)
        self.assertEqual(result_clean, result_dirty)

    def test_lookup_alias_tuler(self):
        """'Tuler' must resolve to 'Toliary' km."""
        km_toliary = get_km_mada_km_for_repere("Toliary")
        km_tuler   = get_km_mada_km_for_repere("Tuler")
        self.assertEqual(km_toliary, km_tuler)

    def test_lookup_alias_ranohira_isalo(self):
        """'Ranohira (Isalo)' must resolve to same km as 'Ranohira'."""
        km_plain = get_km_mada_km_for_repere("Ranohira")
        km_isalo = get_km_mada_km_for_repere("Ranohira (Isalo)")
        self.assertEqual(km_plain, km_isalo)

    def test_duplicate_repere_prefers_nonzero(self):
        """When KM_MADA has duplicate repères, prefer km > 0."""
        rows = self._make_rows([("MORONDAVA", 0), ("MORONDAVA", 741)])
        with patch("utils.excel_handler._load_km_mada_rows", return_value=rows):
            with patch.dict("utils.excel_handler._KM_MADA_CACHE",
                            {"lookup": {}, "rows": rows, "path": None,
                             "mtime": None, "loaded_at": 0.0}):
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


class TestSegmentDistance(unittest.TestCase):
    """get_segment_distance doit retourner abs(km_arr - km_dep)."""

    def test_segment_antananarivo_to_antsirabe(self):
        """Antananarivo = 0 km (origin), Antsirabe = 169 km → distance = 169."""
        km_dep = get_km_mada_km_for_repere("Antananarivo")
        km_arr = get_km_mada_km_for_repere("Antsirabe")
        dist   = get_segment_distance("Antananarivo", "Antsirabe")
        expected = abs(km_arr - km_dep)
        self.assertEqual(dist, expected)

    def test_segment_antsirabe_to_fianarantsoa(self):
        km_dep = get_km_mada_km_for_repere("Antsirabe")
        km_arr = get_km_mada_km_for_repere("Fianarantsoa")
        dist   = get_segment_distance("Antsirabe", "Fianarantsoa")
        expected = abs(km_arr - km_dep)
        self.assertEqual(dist, expected)

    def test_segment_with_dirty_names(self):
        """Dirty city names must normalize before lookup."""
        dist_clean = get_segment_distance("Antsirabe", "Fianarantsoa")
        dist_dirty = get_segment_distance("Antsirabe(2 jours)", "Fianarantsoa(3 jour)")
        self.assertEqual(dist_clean, dist_dirty)

    def test_segment_with_alias(self):
        dist_clean = get_segment_distance("Ranohira", "Toliary")
        dist_dirty = get_segment_distance("Ranohira (Isalo)", "Tulear")
        self.assertEqual(dist_clean, dist_dirty)

    def test_unknown_departure_falls_back_to_arrival_km(self):
        """If depart is unknown, fall back to km(arrivee)."""
        km_arr = get_km_mada_km_for_repere("Antsirabe")
        dist   = get_segment_distance("VilleInconnueXYZ", "Antsirabe")
        self.assertEqual(dist, km_arr)

    def test_both_unknown_returns_zero(self):
        dist = get_segment_distance("InconnuA", "InconnuB")
        self.assertEqual(dist, 0)


def _parse_num_local(v):
    try:
        return float(str(v).replace(",", ".").strip() or 0)
    except Exception:
        return 0.0


if __name__ == "__main__":
    unittest.main()
