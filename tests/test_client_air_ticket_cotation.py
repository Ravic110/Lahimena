"""
Tests pour la cotation avion client :
- helpers de génération, calcul et validation
- navigation home_page → main_content
- comportement _populate_initial_rows
"""

import sys
import os
import types
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from gui.forms.client_air_ticket_cotation import (
    _build_initial_rows,
    _make_row,
    _recompute_row,
    _validate_rows,
)


# ── Génération initiale ────────────────────────────────────────────────────────

class TestBuildInitialRows(unittest.TestCase):

    def _client(self):
        return {
            "compagnie":       "Air Austral",
            "ville_depart":    "Antananarivo",
            "ville_arrivee":   "Nosy Be",
            "nombre_adultes":  "2",
            "nombre_enfants":  "1",
        }

    def test_round_trip_generates_two_rows(self):
        rows = _build_initial_rows(self._client(), "aller-retour")
        self.assertEqual(len(rows), 2)

    def test_round_trip_aller_fields(self):
        rows = _build_initial_rows(self._client(), "aller-retour")
        self.assertEqual(rows[0]["type_trajet"],  "aller")
        self.assertEqual(rows[0]["compagnie"],    "Air Austral")
        self.assertEqual(rows[0]["ville_depart"], "Antananarivo")
        self.assertEqual(rows[0]["ville_arrivee"],"Nosy Be")

    def test_round_trip_retour_inverts_cities(self):
        rows = _build_initial_rows(self._client(), "aller-retour")
        self.assertEqual(rows[1]["type_trajet"],  "retour")
        self.assertEqual(rows[1]["ville_depart"], "Nosy Be")
        self.assertEqual(rows[1]["ville_arrivee"],"Antananarivo")

    def test_aller_simple_generates_one_row(self):
        rows = _build_initial_rows(self._client(), "aller simple")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["type_trajet"], "aller")

    def test_passengers_propagated(self):
        rows = _build_initial_rows(self._client(), "aller-retour")
        for row in rows:
            self.assertEqual(row["nb_adultes"], "2")
            self.assertEqual(row["nb_enfants"], "1")


# ── Recalcul ──────────────────────────────────────────────────────────────────

class TestRecomputeRow(unittest.TestCase):

    def _base_row(self, **kwargs):
        defaults = dict(
            type_trajet="aller",
            compagnie="Air Austral",
            ville_depart="Antananarivo",
            ville_arrivee="Nosy Be",
            nb_adultes="2",
            nb_enfants="1",
            tarif_adulte="500",
            tarif_enfant="200",
            marge_pct="10",
        )
        defaults.update(kwargs)
        return _make_row(**defaults)

    def test_montants_calculés(self):
        row = _recompute_row(self._base_row())
        self.assertEqual(row["montant_adultes"], 1000.0)
        self.assertEqual(row["montant_enfants"], 200.0)
        self.assertEqual(row["sous_total"],      1200.0)

    def test_marge_appliquee(self):
        row = _recompute_row(self._base_row())
        self.assertAlmostEqual(row["total"], 1320.0)

    def test_total_manuel_preserve(self):
        row = _recompute_row(self._base_row(total="1500", total_manuel=True))
        self.assertEqual(row["sous_total"], 1200.0)
        self.assertEqual(row["total"],      1500.0)
        self.assertTrue(row["total_manuel"])

    def test_total_manuel_vide_revert_auto(self):
        row = _recompute_row(self._base_row(total="", total_manuel=True))
        self.assertAlmostEqual(row["total"], 1320.0)
        self.assertFalse(row["total_manuel"])

    def test_zero_pax_gives_zero_total(self):
        row = _recompute_row(self._base_row(nb_adultes="0", nb_enfants="0"))
        self.assertEqual(row["total"], 0.0)


# ── Validation ────────────────────────────────────────────────────────────────

class TestValidateRows(unittest.TestCase):

    def test_empty_rows_no_errors(self):
        self.assertEqual(_validate_rows([]), [])

    def test_missing_ville_depart(self):
        row = _make_row(ville_depart="", ville_arrivee="Nosy Be")
        errors = _validate_rows([row])
        self.assertIn("Ligne 1 : ville_depart et ville_arrivee sont obligatoires.", errors)

    def test_missing_ville_arrivee(self):
        row = _make_row(ville_depart="Antananarivo", ville_arrivee="")
        errors = _validate_rows([row])
        self.assertIn("Ligne 1 : ville_depart et ville_arrivee sont obligatoires.", errors)

    def test_negative_nb_adultes(self):
        row = _make_row(ville_depart="Antananarivo", ville_arrivee="Nosy Be", nb_adultes="-1")
        errors = _validate_rows([row])
        self.assertIn("Ligne 1 : nb_adultes doit etre numerique et >= 0.", errors)

    def test_valid_row_no_errors(self):
        row = _make_row(
            ville_depart="Antananarivo", ville_arrivee="Nosy Be",
            nb_adultes="2", nb_enfants="1",
            tarif_adulte="500", tarif_enfant="200", marge_pct="10",
        )
        self.assertEqual(_validate_rows([row]), [])

    def test_multiple_rows_multiple_errors(self):
        rows = [
            _make_row(ville_depart="", ville_arrivee="Nosy Be"),
            _make_row(ville_depart="Antananarivo", ville_arrivee="Nosy Be", nb_adultes="-1"),
        ]
        errors = _validate_rows(rows)
        self.assertIn("Ligne 1 : ville_depart et ville_arrivee sont obligatoires.", errors)
        self.assertIn("Ligne 2 : nb_adultes doit etre numerique et >= 0.", errors)


# ── Navigation ─────────────────────────────────────────────────────────────────

class TestNavigation(unittest.TestCase):

    def test_home_page_air_ticket_cotation_routes_selected_client(self):
        from gui.forms.home_page import HomePage
        calls = []
        page = HomePage.__new__(HomePage)
        page.navigate_callback = lambda route, **kwargs: calls.append((route, kwargs))
        client = {"ref_client": "CLI001"}

        HomePage._on_air_ticket_cotation(page, client)

        self.assertEqual(calls, [("client_air_ticket_cotation", {"client": client})])

    def test_main_content_show_client_air_ticket_cotation_uses_nav_kwargs(self):
        from gui.main_content import MainContent
        captured = {}

        class DummyScreen:
            def __init__(self, parent, client, on_back=None):
                captured["parent"] = parent
                captured["client"] = client
                captured["on_back"] = on_back

        dummy_module = types.ModuleType("gui.forms.client_air_ticket_cotation")
        dummy_module.ClientAirTicketCotation = DummyScreen
        sys.modules["gui.forms.client_air_ticket_cotation"] = dummy_module

        content = MainContent.__new__(MainContent)
        content.main_scroll = object()
        content._nav_kwargs = {"client": {"ref_client": "CLI001"}}
        content.update_content = lambda route: captured.setdefault("route", route)

        MainContent._show_client_air_ticket_cotation(content)

        self.assertEqual(captured["client"], {"ref_client": "CLI001"})
        self.assertTrue(callable(captured["on_back"]))

        # Nettoyage
        del sys.modules["gui.forms.client_air_ticket_cotation"]


# ── Comportement _populate_initial_rows ───────────────────────────────────────

class TestPopulateInitialRows(unittest.TestCase):

    def _fake_screen(self, client=None):
        from gui.forms.client_air_ticket_cotation import ClientAirTicketCotation
        fake = ClientAirTicketCotation.__new__(ClientAirTicketCotation)
        fake.client = client or {"ref_client": "CLI001"}
        fake.trip_mode_var = types.SimpleNamespace(get=lambda: "aller-retour")
        fake._rows = []
        return fake

    def test_prefer_saved_rows(self):
        import unittest.mock as mock
        from gui.forms.client_air_ticket_cotation import ClientAirTicketCotation

        saved = [_make_row(
            type_trajet="aller", compagnie="Air Austral",
            ville_depart="Antananarivo", ville_arrivee="Nosy Be",
            nb_adultes="2", nb_enfants="1",
            tarif_adulte="500", tarif_enfant="200",
            marge_pct="10", total="1500", total_manuel=True,
        )]

        fake = self._fake_screen()
        with mock.patch(
            "gui.forms.client_air_ticket_cotation.load_client_air_ticket_cotation",
            return_value=saved,
        ):
            ClientAirTicketCotation._populate_initial_rows(fake)

        self.assertEqual(len(fake._rows), 1)
        self.assertTrue(fake._rows[0]["total_manuel"])
        self.assertEqual(fake._rows[0]["total"], 1500.0)

    def test_falls_back_to_generation_when_no_saved(self):
        import unittest.mock as mock
        from gui.forms.client_air_ticket_cotation import ClientAirTicketCotation

        client = {
            "compagnie": "Air Austral",
            "ville_depart": "Antananarivo",
            "ville_arrivee": "Nosy Be",
            "nombre_adultes": "2",
            "nombre_enfants": "1",
        }
        fake = self._fake_screen(client=client)
        with mock.patch(
            "gui.forms.client_air_ticket_cotation.load_client_air_ticket_cotation",
            return_value=[],
        ):
            ClientAirTicketCotation._populate_initial_rows(fake)

        self.assertEqual(len(fake._rows), 2)
        self.assertEqual(fake._rows[0]["type_trajet"], "aller")
        self.assertEqual(fake._rows[1]["type_trajet"], "retour")


if __name__ == "__main__":
    unittest.main()
