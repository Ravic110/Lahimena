"""Correctifs appliques a CustomTkinter au demarrage.

Ces correctifs atteignent onze elements d'API privee de CustomTkinter, par des
chemins de modules profonds (`customtkinter.windows.widgets.scaling.
scaling_tracker`, entre autres). Tant qu'ils vivaient dans main.py sous forme
d'imports nus, une version de CustomTkinter deplacant l'un de ces modules
empechait l'application de demarrer -- et l'echec survenait avant que le
gestionnaire d'erreur de main() n'existe, donc sans message pour l'utilisateur.

Ils sont desormais facultatifs : absents, l'application demarre sans eux.
"""

import builtins
import tkinter as tk

import pytest

from gui import ctk_patch


@pytest.fixture(autouse=True)
def _etat_vierge(monkeypatch):
    """Chaque test part d'un correctif non applique."""
    import customtkinter as ctk

    monkeypatch.setattr(ctk, "_lahimena_tracker_patch_applied", False, raising=False)
    yield


class TestApplication:
    def test_appliquer_reussit_avec_la_version_installee(self):
        assert ctk_patch.appliquer() is True

    def test_deuxieme_appel_sans_effet(self):
        assert ctk_patch.appliquer() is True
        assert ctk_patch.appliquer() is True

    def test_marque_customtkinter_une_fois_applique(self):
        import customtkinter as ctk

        ctk_patch.appliquer()
        assert getattr(ctk, "_lahimena_tracker_patch_applied", False) is True


class TestToleranceAuxVersions:
    """Un interne deplace doit degrader, jamais empecher le demarrage."""

    def test_import_manquant_renvoie_faux_sans_lever(self, monkeypatch):
        vrai_import = builtins.__import__

        def _import_qui_echoue(nom, *args, **kwargs):
            if "scaling_tracker" in nom:
                raise ImportError(f"No module named {nom!r}")
            return vrai_import(nom, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", _import_qui_echoue)
        assert ctk_patch.appliquer() is False

    def test_attribut_manquant_renvoie_faux_sans_lever(self, monkeypatch):
        def _charger_incomplet():
            raise AttributeError("ScalingTracker n'a plus window_widgets_dict")

        monkeypatch.setattr(ctk_patch, "_charger_les_traceurs", _charger_incomplet)
        assert ctk_patch.appliquer() is False

    def test_l_echec_est_journalise(self, monkeypatch, caplog):
        def _charger_incomplet():
            raise AttributeError("interne deplace")

        monkeypatch.setattr(ctk_patch, "_charger_les_traceurs", _charger_incomplet)
        with caplog.at_level("WARNING"):
            ctk_patch.appliquer()
        assert any("CustomTkinter" in m for m in caplog.messages)


class TestSuppressionDeCommandeTcl:
    """`deletecommand` doit avaler la double-suppression, race connue de CTk."""

    def test_supprimer_deux_fois_ne_leve_pas(self):
        ctk_patch.appliquer()
        racine = tk.Tk()
        racine.withdraw()
        try:
            nom = racine.register(lambda: None)
            racine.deletecommand(nom)
            racine.deletecommand(nom)  # ne doit pas lever
        finally:
            racine.destroy()
