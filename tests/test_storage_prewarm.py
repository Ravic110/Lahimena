"""Prechauffage des caches de reference au demarrage."""

import threading

import pytest

from utils.storage import prewarm


@pytest.fixture
def chargeurs_factices(monkeypatch):
    appels = []

    def _faire(nom, echoue=False):
        def _charger():
            appels.append(nom)
            if echoue:
                raise RuntimeError(f"{nom} indisponible")
            return [nom]

        return _charger

    monkeypatch.setattr(
        prewarm,
        "_chargeurs",
        lambda: [("un", _faire("un")), ("deux", _faire("deux"))],
    )
    return appels


class TestPrechauffage:
    def test_bloquant_appelle_tous_les_chargeurs(self, chargeurs_factices):
        prewarm.prechauffer(bloquant=True)
        assert chargeurs_factices == ["un", "deux"]

    def test_non_bloquant_rend_la_main_et_travaille(self, chargeurs_factices):
        fil = prewarm.prechauffer()
        assert isinstance(fil, threading.Thread)
        fil.join(timeout=5)
        assert not fil.is_alive()
        assert sorted(chargeurs_factices) == ["deux", "un"]

    def test_le_fil_est_demon(self, chargeurs_factices):
        """L'application ne doit pas rester ouverte a cause du prechauffage."""
        fil = prewarm.prechauffer()
        assert fil.daemon
        fil.join(timeout=5)

    def test_un_chargeur_en_echec_n_arrete_pas_les_autres(self, monkeypatch):
        """Le prechauffage est une optimisation : son echec est sans effet."""
        appels = []

        def _casse():
            appels.append("casse")
            raise RuntimeError("classeur absent")

        def _marche():
            appels.append("marche")

        monkeypatch.setattr(
            prewarm, "_chargeurs", lambda: [("a", _casse), ("b", _marche)]
        )
        prewarm.prechauffer(bloquant=True)
        assert appels == ["casse", "marche"]

    def test_sans_classeur_ne_leve_pas(self, monkeypatch, tmp_path):
        """Cas du poste neuf : les classeurs n'existent pas encore."""
        import config

        monkeypatch.setattr(config, "CLIENT_EXCEL_PATH", str(tmp_path / "absent.xlsx"))
        monkeypatch.setattr(config, "HOTEL_EXCEL_PATH", str(tmp_path / "absent2.xlsx"))
        prewarm.prechauffer(bloquant=True)


class TestChargeurs:
    def test_la_liste_est_composee_d_appelables(self):
        for libelle, charger in prewarm._chargeurs():
            assert isinstance(libelle, str) and libelle
            assert callable(charger)

    def test_le_catalogue_hotels_vient_en_premier(self):
        """C'est la lecture la plus couteuse : ~1,7 s a froid."""
        assert prewarm._chargeurs()[0][0] == "catalogue hotels"
