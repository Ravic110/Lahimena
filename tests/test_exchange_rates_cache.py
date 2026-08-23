"""Mise en cache des taux de change.

`get_exchange_rates()` interroge une API HTTP depuis le fil de l'interface.
Deux ecrans de cotation -- hotellerie et restauration -- l'appellent dans leur
constructeur. Sans cache, chaque ouverture d'ecran gele l'interface le temps de
l'aller-retour reseau, jusqu'a 5 secondes de timeout quand la connexion est
mauvaise. Le decorateur `cached_exchange_rates` existait deja dans
utils/cache.py mais n'etait pose sur rien.
"""

import pytest

from utils import validators
from utils.cache import _exchange_rate_cache


@pytest.fixture(autouse=True)
def _cache_vierge():
    _exchange_rate_cache.clear()
    yield
    _exchange_rate_cache.clear()


class TestCache:
    def test_un_seul_appel_reseau_pour_deux_lectures(self, monkeypatch):
        appels = []

        class _Reponse:
            @staticmethod
            def json():
                appels.append(1)
                return {"rates": {"EUR": 1 / 5000.0, "USD": 1 / 4500.0}}

        monkeypatch.setattr(
            validators.requests, "get", lambda *a, **k: _Reponse(), raising=False
        )

        premier = validators.get_exchange_rates()
        second = validators.get_exchange_rates()

        assert len(appels) == 1
        assert premier == second

    def test_les_taux_restent_justes(self, monkeypatch):
        """Le cache ne doit pas alterer la conversion : MGA par unite."""

        class _Reponse:
            @staticmethod
            def json():
                return {"rates": {"EUR": 1 / 5000.0, "USD": 1 / 4500.0}}

        monkeypatch.setattr(
            validators.requests, "get", lambda *a, **k: _Reponse(), raising=False
        )

        taux = validators.get_exchange_rates()
        assert taux["EUR"] == pytest.approx(5000.0)
        assert taux["USD"] == pytest.approx(4500.0)

    def test_l_echec_reseau_retombe_sur_les_taux_de_secours(self, monkeypatch):
        def _echouer(*a, **k):
            raise OSError("reseau indisponible")

        monkeypatch.setattr(validators.requests, "get", _echouer, raising=False)

        taux = validators.get_exchange_rates()
        assert taux["EUR"] == 5235.0
        assert taux["USD"] == 4900.0

    def test_le_repli_est_aussi_mis_en_cache(self, monkeypatch):
        """Sans cela, une agence hors ligne refait un timeout par ecran."""
        appels = []

        def _echouer(*a, **k):
            appels.append(1)
            raise OSError("reseau indisponible")

        monkeypatch.setattr(validators.requests, "get", _echouer, raising=False)

        validators.get_exchange_rates()
        validators.get_exchange_rates()
        validators.get_exchange_rates()

        assert len(appels) == 1
