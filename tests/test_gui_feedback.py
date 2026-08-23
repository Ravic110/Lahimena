"""Traduction des sentinelles de stockage en messages utilisateur."""

import pytest

from gui.feedback import ECHEC, FICHIER_VERROUILLE, message_d_echec


class TestSucces:
    @pytest.mark.parametrize("resultat", [1, 2, 42, 0, True, None])
    def test_aucun_message(self, resultat):
        """0 est le succes des update_*, un entier positif celui des save_*."""
        assert message_d_echec(resultat) is None


class TestEchec:
    def test_fichier_verrouille_nomme_le_fichier(self):
        message = message_d_echec(FICHIER_VERROUILLE, fichier="data-hotel.xlsx")
        assert "data-hotel.xlsx" in message
        assert "Fermez" in message

    def test_fichier_par_defaut(self):
        assert "data.xlsx" in message_d_echec(FICHIER_VERROUILLE)

    def test_echec_generique_renvoie_vers_les_journaux(self):
        message = message_d_echec(ECHEC)
        assert message is not None
        assert "journaux" in message

    def test_false_est_l_echec_des_suppressions(self):
        assert message_d_echec(False) is not None

    def test_true_est_leur_succes(self):
        assert message_d_echec(True) is None


class TestRobustesse:
    def test_false_n_est_pas_confondu_avec_zero(self):
        """En Python False == 0 : la distinction doit etre faite par identite."""
        assert message_d_echec(0) is None
        assert message_d_echec(False) is not None

    def test_true_n_est_pas_confondu_avec_un(self):
        assert message_d_echec(1) is None
        assert message_d_echec(True) is None

    @pytest.mark.parametrize("resultat", ["", "erreur", [], {}, 3.5])
    def test_valeur_inattendue_ne_leve_pas(self, resultat):
        assert message_d_echec(resultat) is None
