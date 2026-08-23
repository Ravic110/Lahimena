"""Cache des tables de reference, invalide par la date du fichier.

Quatre chargeurs -- circuits, transport, avion, visites/excursions --
reparsaient le classeur entier a chaque appel, soit ~500 ms pieces sur les
donnees reelles. L'interface les invoque depuis les rappels de listes
deroulantes : choisir un prestataire gelait l'ecran une demi-seconde, puis
encore autant au choix du vehicule.

L'invalidation se fait sur la date de modification du fichier, et non par des
appels manuels : une ecriture change le mtime, donc le cache se perime tout
seul. Aucun risque d'oublier une invalidation apres un `save_`.
"""

import os
import time

import pytest

from utils.storage.source_cache import mtime_cached, vider_les_caches_de_source


@pytest.fixture(autouse=True)
def _caches_vierges():
    vider_les_caches_de_source()
    yield
    vider_les_caches_de_source()


@pytest.fixture
def fichier(tmp_path):
    chemin = tmp_path / "source.txt"
    chemin.write_text("v1", encoding="utf-8")
    return str(chemin)


class TestCache:
    def test_le_second_appel_ne_relit_pas(self, fichier):
        appels = []

        @mtime_cached(lambda: fichier)
        def charger():
            appels.append(1)
            return open(fichier, encoding="utf-8").read()

        assert charger() == "v1"
        assert charger() == "v1"
        assert len(appels) == 1

    def test_une_modification_perime_le_cache(self, fichier):
        appels = []

        @mtime_cached(lambda: fichier)
        def charger():
            appels.append(1)
            return open(fichier, encoding="utf-8").read()

        assert charger() == "v1"
        time.sleep(0.01)
        with open(fichier, "w", encoding="utf-8") as f:
            f.write("v2")

        assert charger() == "v2"
        assert len(appels) == 2

    def test_une_ecriture_de_meme_taille_perime_aussi(self, fichier):
        """Le mtime seul suffit : nanosecondes, pas secondes."""
        appels = []

        @mtime_cached(lambda: fichier)
        def charger():
            appels.append(1)
            return open(fichier, encoding="utf-8").read()

        charger()
        time.sleep(0.01)
        with open(fichier, "w", encoding="utf-8") as f:
            f.write("v9")  # meme longueur que "v1"

        assert charger() == "v9"
        assert len(appels) == 2

    def test_fichier_absent_n_est_pas_mis_en_cache(self, tmp_path):
        """Sinon un classeur cree apres coup resterait invisible."""
        chemin = str(tmp_path / "pas-encore.txt")
        appels = []

        @mtime_cached(lambda: chemin)
        def charger():
            appels.append(1)
            if not os.path.exists(chemin):
                return []
            return open(chemin, encoding="utf-8").read()

        assert charger() == []
        assert charger() == []
        assert len(appels) == 2  # rejoue tant que le fichier n'existe pas

        with open(chemin, "w", encoding="utf-8") as f:
            f.write("enfin")
        assert charger() == "enfin"

    def test_deux_fonctions_ont_des_entrees_distinctes(self, fichier):
        @mtime_cached(lambda: fichier)
        def une():
            return "A"

        @mtime_cached(lambda: fichier)
        def deux():
            return "B"

        assert une() == "A"
        assert deux() == "B"

    def test_le_chemin_est_relu_a_chaque_appel(self, tmp_path):
        """Les tests redirigent les classeurs : le chemin ne doit pas etre fige."""
        a = tmp_path / "a.txt"
        b = tmp_path / "b.txt"
        a.write_text("A", encoding="utf-8")
        b.write_text("B", encoding="utf-8")
        courant = {"chemin": str(a)}

        @mtime_cached(lambda: courant["chemin"])
        def charger():
            return open(courant["chemin"], encoding="utf-8").read()

        assert charger() == "A"
        courant["chemin"] = str(b)
        assert charger() == "B"

    def test_vider_force_la_relecture(self, fichier):
        appels = []

        @mtime_cached(lambda: fichier)
        def charger():
            appels.append(1)
            return "x"

        charger()
        vider_les_caches_de_source()
        charger()
        assert len(appels) == 2

    def test_le_nom_et_la_docstring_sont_conserves(self):
        @mtime_cached(lambda: "/inexistant")
        def charger_quelque_chose():
            """Docstring d'origine."""
            return None

        assert charger_quelque_chose.__name__ == "charger_quelque_chose"
        assert charger_quelque_chose.__doc__ == "Docstring d'origine."


class TestProtectionDuCache:
    """Un appelant qui modifie la liste recue ne doit pas corrompre le cache."""

    def test_modifier_la_liste_recue_n_affecte_pas_les_appels_suivants(self, fichier):
        @mtime_cached(lambda: fichier)
        def charger():
            return [{"nom": "A"}, {"nom": "B"}]

        premiere = charger()
        premiere.append({"nom": "INTRUS"})
        premiere.sort(key=lambda r: r["nom"], reverse=True)

        seconde = charger()
        assert len(seconde) == 2
        assert [r["nom"] for r in seconde] == ["A", "B"]

    def test_les_valeurs_non_listes_traversent_telles_quelles(self, fichier):
        @mtime_cached(lambda: fichier)
        def charger():
            return {"cle": "valeur"}

        assert charger() == {"cle": "valeur"}
