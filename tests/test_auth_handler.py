"""Tests du gestionnaire d'authentification.

Module critique reste longtemps sans filet : 314 instructions, 0 % de
couverture. Il porte le hachage, la migration SHA-256 -> PBKDF2, le
verrouillage anti-force-brute, l'expiration des mots de passe et la protection
du dernier administrateur.
"""

import json
import os
from datetime import datetime, timedelta

import pytest

import utils.auth_handler as ah


@pytest.fixture(autouse=True)
def bac_a_sable(tmp_path, monkeypatch):
    """Fichier utilisateurs isole, journal neutralise, etat global remis a zero."""
    monkeypatch.setattr(ah, "USERS_FILE", str(tmp_path / "users.json"))
    monkeypatch.setattr(
        "utils.activity_log.log_activity", lambda *a, **k: None, raising=False
    )
    ah._failed_attempts.clear()
    ah.set_current_user(None)
    yield
    ah._failed_attempts.clear()
    ah.set_current_user(None)


def _horodatage(jours_avant=0):
    return (datetime.now() - timedelta(days=jours_avant)).strftime("%Y-%m-%d %H:%M:%S")


def _ecrire_utilisateurs(entrees):
    with open(ah.USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(entrees, f)


# ── Hachage ───────────────────────────────────────────────────────────────────


class TestHachage:
    def test_le_sel_est_unique_a_chaque_appel(self):
        assert ah._generate_salt() != ah._generate_salt()

    def test_le_sel_fait_256_bits(self):
        assert len(ah._generate_salt()) == 64  # 32 octets en hexadecimal

    def test_le_hachage_est_deterministe(self):
        assert ah._hash_password("motdepasse", "sel") == ah._hash_password(
            "motdepasse", "sel"
        )

    def test_deux_sels_donnent_deux_empreintes(self):
        assert ah._hash_password("motdepasse", "selA") != ah._hash_password(
            "motdepasse", "selB"
        )

    def test_l_algorithme_courant_est_le_v2(self):
        assert ah._hash_password("m", "s") == ah._hash_password_v2("m", "s")
        assert ah._hash_password("m", "s") != ah._hash_password_v1("m", "s")


# ── Creation de comptes ───────────────────────────────────────────────────────


class TestCreationDeCompte:
    def test_creation_nominale(self):
        ok, err = ah.create_user("victorien", "motdepasse", "admin")
        assert ok and err == ""
        assert ah.has_users()

    def test_le_mot_de_passe_n_est_pas_stocke_en_clair(self):
        ah.create_user("victorien", "monSecret123", "admin")
        contenu = open(ah.USERS_FILE, encoding="utf-8").read()
        assert "monSecret123" not in contenu

    @pytest.mark.parametrize(
        "nom, mdp, role",
        [
            ("", "motdepasse", "admin"),
            ("ab", "motdepasse", "admin"),
            ("victorien", "", "admin"),
            ("victorien", "court", "admin"),
            ("victorien", "motdepasse", "sorcier"),
        ],
    )
    def test_saisies_refusees(self, nom, mdp, role):
        ok, err = ah.create_user(nom, mdp, role)
        assert not ok and err

    def test_doublon_refuse_sans_tenir_compte_de_la_casse(self):
        ah.create_user("victorien", "motdepasse", "admin")
        ok, err = ah.create_user("VICTORIEN", "autrepasse", "agent")
        assert not ok and "existe" in err


# ── Authentification ──────────────────────────────────────────────────────────


class TestAuthentification:
    def test_connexion_reussie(self):
        ah.create_user("victorien", "motdepasse", "admin")
        ok, user, motif = ah.authenticate("victorien", "motdepasse")
        assert ok and motif == "" and user["role"] == "admin"

    def test_le_mot_de_passe_n_est_pas_renvoye(self):
        ah.create_user("victorien", "motdepasse", "admin")
        _, user, _ = ah.authenticate("victorien", "motdepasse")
        assert "password_hash" not in user and "salt" not in user

    def test_mauvais_mot_de_passe(self):
        ah.create_user("victorien", "motdepasse", "admin")
        ok, user, _ = ah.authenticate("victorien", "MAUVAIS")
        assert not ok and user is None

    def test_utilisateur_inconnu(self):
        ok, user, _ = ah.authenticate("fantome", "motdepasse")
        assert not ok and user is None

    def test_la_casse_du_nom_est_ignoree(self):
        ah.create_user("Victorien", "motdepasse", "admin")
        ok, _, _ = ah.authenticate("VICTORIEN", "motdepasse")
        assert ok

    def test_compte_suspendu(self):
        ah.create_user("victorien", "motdepasse", "admin")
        ah.create_user("second", "motdepasse", "admin")
        ah.suspend_user("victorien")
        ok, _, motif = ah.authenticate("victorien", "motdepasse")
        assert not ok and motif == "suspended"

    def test_mot_de_passe_expire_laisse_entrer_avec_un_signal(self):
        ah.create_user("victorien", "motdepasse", "admin")
        entrees = json.load(open(ah.USERS_FILE, encoding="utf-8"))
        entrees[0]["password_changed_at"] = _horodatage(jours_avant=100)
        _ecrire_utilisateurs(entrees)
        ok, user, motif = ah.authenticate("victorien", "motdepasse")
        assert ok and motif == "expired" and user is not None


class TestMigrationDeHachage:
    """Les comptes en SHA-256 doivent basculer en PBKDF2 a la connexion."""

    def _compte_v1(self, nom="ancien", mdp="motdepasse"):
        sel = ah._generate_salt()
        _ecrire_utilisateurs(
            [
                {
                    "username": nom,
                    "password_hash": ah._hash_password_v1(mdp, sel),
                    "salt": sel,
                    "hash_version": 1,
                    "role": "admin",
                    "created_at": _horodatage(),
                    "password_changed_at": _horodatage(),
                }
            ]
        )

    def test_un_compte_v1_peut_se_connecter(self):
        self._compte_v1()
        ok, _, _ = ah.authenticate("ancien", "motdepasse")
        assert ok

    def test_la_connexion_bascule_le_compte_en_v2(self):
        self._compte_v1()
        ah.authenticate("ancien", "motdepasse")
        entree = json.load(open(ah.USERS_FILE, encoding="utf-8"))[0]
        assert entree["hash_version"] == 2

    def test_le_sel_est_renouvele_a_la_migration(self):
        self._compte_v1()
        avant = json.load(open(ah.USERS_FILE, encoding="utf-8"))[0]["salt"]
        ah.authenticate("ancien", "motdepasse")
        apres = json.load(open(ah.USERS_FILE, encoding="utf-8"))[0]["salt"]
        assert avant != apres

    def test_le_compte_migre_se_reconnecte(self):
        self._compte_v1()
        ah.authenticate("ancien", "motdepasse")
        ok, _, _ = ah.authenticate("ancien", "motdepasse")
        assert ok

    def test_un_mauvais_mot_de_passe_ne_migre_pas(self):
        self._compte_v1()
        ah.authenticate("ancien", "MAUVAIS")
        entree = json.load(open(ah.USERS_FILE, encoding="utf-8"))[0]
        assert entree["hash_version"] == 1


# ── Verrouillage anti-force-brute ─────────────────────────────────────────────


class TestVerrouillage:
    def test_pas_de_verrou_avant_le_seuil(self):
        for _ in range(ah._LOCKOUT_MAX_FAILURES - 1):
            ah._record_failed_attempt("victime")
        assert ah.check_lockout("victime")[0] is False

    def test_verrou_au_seuil(self):
        for _ in range(ah._LOCKOUT_MAX_FAILURES):
            ah._record_failed_attempt("victime")
        verrouille, restant = ah.check_lockout("victime")
        assert verrouille and restant > 0

    def test_une_connexion_reussie_efface_le_compteur(self):
        ah.create_user("victorien", "motdepasse", "admin")
        for _ in range(ah._LOCKOUT_MAX_FAILURES - 1):
            ah.authenticate("victorien", "MAUVAIS")
        ah.authenticate("victorien", "motdepasse")
        assert ah.check_lockout("victorien")[0] is False

    def test_authenticate_signale_le_verrou(self):
        ah.create_user("victorien", "motdepasse", "admin")
        for _ in range(ah._LOCKOUT_MAX_FAILURES):
            ah.authenticate("victorien", "MAUVAIS")
        ok, _, motif = ah.authenticate("victorien", "motdepasse")
        assert not ok and motif.startswith("locked:")

    def test_les_tentatives_anciennes_sortent_de_la_fenetre(self):
        vieux = datetime.now() - timedelta(minutes=ah._LOCKOUT_WINDOW_MIN + 1)
        ah._failed_attempts["victime"] = [vieux] * 10
        assert ah.check_lockout("victime")[0] is False


# ── Protection du dernier administrateur ──────────────────────────────────────


class TestDernierAdministrateur:
    def test_suppression_refusee(self):
        ah.create_user("chef", "motdepasse", "admin")
        ok, err = ah.delete_user("chef")
        assert not ok and "dernier administrateur" in err

    def test_suspension_refusee(self):
        ah.create_user("chef", "motdepasse", "admin")
        ok, err = ah.suspend_user("chef")
        assert not ok and "dernier administrateur" in err

    def test_retrogradation_refusee(self):
        ah.create_user("chef", "motdepasse", "admin")
        ok, err = ah.update_user_role("chef", "agent")
        assert not ok and "dernier administrateur" in err

    def test_autorise_des_qu_un_second_admin_existe(self):
        ah.create_user("chef", "motdepasse", "admin")
        ah.create_user("adjoint", "motdepasse", "admin")
        ok, _ = ah.delete_user("chef")
        assert ok

    def test_on_ne_supprime_pas_son_propre_compte(self):
        ah.create_user("chef", "motdepasse", "admin")
        ah.create_user("adjoint", "motdepasse", "admin")
        ah.set_current_user({"username": "chef", "role": "admin"})
        ok, err = ah.delete_user("chef")
        assert not ok and "propre compte" in err


# ── Statut d'acces ────────────────────────────────────────────────────────────


class TestStatutDAcces:
    def test_compte_sain(self):
        assert ah.access_status({"password_changed_at": _horodatage()}) == "active"

    def test_suspension_prime_sur_le_reste(self):
        user = {"suspended": True, "password_changed_at": _horodatage(jours_avant=999)}
        assert ah.access_status(user) == "suspended"

    def test_mot_de_passe_expire(self):
        user = {"password_changed_at": _horodatage(jours_avant=91)}
        assert ah.access_status(user) == "pw_expired"

    def test_acces_expire(self):
        hier = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        user = {"access_expires_at": hier, "password_changed_at": _horodatage()}
        assert ah.access_status(user) == "access_expired"

    def test_acces_sans_date_est_illimite(self):
        assert ah.is_access_expired({"access_expires_at": ""}) is False

    def test_horodatage_de_mot_de_passe_corrompu_echoue_en_securite(self):
        assert ah.is_password_expired({"password_changed_at": "n'importe quoi"}) is True

    def test_jours_restants(self):
        user = {"password_changed_at": _horodatage(jours_avant=1)}
        assert 87 <= ah.password_days_left(user) <= 89


# ── Session ───────────────────────────────────────────────────────────────────


class TestSession:
    def test_role_courant_et_admin(self):
        ah.set_current_user({"username": "chef", "role": "admin"})
        assert ah.current_role() == "admin" and ah.is_admin()

    def test_sans_session(self):
        ah.set_current_user(None)
        assert ah.current_role() == "" and not ah.is_admin()

    def test_un_agent_n_est_pas_admin(self):
        ah.set_current_user({"username": "a", "role": "agent"})
        assert not ah.is_admin()


# ── Robustesse du fichier ─────────────────────────────────────────────────────


class TestFichierAbime:
    def test_fichier_absent(self):
        assert ah._load_users() == [] and not ah.has_users()

    def test_json_illisible(self):
        with open(ah.USERS_FILE, "w", encoding="utf-8") as f:
            f.write("{ pas du json")
        assert ah._load_users() == []

    def test_entrees_incompletes_ecartees(self):
        _ecrire_utilisateurs(
            [
                {"username": "bon", "role": "admin"},
                {"username": "", "role": "admin"},
                {"username": "sansrole"},
                "pas un dictionnaire",
            ]
        )
        assert [u["username"] for u in ah._load_users()] == ["bon"]


# ── Defauts corriges ──────────────────────────────────────────────────────────


class TestEntreeSansSel:
    """Une entree sans sel faisait planter la connexion de tout le monde.

    `_valid_user_entries` ne controlait que le nom et le role. Un users.json
    edite a la main, tronque par une ecriture interrompue ou issu d'une version
    anterieure levait `KeyError: 'salt'` dans `authenticate`, exception non
    rattrapee : plus aucun compte ne pouvait se connecter, y compris les
    entrees saines situees plus loin dans le fichier.
    """

    def test_une_entree_sans_sel_ne_fait_pas_planter(self):
        _ecrire_utilisateurs(
            [{"username": "abime", "role": "admin", "password_hash": "x"}]
        )
        ok, user, _ = ah.authenticate("abime", "peu importe")
        assert not ok and user is None

    def test_une_entree_sans_empreinte_ne_fait_pas_planter(self):
        _ecrire_utilisateurs([{"username": "abime", "role": "admin", "salt": "abcd"}])
        ok, user, _ = ah.authenticate("abime", "peu importe")
        assert not ok and user is None

    def test_une_entree_abimee_n_empeche_pas_les_saines(self):
        sel = ah._generate_salt()
        _ecrire_utilisateurs(
            [
                {"username": "abime", "role": "admin"},
                {
                    "username": "sain",
                    "password_hash": ah._hash_password("motdepasse", sel),
                    "salt": sel,
                    "hash_version": 2,
                    "role": "admin",
                    "created_at": _horodatage(),
                    "password_changed_at": _horodatage(),
                },
            ]
        )
        ok, _, _ = ah.authenticate("sain", "motdepasse")
        assert ok


class TestDateDAccesCorrompue:
    """L'expiration d'acces echouait en sens ouvert.

    `is_password_expired` considere un horodatage illisible comme expire
    (fail-secure). `is_access_expired` faisait l'inverse et accordait l'acces.
    Une date abimee transformait donc un acces a duree limitee en acces
    illimite, sans aucun signal.
    """

    @pytest.mark.parametrize(
        "valeur", ["n'importe quoi", "32/13/2020", "2020-13-45", "0000"]
    )
    def test_une_date_illisible_refuse_l_acces(self, valeur):
        assert ah.is_access_expired({"access_expires_at": valeur}) is True

    def test_le_statut_correspond(self):
        user = {"access_expires_at": "abime", "password_changed_at": _horodatage()}
        assert ah.access_status(user) == "access_expired"

    def test_une_date_valide_dans_le_futur_laisse_passer(self):
        demain = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
        assert ah.is_access_expired({"access_expires_at": demain}) is False


class TestComparaisonATempsConstant:
    """La comparaison d'empreintes doit etre a temps constant.

    `expected != u["password_hash"]` s'arrete au premier octet different : le
    temps de reponse renseigne sur le nombre d'octets devines.
    """

    def test_authenticate_utilise_compare_digest(self):
        import inspect

        assert "compare_digest" in inspect.getsource(ah._empreintes_egales)

    def test_les_empreintes_identiques_sont_egales(self):
        assert ah._empreintes_egales("abcd", "abcd") is True

    def test_les_empreintes_differentes_ne_le_sont_pas(self):
        assert ah._empreintes_egales("abcd", "abce") is False

    def test_une_empreinte_absente_ne_leve_pas(self):
        assert ah._empreintes_egales("abcd", None) is False


class TestVerrouillagePersistant:
    """Le verrouillage ne doit pas disparaitre au redemarrage.

    Le compteur vivait dans un dictionnaire en memoire. Relancer l'application
    remettait a zero les cinq tentatives, ce qui reduisait a rien la protection
    annoncee au README : il suffisait de fermer et rouvrir entre chaque salve.
    """

    def test_le_verrou_survit_a_un_redemarrage(self):
        ah.create_user("victime", "motdepasse", "admin")
        for _ in range(ah._LOCKOUT_MAX_FAILURES):
            ah.authenticate("victime", "MAUVAIS")
        assert ah.check_lockout("victime")[0]

        ah._failed_attempts.clear()  # equivalent d'un redemarrage

        verrouille, restant = ah.check_lockout("victime")
        assert verrouille and restant > 0

    def test_une_connexion_reussie_efface_aussi_l_etat_persistant(self):
        ah.create_user("victime", "motdepasse", "admin")
        for _ in range(ah._LOCKOUT_MAX_FAILURES - 1):
            ah.authenticate("victime", "MAUVAIS")
        ah.authenticate("victime", "motdepasse")

        ah._failed_attempts.clear()
        assert ah.check_lockout("victime")[0] is False

    def test_un_etat_persistant_illisible_ne_fait_pas_planter(self, tmp_path):
        chemin = ah._chemin_des_tentatives()
        os.makedirs(os.path.dirname(chemin), exist_ok=True)
        with open(chemin, "w", encoding="utf-8") as f:
            f.write("{ pas du json")
        assert ah.check_lockout("qui que ce soit")[0] is False
