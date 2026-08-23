"""Filet de securite des cinq ecrans de cotation client.

`gui/` etait a 0 % de couverture. Ces cinq classes -- 4 500 lignes au squelette
identique -- sont la cible d'une factorisation ; les deplacer sans filet serait
imprudent. Ces tests ne decrivent pas un comportement souhaitable, ils
constatent l'existant : l'ecran s'instancie, sa table porte les colonnes
declarees, les lignes sauvegardees sont restituees, et un echec d'ecriture
remonte bien sa sentinelle.

La couche de donnees est entierement neutralisee : aucun classeur n'est lu.
Les tests s'ignorent d'eux-memes la ou aucun affichage n'est disponible.
"""

import importlib

import pytest

from gui.forms import cotation_screen

pytestmark = pytest.mark.gui


# (module, classe, fonctions excel_handler a neutraliser, chargeur, enregistreur)
ECRANS = [
    (
        "client_hotel_cotation",
        "ClientHotelCotation",
        ["load_all_hotels"],
        "load_client_hotel_cotation",
        "save_client_hotel_cotation_to_excel",
    ),
    (
        "client_restauration_cotation",
        "ClientRestaurationCotation",
        ["load_all_hotels", "load_client_hotel_cotation"],
        "load_client_restauration_cotation",
        "save_client_restauration_cotation_to_excel",
    ),
    (
        "client_transport_cotation",
        "ClientTransportCotation",
        [
            "get_km_mada_reperes",
            "get_transport_prestataires",
            "get_transport_vehicle_types",
        ],
        "load_client_transport_cotation",
        "save_client_transport_cotation_to_excel",
    ),
    (
        "client_air_ticket_cotation",
        "ClientAirTicketCotation",
        ["get_avion_compagnies"],
        "load_client_air_ticket_cotation",
        "save_client_air_ticket_cotation_to_excel",
    ),
    (
        "client_collective_cotation",
        "ClientCollectiveCotation",
        ["get_collective_expense_prestataires"],
        "load_client_collective_cotation",
        "save_client_collective_cotation_to_excel",
    ),
]

IDS = [e[0].replace("client_", "").replace("_cotation", "") for e in ECRANS]


@pytest.fixture(params=ECRANS, ids=IDS)
def ecran(request, monkeypatch):
    """Un ecran de cotation, sa couche de donnees remplacee par du vide."""
    nom_module, nom_classe, listes_vides, chargeur, enregistreur = request.param
    module = importlib.import_module(f"gui.forms.{nom_module}")

    for fonction in listes_vides:
        monkeypatch.setattr(module, fonction, lambda *a, **k: [])
    monkeypatch.setattr(module, chargeur, lambda *a, **k: [])

    appels = []

    def _enregistrer(*args, **kwargs):
        appels.append((args, kwargs))
        return 1

    monkeypatch.setattr(module, enregistreur, _enregistrer)

    # Sans cela, la premiere boite de dialogue attend un clic et la suite gele.
    alertes = []
    for niveau in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(
            module.messagebox,
            niveau,
            lambda titre, message, _n=niveau, **k: alertes.append((_n, titre, message)),
        )
    # askyesno attend une reponse : sans defaut, la suppression gele la suite.
    monkeypatch.setattr(cotation_screen.messagebox, "askyesno", lambda *a, **k: True)

    return {
        "alertes": alertes,
        "module": module,
        "classe": getattr(module, nom_classe),
        "chargeur": chargeur,
        "enregistreur": enregistreur,
        "appels": appels,
    }


class TestInstanciation:
    def test_l_ecran_se_construit(self, ecran, tk_root, client_type):
        vue = ecran["classe"](tk_root, client_type)
        assert vue is not None

    def test_la_table_porte_les_colonnes_declarees(self, ecran, tk_root, client_type):
        vue = ecran["classe"](tk_root, client_type)
        attendues = [cle for cle, *_ in vue._COLS]
        assert list(vue._tree["columns"]) == attendues

    def test_chaque_colonne_a_un_intitule(self, ecran, tk_root, client_type):
        vue = ecran["classe"](tk_root, client_type)
        for cle, libelle, *_ in vue._COLS:
            assert vue._tree.heading(cle)["text"] == libelle

    def test_le_client_est_conserve(self, ecran, tk_root, client_type):
        vue = ecran["classe"](tk_root, client_type)
        assert vue.client is client_type


class TestRestitutionDesLignes:
    def test_les_lignes_enregistrees_reapparaissent(
        self, ecran, tk_root, client_type, monkeypatch
    ):
        """Ce qui a ete sauvegarde doit etre reaffiche a la reouverture."""
        vue_vierge = ecran["classe"](tk_root, client_type)
        gabarit = list(vue_vierge._rows)
        if not gabarit:
            pytest.skip("cet ecran ne pre-remplit aucune ligne")

        monkeypatch.setattr(
            ecran["module"], ecran["chargeur"], lambda *a, **k: list(gabarit)
        )
        vue = ecran["classe"](tk_root, client_type)

        assert len(vue._rows) == len(gabarit)
        assert len(vue._tree.get_children()) == len(gabarit)

    def test_le_rafraichissement_est_idempotent(self, ecran, tk_root, client_type):
        """Rafraichir deux fois ne doit pas dupliquer les lignes affichees."""
        vue = ecran["classe"](tk_root, client_type)
        vue._refresh_tree()
        avant = len(vue._tree.get_children())
        vue._refresh_tree()
        assert len(vue._tree.get_children()) == avant

    def test_les_totaux_se_recalculent_sans_erreur(self, ecran, tk_root, client_type):
        vue = ecran["classe"](tk_root, client_type)
        vue._refresh_totals()


class TestEnregistrement:
    """`_save_to_excel` est rigoureusement identique dans les cinq ecrans :
    seul le nom de la fonction de la couche de donnees change."""

    @staticmethod
    def _garnir(vue):
        """Assure qu'il y a au moins une ligne a enregistrer.

        Trois ecrans pre-remplissent depuis la fiche client ; les deux autres
        (charges collectives, avion sans trajet) partent d'un tableau vide.
        """
        if not vue._rows:
            vue._rows = [{"designation": "ligne de test", "total": 1000}]
        return vue

    def test_enregistrer_appelle_la_couche_de_donnees(
        self, ecran, tk_root, client_type
    ):
        vue = self._garnir(ecran["classe"](tk_root, client_type))
        vue._save_to_excel()
        assert len(ecran["appels"]) == 1

    def test_le_client_et_les_lignes_sont_transmis(self, ecran, tk_root, client_type):
        vue = self._garnir(ecran["classe"](tk_root, client_type))
        vue._save_to_excel()
        args, _ = ecran["appels"][0]
        assert args[0] is client_type
        assert args[1] == vue._rows

    def test_le_succes_est_confirme_a_l_utilisateur(self, ecran, tk_root, client_type):
        vue = self._garnir(ecran["classe"](tk_root, client_type))
        vue._save_to_excel()
        assert any(niveau == "showinfo" for niveau, _, _ in ecran["alertes"])

    def test_tableau_vide_le_comportement_diverge_selon_l_ecran(
        self, ecran, tk_root, client_type
    ):
        """Constat, non prescription : les cinq ecrans ne s'accordent pas.

        Quatre refusent d'enregistrer un tableau vide (`if not self._rows`).
        `air_ticket` passe par `_validate_rows`, qui ne dit rien d'une liste
        vide : il enregistre. Or `save_client_air_ticket_cotation_to_excel`
        supprime d'abord les lignes du client -- enregistrer une table vide
        efface donc ses billets sans avertissement.

        C'est peut-etre le geste voulu pour vider une cotation. Le test fige
        l'ecart pour qu'une factorisation ne le supprime pas par megarde.
        """
        vue = ecran["classe"](tk_root, client_type)
        vue._rows = []
        ecran["alertes"].clear()
        vue._save_to_excel()

        if ecran["module"].__name__.endswith("client_air_ticket_cotation"):
            assert len(ecran["appels"]) == 1
            args, _ = ecran["appels"][0]
            assert args[1] == []
        else:
            assert ecran["appels"] == []
            assert any(niveau == "showwarning" for niveau, _, _ in ecran["alertes"])

    @pytest.mark.parametrize("sentinelle", [-1, -2])
    def test_un_echec_d_ecriture_est_signale_sans_lever(
        self, ecran, tk_root, client_type, monkeypatch, sentinelle
    ):
        """L'echec remonte a l'utilisateur, pas en exception.

        -1 : erreur d'ecriture. -2 : classeur ouvert dans Excel.
        """
        monkeypatch.setattr(
            ecran["module"], ecran["enregistreur"], lambda *a, **k: sentinelle
        )
        vue = self._garnir(ecran["classe"](tk_root, client_type))
        ecran["alertes"].clear()
        vue._save_to_excel()

        assert any(niveau == "showerror" for niveau, _, _ in ecran["alertes"])
        assert not any(niveau == "showinfo" for niveau, _, _ in ecran["alertes"])

    def test_le_verrouillage_nomme_le_fichier_a_fermer(
        self, ecran, tk_root, client_type, monkeypatch
    ):
        monkeypatch.setattr(ecran["module"], ecran["enregistreur"], lambda *a, **k: -2)
        vue = self._garnir(ecran["classe"](tk_root, client_type))
        ecran["alertes"].clear()
        vue._save_to_excel()

        messages = [message for niveau, _, message in ecran["alertes"]]
        assert any("data.xlsx" in m for m in messages)


class TestSuppressionEtModification:
    """Methodes remontees dans ClientCotationScreen.

    `_delete_selected` existait en deux versions : quatre ecrans demandaient
    confirmation, `air_ticket` supprimait au premier clic. La factorisation
    aligne les cinq sur la confirmation -- une ligne de cotation effacee par
    megarde est une saisie a refaire.
    """

    @staticmethod
    def _garnir(vue, module, n=2):
        """Place `n` lignes affichables dans le tableau.

        Des dicts inventes ne suffisent pas : `_refresh_tree` lit des champs
        propres a chaque entite (`room_prices` pour l'hotellerie...). On part
        donc des lignes que l'ecran sait produire lui-meme.
        """
        import copy

        if vue._rows:
            gabarit = vue._rows[0]
        else:
            gabarit = module._make_row()

        vue._rows = [copy.deepcopy(gabarit) for _ in range(n)]
        for index, ligne in enumerate(vue._rows):
            ligne["_repere_test"] = index
        vue._refresh_tree()
        return vue

    def test_supprimer_sans_selection_avertit(self, ecran, tk_root, client_type):
        vue = self._garnir(ecran["classe"](tk_root, client_type), ecran["module"])
        ecran["alertes"].clear()
        vue._delete_selected()

        assert len(vue._rows) == 2
        assert any(niveau == "showwarning" for niveau, _, _ in ecran["alertes"])

    def test_supprimer_demande_confirmation(
        self, ecran, tk_root, client_type, monkeypatch
    ):
        refus = []
        monkeypatch.setattr(
            cotation_screen.messagebox,
            "askyesno",
            lambda *a, **k: refus.append(1) or False,
        )
        vue = self._garnir(ecran["classe"](tk_root, client_type), ecran["module"])
        vue._tree.selection_set("0")
        vue._delete_selected()

        assert len(refus) == 1
        assert len(vue._rows) == 2  # refus respecte

    def test_supprimer_retire_la_ligne_choisie(self, ecran, tk_root, client_type):
        vue = self._garnir(ecran["classe"](tk_root, client_type), ecran["module"])
        vue._tree.selection_set("0")
        vue._delete_selected()

        assert len(vue._rows) == 1
        assert vue._rows[0]["_repere_test"] == 1  # la seconde ligne a survecu
        assert len(vue._tree.get_children()) == 1

    def test_modifier_sans_selection_avertit(self, ecran, tk_root, client_type):
        vue = self._garnir(ecran["classe"](tk_root, client_type), ecran["module"])
        if not hasattr(vue, "_edit_selected"):
            pytest.skip("cet ecran n'expose pas _edit_selected")
        ecran["alertes"].clear()
        vue._edit_selected()

        assert any(niveau == "showwarning" for niveau, _, _ in ecran["alertes"])

    def test_modifier_ouvre_le_dialogue_sur_la_bonne_ligne(
        self, ecran, tk_root, client_type, monkeypatch
    ):
        vue = self._garnir(ecran["classe"](tk_root, client_type), ecran["module"])
        if not hasattr(vue, "_edit_selected"):
            pytest.skip("cet ecran n'expose pas _edit_selected")

        recus = []
        monkeypatch.setattr(
            vue, "_open_row_dialog", lambda *a, **k: recus.append((a, k))
        )
        vue._tree.selection_set("1")
        vue._edit_selected()

        assert len(recus) == 1
        args, kwargs = recus[0]
        assert kwargs.get("row_index") == 1
        assert args[0] is vue._rows[1]
