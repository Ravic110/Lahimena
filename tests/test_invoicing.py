"""Tests for invoice calculations."""

from utils.excel_handler import (
    INVOICE_STATUS_PAID,
    INVOICE_STATUS_PARTIAL,
    INVOICE_STATUS_UNPAID,
    calculate_invoice_totals,
)


def test_invoice_totals_with_margin_and_vat():
    result = calculate_invoice_totals(
        montant_ht=1000,
        cout_ht=700,
        marge_pct=10,
        tva_pct=20,
        acompte=100,
        statut=INVOICE_STATUS_PARTIAL,
    )

    assert result["Montant_HT"] == 1000
    assert result["Marge_Montant"] == 100
    assert result["Base_Taxable_HT"] == 1100
    assert result["TVA_Montant"] == 220
    assert result["Total_TTC"] == 1320
    assert result["Acompte"] == 100
    assert result["Reste_A_Payer"] == 1220
    assert result["Statut"] == INVOICE_STATUS_PARTIAL


def test_status_paid_sets_remaining_to_zero():
    result = calculate_invoice_totals(
        montant_ht=500,
        marge_pct=0,
        tva_pct=0,
        acompte=20,
        statut=INVOICE_STATUS_PAID,
    )

    assert result["Total_TTC"] == 500
    assert result["Acompte"] == 500
    assert result["Reste_A_Payer"] == 0
    assert result["Statut"] == INVOICE_STATUS_PAID


def test_status_unpaid_forces_zero_deposit():
    result = calculate_invoice_totals(
        montant_ht=800,
        marge_pct=0,
        tva_pct=0,
        acompte=120,
        statut=INVOICE_STATUS_UNPAID,
    )

    assert result["Total_TTC"] == 800
    assert result["Acompte"] == 0
    assert result["Reste_A_Payer"] == 800
    assert result["Statut"] == INVOICE_STATUS_UNPAID


class TestMargeExplicitementNulle:
    """Une marge de 0 % demandee doit rester nulle.

    Le repli historique -- `si marge_montant == 0 et cout_ht > 0, alors
    marge = montant_ht - cout_ht` -- ne distinguait pas "marge non renseignee"
    de "marge volontairement nulle". Facturer a prix coutant etait donc
    inexprimable des lors que le cout etait connu : sur 1000 HT pour 700 de
    cout, le client se voyait facturer 1300.
    """

    def test_marge_zero_explicite_est_respectee(self):
        result = calculate_invoice_totals(
            montant_ht=1000, cout_ht=700, marge_pct=0, tva_pct=0
        )
        assert result["Marge_Montant"] == 0
        assert result["Total_TTC"] == 1000

    def test_marge_zero_en_chaine_est_respectee(self):
        """L'interface transmet le contenu brut du champ de saisie."""
        result = calculate_invoice_totals(
            montant_ht=1000, cout_ht=700, marge_pct="0", tva_pct=0
        )
        assert result["Marge_Montant"] == 0
        assert result["Total_TTC"] == 1000

    def test_marge_absente_reste_deduite_du_cout(self):
        """Comportement historique conserve quand rien n'est renseigne."""
        result = calculate_invoice_totals(montant_ht=1000, cout_ht=700, tva_pct=0)
        assert result["Marge_Montant"] == 300
        assert result["Total_TTC"] == 1300

    def test_champ_vide_vaut_non_renseigne(self):
        result = calculate_invoice_totals(
            montant_ht=1000, cout_ht=700, marge_pct="", tva_pct=0
        )
        assert result["Marge_Montant"] == 300

    def test_champ_blanc_vaut_non_renseigne(self):
        result = calculate_invoice_totals(
            montant_ht=1000, cout_ht=700, marge_pct="   ", tva_pct=0
        )
        assert result["Marge_Montant"] == 300

    def test_none_vaut_non_renseigne(self):
        result = calculate_invoice_totals(
            montant_ht=1000, cout_ht=700, marge_pct=None, tva_pct=0
        )
        assert result["Marge_Montant"] == 300

    def test_marge_positive_prime_toujours_sur_le_cout(self):
        result = calculate_invoice_totals(
            montant_ht=1000, cout_ht=700, marge_pct=10, tva_pct=0
        )
        assert result["Marge_Montant"] == 100
        assert result["Total_TTC"] == 1100

    def test_marge_zero_sans_cout_reste_nulle(self):
        result = calculate_invoice_totals(montant_ht=1000, marge_pct=0, tva_pct=0)
        assert result["Marge_Montant"] == 0
        assert result["Total_TTC"] == 1000

    def test_tva_s_applique_a_la_base_sans_marge(self):
        result = calculate_invoice_totals(
            montant_ht=1000, cout_ht=700, marge_pct=0, tva_pct=20
        )
        assert result["Base_Taxable_HT"] == 1000
        assert result["TVA_Montant"] == 200
        assert result["Total_TTC"] == 1200
