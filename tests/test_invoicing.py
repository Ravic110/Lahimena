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
