"""Tests for the client quote/invoice aggregation service."""

from utils.client_billing import (
    CATEGORY_AIR_TICKET,
    CATEGORY_COLLECTIVE,
    CATEGORY_HOTEL,
    CATEGORY_RESTAURATION,
    CATEGORY_TRANSPORT,
    CATEGORY_VISIT,
    apply_margin_to_quote_line,
    build_client_quote,
    convert_quote_to_invoice,
)


def _client():
    return {
        "ref_client": "CLI001",
        "numero_dossier": "DOS001",
        "nom": "Rakoto",
        "prenom": "Aina",
    }


def test_build_client_quote_aggregates_all_categories():
    quote = build_client_quote(
        _client(),
        source_rows={
            "hotel": [
                {
                    "ville": "Antananarivo",
                    "hotel": "Colbert",
                    "nuits": "2",
                    "prix_unitaire": 100,
                    "depense": 200,
                    "marge": "15",
                    "total": 230,
                }
            ],
            "restauration": [
                {
                    "ville": "Antananarivo",
                    "hotel": "Colbert",
                    "nuits": "2",
                    "forfait": "Demi-pension",
                    "prix_unitaire": 50,
                    "total": 100,
                }
            ],
            "transport": [
                {
                    "depart": "Antananarivo",
                    "arrivee": "Morondava",
                    "type_voiture": "4x4",
                    "nb_vehicules": "1",
                    "total": 300,
                }
            ],
            "air_ticket": [
                {
                    "type_trajet": "aller",
                    "compagnie": "Air Austral",
                    "ville_depart": "TNR",
                    "ville_arrivee": "NOS",
                    "nb_adultes": "2",
                    "sous_total": 400,
                    "marge_pct": "20",
                    "total": 480,
                }
            ],
            "visite_excursion": [
                {
                    "ID_CLIENT": "CLI001",
                    "Nom": "Rakoto",
                    "Prénom": "Aina",
                    "Prestation": "Guide",
                    "Désignation": "Baobab",
                    "Quantité": "2",
                    "Montant": "30",
                    "Total": "60",
                }
            ],
            "collective": [
                {
                    "prestataire": "Guide",
                    "designation": "Accueil",
                    "quantite": "2",
                    "prix_unitaire": "20",
                    "marge": "10",
                    "depense": 40,
                    "total": 44,
                }
            ],
        },
    )

    categories = {line["category"] for line in quote["lines"]}
    assert categories == {
        CATEGORY_HOTEL,
        CATEGORY_RESTAURATION,
        CATEGORY_TRANSPORT,
        CATEGORY_AIR_TICKET,
        CATEGORY_VISIT,
        CATEGORY_COLLECTIVE,
    }
    assert quote["line_count"] == 6
    assert quote["total_cost"] == 1100
    assert quote["total_price"] == 1214

    restauration_line = next(
        line for line in quote["lines"] if line["category"] == CATEGORY_RESTAURATION
    )
    assert restauration_line["margin_pct"] == 0
    assert restauration_line["margin_editable"] is False
    assert restauration_line["total_price"] == 100


def test_apply_margin_to_quote_line_keeps_restauration_at_zero():
    line = {
        "category": CATEGORY_RESTAURATION,
        "quantity": 2,
        "cost_unit": 50,
        "cost_total": 100,
        "margin_pct": 0,
        "margin_editable": False,
    }

    updated = apply_margin_to_quote_line(line, 25)

    assert updated["margin_pct"] == 0
    assert updated["margin_amount"] == 0
    assert updated["total_price"] == 100
    assert updated["unit_price"] == 50


def test_convert_quote_to_invoice_groups_lines_by_category():
    quote = {
        "client_id": "CLI001",
        "client_name": "Aina Rakoto",
        "numero_dossier": "DOS001",
        "currency": "Ariary",
        "lines": [
            {
                "category": CATEGORY_HOTEL,
                "designation": "Hotel A",
                "quantity": 2,
                "total_price": 230,
            },
            {
                "category": CATEGORY_HOTEL,
                "designation": "Hotel B",
                "quantity": 1,
                "total_price": 115,
            },
            {
                "category": CATEGORY_RESTAURATION,
                "designation": "Restauration",
                "quantity": 1,
                "total_price": 100,
            },
        ],
    }

    invoice = convert_quote_to_invoice(quote)

    assert invoice["line_count"] == 2
    assert invoice["total_price"] == 445

    hotel_line = next(
        line for line in invoice["lines"] if line["category"] == CATEGORY_HOTEL
    )
    assert hotel_line["designation"] == CATEGORY_HOTEL
    assert hotel_line["quantity"] == 1
    assert hotel_line["unit_price"] == 345
    assert "margin_pct" not in hotel_line
