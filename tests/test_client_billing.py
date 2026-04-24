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
    invoice_requires_detail_refresh,
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


def test_convert_quote_to_invoice_preserves_detailed_lines():
    quote = {
        "client_id": "CLI001",
        "client_name": "Aina Rakoto",
        "numero_dossier": "DOS001",
        "currency": "Ariary",
        "lines": [
            {
                "category": CATEGORY_HOTEL,
                "designation": "Colbert - Antananarivo",
                "quantity": 2,
                "unit": "nuit",
                "unit_price": 115,
                "total_price": 230,
                "currency": "Ariary",
            },
            {
                "category": CATEGORY_TRANSPORT,
                "designation": "Antananarivo -> Morondava - 4x4",
                "quantity": 1,
                "unit": "trajet",
                "unit_price": 300,
                "total_price": 300,
                "currency": "Ariary",
            },
        ],
    }

    invoice = convert_quote_to_invoice(quote)

    assert invoice["line_count"] == 2
    assert invoice["total_price"] == 530
    assert invoice["lines"][0]["designation"] == "Colbert - Antananarivo"
    assert invoice["lines"][0]["quantity"] == 2
    assert invoice["lines"][0]["unit_price"] == 115
    assert invoice["lines"][1]["designation"] == "Antananarivo -> Morondava - 4x4"
    assert invoice["lines"][1]["quantity"] == 1
    assert invoice["lines"][1]["unit_price"] == 300


def test_convert_quote_to_invoice_falls_back_to_category_for_empty_designation():
    quote = {
        "client_id": "CLI001",
        "client_name": "Aina Rakoto",
        "numero_dossier": "DOS001",
        "currency": "Ariary",
        "lines": [
            {
                "category": CATEGORY_HOTEL,
                "designation": "",
                "quantity": 1,
                "unit_price": 230,
                "total_price": 230,
            }
        ],
    }

    invoice = convert_quote_to_invoice(quote)

    assert invoice["line_count"] == 1
    assert invoice["lines"][0]["designation"] == CATEGORY_HOTEL


def test_invoice_requires_detail_refresh_for_legacy_grouped_invoice():
    invoice = {
        "lines": [
            {
                "category": CATEGORY_HOTEL,
                "designation": CATEGORY_HOTEL,
                "quantity": 1,
                "unit_price": 345,
                "total_price": 345,
            }
        ]
    }
    quote = {
        "lines": [
            {
                "category": CATEGORY_HOTEL,
                "designation": "Colbert - Antananarivo",
                "quantity": 2,
                "unit_price": 115,
                "total_price": 230,
            },
            {
                "category": CATEGORY_HOTEL,
                "designation": "Carlton - Antananarivo",
                "quantity": 1,
                "unit_price": 115,
                "total_price": 115,
            },
        ]
    }

    assert invoice_requires_detail_refresh(invoice, quote) is True


def test_invoice_requires_detail_refresh_ignores_already_detailed_invoice():
    invoice = {
        "lines": [
            {
                "category": CATEGORY_HOTEL,
                "designation": "Colbert - Antananarivo",
                "quantity": 2,
                "unit_price": 115,
                "total_price": 230,
            }
        ]
    }
    quote = {
        "lines": [
            {
                "category": CATEGORY_HOTEL,
                "designation": "Colbert - Antananarivo",
                "quantity": 2,
                "unit_price": 115,
                "total_price": 230,
            }
        ]
    }

    assert invoice_requires_detail_refresh(invoice, quote) is False
