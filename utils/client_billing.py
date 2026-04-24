"""Client quote/invoice aggregation helpers."""

from copy import deepcopy


CATEGORY_HOTEL = "Hébergement"
CATEGORY_RESTAURATION = "Restauration"
CATEGORY_TRANSPORT = "Transport"
CATEGORY_AIR_TICKET = "Billets d’avion"
CATEGORY_VISIT = "Visites/Excursions"
CATEGORY_COLLECTIVE = "Charges collectives"


def _to_float(value, default=0.0):
    try:
        text = str(value).replace(" ", "").replace(",", ".").strip()
        return float(text or default)
    except (TypeError, ValueError):
        return default


def _to_int(value, default=0):
    try:
        return int(_to_float(value, default))
    except (TypeError, ValueError):
        return default


def _safe_strip(value):
    return str(value or "").strip()


def _normalize_text(value):
    return " ".join(_safe_strip(value).lower().replace("_", " ").split())


def _client_matches_generic_row(client, row):
    client_ref = _normalize_text(client.get("ref_client"))
    client_nom = _normalize_text(client.get("nom"))
    client_prenom = _normalize_text(client.get("prenom"))

    row_id = _normalize_text(
        row.get("ID_CLIENT")
        or row.get("ID_Client")
        or row.get("client_id")
        or row.get("Référence")
        or row.get("reference")
    )
    if client_ref and row_id:
        return row_id == client_ref

    row_nom = _normalize_text(row.get("Nom") or row.get("Nom_Client") or row.get("nom"))
    row_prenom = _normalize_text(
        row.get("Prénom") or row.get("Prénom_Client") or row.get("prenom")
    )
    if client_nom and row_nom and row_nom != client_nom:
        return False
    if client_prenom and row_prenom and row_prenom != client_prenom:
        return False
    return bool(row_nom or row_prenom)


def _line(
    category,
    designation,
    quantity,
    cost_unit,
    cost_total,
    margin_pct=0.0,
    margin_editable=True,
    source_module="",
    currency="Ariary",
):
    quantity = max(1, _to_int(quantity, 1))
    cost_unit = max(0.0, _to_float(cost_unit))
    cost_total = max(0.0, _to_float(cost_total))
    margin_pct = max(0.0, _to_float(margin_pct))

    base = {
        "category": category,
        "designation": designation,
        "quantity": quantity,
        "unit": "unité",
        "cost_unit": cost_unit if cost_unit > 0 else (cost_total / quantity if quantity else 0.0),
        "cost_total": cost_total,
        "margin_pct": margin_pct,
        "margin_editable": bool(margin_editable),
        "source_module": source_module,
        "currency": currency or "Ariary",
    }
    return apply_margin_to_quote_line(base, margin_pct)


def apply_margin_to_quote_line(line, margin_pct):
    """Reprice one quote line while enforcing the restauration rule."""
    updated = deepcopy(line)
    quantity = max(1, _to_int(updated.get("quantity", 1), 1))
    cost_total = max(0.0, _to_float(updated.get("cost_total", 0)))

    if updated.get("category") == CATEGORY_RESTAURATION or not updated.get(
        "margin_editable", True
    ):
        margin_pct_value = 0.0
    else:
        margin_pct_value = max(0.0, _to_float(margin_pct))

    margin_amount = cost_total * (margin_pct_value / 100.0)
    total_price = cost_total + margin_amount
    unit_price = total_price / quantity if quantity else total_price

    updated["quantity"] = quantity
    updated["margin_pct"] = margin_pct_value
    updated["margin_amount"] = margin_amount
    updated["total_price"] = total_price
    updated["unit_price"] = unit_price
    return updated


def _default_source_rows(client):
    from utils.excel_handler import (
        load_all_visite_excursion_quotations,
        load_client_air_ticket_cotation,
        load_client_collective_cotation,
        load_client_hotel_cotation,
        load_client_restauration_cotation,
        load_client_transport_cotation,
    )

    visite_rows = [
        row
        for row in load_all_visite_excursion_quotations()
        if _client_matches_generic_row(client, row)
    ]
    return {
        "hotel": load_client_hotel_cotation(client),
        "restauration": load_client_restauration_cotation(client),
        "transport": load_client_transport_cotation(client),
        "air_ticket": load_client_air_ticket_cotation(client),
        "visite_excursion": visite_rows,
        "collective": load_client_collective_cotation(client),
    }


def build_client_quote(client, source_rows=None):
    """Build a normalized active quote for one client."""
    rows = source_rows or _default_source_rows(client)
    lines = []

    for row in rows.get("hotel", []):
        lines.append(
            _line(
                CATEGORY_HOTEL,
                f"{_safe_strip(row.get('hotel'))} - {_safe_strip(row.get('ville'))}".strip(" -"),
                row.get("nuits", 1),
                row.get("prix_unitaire", 0),
                row.get("depense", 0),
                row.get("marge", 0),
                True,
                "hotel",
            )
        )

    for row in rows.get("restauration", []):
        total = _to_float(row.get("total", 0))
        lines.append(
            _line(
                CATEGORY_RESTAURATION,
                f"{_safe_strip(row.get('hotel'))} - {_safe_strip(row.get('forfait'))}".strip(" -"),
                row.get("nuits", 1),
                row.get("prix_unitaire", 0),
                total,
                0,
                False,
                "restauration",
            )
        )

    for row in rows.get("transport", []):
        total = _to_float(row.get("total", 0))
        lines.append(
            _line(
                CATEGORY_TRANSPORT,
                f"{_safe_strip(row.get('depart'))} -> {_safe_strip(row.get('arrivee'))} - {_safe_strip(row.get('type_voiture'))}".strip(" -"),
                row.get("nb_vehicules", 1),
                total,
                total,
                0,
                True,
                "transport",
            )
        )

    for row in rows.get("air_ticket", []):
        cost_total = _to_float(row.get("sous_total", 0))
        if cost_total <= 0:
            cost_total = _to_float(row.get("montant_adultes", 0)) + _to_float(
                row.get("montant_enfants", 0)
            )
        lines.append(
            _line(
                CATEGORY_AIR_TICKET,
                f"{_safe_strip(row.get('compagnie'))} - {_safe_strip(row.get('ville_depart'))} -> {_safe_strip(row.get('ville_arrivee'))}".strip(" -"),
                1,
                cost_total,
                cost_total,
                row.get("marge_pct", 0),
                True,
                "air_ticket",
            )
        )

    for row in rows.get("visite_excursion", []):
        quantity = row.get("Quantité", 1)
        cost_unit = row.get("Montant", 0)
        cost_total = row.get("Total", 0) or (_to_float(cost_unit) * max(1, _to_int(quantity, 1)))
        lines.append(
            _line(
                CATEGORY_VISIT,
                f"{_safe_strip(row.get('Prestation'))} - {_safe_strip(row.get('Désignation'))}".strip(" -"),
                quantity,
                cost_unit,
                cost_total,
                0,
                True,
                "visite_excursion",
            )
        )

    for row in rows.get("collective", []):
        lines.append(
            _line(
                CATEGORY_COLLECTIVE,
                f"{_safe_strip(row.get('prestataire'))} - {_safe_strip(row.get('designation'))}".strip(" -"),
                row.get("quantite", 1),
                row.get("prix_unitaire", 0),
                row.get("depense", 0),
                row.get("marge", 0),
                True,
                "collective",
            )
        )

    lines = [line for line in lines if _to_float(line.get("total_price", 0)) > 0]
    total_cost = sum(_to_float(line.get("cost_total", 0)) for line in lines)
    total_price = sum(_to_float(line.get("total_price", 0)) for line in lines)
    total_margin = sum(_to_float(line.get("margin_amount", 0)) for line in lines)

    return {
        "client_id": _safe_strip(client.get("ref_client")),
        "client_name": f"{_safe_strip(client.get('prenom'))} {_safe_strip(client.get('nom'))}".strip(),
        "numero_dossier": _safe_strip(client.get("numero_dossier")),
        "currency": "Ariary",
        "lines": lines,
        "line_count": len(lines),
        "total_cost": total_cost,
        "total_margin": total_margin,
        "total_price": total_price,
    }


def convert_quote_to_invoice(quote):
    """Convert a quote into an invoice while preserving detailed service lines."""
    lines = []
    for line in quote.get("lines", []):
        category = _safe_strip(line.get("category"))
        if not category:
            continue

        quantity = max(1, _to_int(line.get("quantity", 1), 1))
        unit_price = max(
            0.0,
            _to_float(line.get("unit_price", line.get("total_price", 0))),
        )
        total_price = max(
            0.0,
            _to_float(line.get("total_price", unit_price * quantity)),
        )
        lines.append(
            {
                "category": category,
                "designation": _safe_strip(line.get("designation")) or category,
                "quantity": quantity,
                "unit": _safe_strip(line.get("unit")) or "unité",
                "unit_price": unit_price,
                "total_price": total_price,
                "currency": line.get("currency", quote.get("currency", "Ariary")),
            }
        )
    total_price = sum(_to_float(line.get("total_price", 0)) for line in lines)
    return {
        "client_id": quote.get("client_id", ""),
        "client_name": quote.get("client_name", ""),
        "numero_dossier": quote.get("numero_dossier", ""),
        "currency": quote.get("currency", "Ariary"),
        "lines": lines,
        "line_count": len(lines),
        "total_price": total_price,
    }


def invoice_requires_detail_refresh(invoice_document, quote_document):
    """Detect legacy grouped invoices that should be rebuilt from the detailed quote."""
    invoice_lines = invoice_document.get("lines", [])
    quote_lines = quote_document.get("lines", [])
    if not invoice_lines or not quote_lines:
        return False
    if len(invoice_lines) >= len(quote_lines):
        return False

    for line in invoice_lines:
        category = _safe_strip(line.get("category"))
        designation = _safe_strip(line.get("designation"))
        if not category or designation != category:
            return False
    return True
