"""
PDF generation utilities for quotations
Uses ReportLab to create professional PDF documents
"""

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm, inch
    from reportlab.platypus import (
        Image as RLImage,
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

import os
from datetime import datetime

from config import (
    COMPANY_NAME,
    COMPANY_TAGLINE,
    DEVIS_FOLDER,
    LOGO_PATH,
    PDF_FOOTER_TEXT,
)
from utils.logger import logger


class QuotationPDF:
    """Generate professional PDF quotations"""

    def __init__(self, filename=None):
        """
        Initialize PDF generator

        Args:
            filename: Output filename (optional, auto-generated if not provided)
        """
        if not REPORTLAB_AVAILABLE:
            raise ImportError(
                "ReportLab is required for PDF generation. Install with: pip install reportlab"
            )

        # Generate filename if not provided
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"DEVIS_{timestamp}.pdf"

        # Determine full output path
        if os.path.isabs(filename) or os.path.dirname(filename):
            self.filepath = filename
        else:
            self.filepath = os.path.join(DEVIS_FOLDER, filename)

        output_dir = os.path.dirname(self.filepath) or DEVIS_FOLDER
        os.makedirs(output_dir, exist_ok=True)
        self.doc = SimpleDocTemplate(
            self.filepath,
            pagesize=A4,
            rightMargin=2 * cm,
            leftMargin=2 * cm,
            topMargin=2 * cm,
            bottomMargin=2 * cm,
        )

        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
        self.elements = []

        logger.info(f"PDF generator initialized for: {filename}")

    def _setup_custom_styles(self):
        """Setup custom paragraph styles"""
        # Title style
        self.styles.add(
            ParagraphStyle(
                name="CustomTitle",
                parent=self.styles["Heading1"],
                fontSize=24,
                textColor=colors.HexColor("#1E3A5F"),
                spaceAfter=30,
                alignment=TA_CENTER,
                fontName="Helvetica-Bold",
            )
        )

        # Subtitle style
        self.styles.add(
            ParagraphStyle(
                name="CustomSubtitle",
                parent=self.styles["Normal"],
                fontSize=11,
                textColor=colors.HexColor("#666666"),
                spaceAfter=12,
                alignment=TA_CENTER,
            )
        )

        # Header style
        self.styles.add(
            ParagraphStyle(
                name="CustomHeader",
                parent=self.styles["Heading2"],
                fontSize=14,
                textColor=colors.HexColor("#27AE60"),
                spaceAfter=10,
                spaceBefore=10,
            )
        )

    def add_header(
        self, company_name="Lahimena Tours", address="Madagascar", logo_path=None
    ):
        """Add document header with company info"""
        title_block = [
            Paragraph(f"<b>{company_name}</b>", self.styles["CustomTitle"]),
            Paragraph(address, self.styles["CustomSubtitle"]),
            Paragraph("DEVIS / QUOTATION", self.styles["CustomSubtitle"]),
        ]

        if logo_path and os.path.exists(logo_path):
            try:
                logo = RLImage(logo_path, width=2.6 * cm, height=2.6 * cm)
                header_data = [[logo, title_block[0]], ["", title_block[1]], ["", title_block[2]]]
                table = Table(header_data, colWidths=[3.0 * cm, 16 * cm])
                table.setStyle(
                    TableStyle(
                        [
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("ALIGN", (0, 0), (0, 0), "LEFT"),
                        ]
                    )
                )
                self.elements.append(table)
            except Exception:
                header_data = [[title_block[0]], [title_block[1]], [title_block[2]]]
                self.elements.append(Table(header_data, colWidths=[19 * cm]))
        else:
            header_data = [[title_block[0]], [title_block[1]], [title_block[2]]]
            self.elements.append(Table(header_data, colWidths=[19 * cm]))
        self.elements.append(Spacer(1, 0.3 * inch))

        logger.debug("Header added to PDF")

    def add_quotation_info(self, quote_number, quote_date, client_name, client_email):
        """Add quotation metadata"""
        info_data = [
            ["Devis N°:", quote_number],
            ["Date:", quote_date],
            ["Client:", client_name],
            ["Email:", client_email],
        ]

        table = Table(info_data, colWidths=[4 * cm, 12 * cm])
        table.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 10),
                    ("FONT", (1, 0), (1, -1), "Helvetica", 10),
                    ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("BOTTOMMARGIN", (0, 0), (-1, -1), 5),
                ]
            )
        )

        self.elements.append(table)
        self.elements.append(Spacer(1, 0.3 * inch))

        logger.debug("Quotation info added to PDF")

    def add_client_contact(self, client_phone):
        """Add optional client phone line"""
        if not client_phone:
            return
        info_data = [["Téléphone:", client_phone]]
        table = Table(info_data, colWidths=[4 * cm, 12 * cm])
        table.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 10),
                    ("FONT", (1, 0), (1, -1), "Helvetica", 10),
                    ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("BOTTOMMARGIN", (0, 0), (-1, -1), 5),
                ]
            )
        )
        self.elements.append(table)
        self.elements.append(Spacer(1, 0.2 * inch))
        logger.debug("Client contact added to PDF")

    def add_section_title(self, title):
        """Add section title"""
        self.elements.append(Paragraph(title, self.styles["CustomHeader"]))
        self.elements.append(Spacer(1, 0.1 * inch))

    def add_quotation_details(
        self,
        nights,
        adults,
        children,
        room_type,
        price_per_night,
        total_price,
        currency="MGA",
    ):
        """Add quotation details table"""
        details_data = [
            ["Description", "Quantité", "Prix Unitaire", "Total"],
            [
                f"Hébergement - {room_type}",
                f"{nights} nuits",
                f"{price_per_night:,.0f} {currency}",
                f"{total_price:,.0f} {currency}",
            ],
        ]

        table = Table(details_data, colWidths=[8 * cm, 3 * cm, 4 * cm, 4 * cm])
        table.setStyle(
            TableStyle(
                [
                    # Header style
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#27AE60")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 11),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                    # Data rows
                    ("ALIGN", (0, 1), (-1, -1), "RIGHT"),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 1), (-1, -1), 10),
                    ("BOTTOMPADDING", (0, 1), (-1, -1), 8),
                    ("GRID", (0, 0), (-1, -1), 1, colors.grey),
                ]
            )
        )

        self.elements.append(table)
        self.elements.append(Spacer(1, 0.2 * inch))

        logger.debug("Quotation details added to PDF")

    def add_line_items_table(self, items, currency="MGA"):
        """Add a line items table for client quotations"""
        details_data = [["Description", "Quantité", "Prix Unitaire", "Total"]]

        for item in items:
            details_data.append(
                [
                    item.get("designation", ""),
                    f"{item.get('nights', 0)} nuits",
                    f"{item.get('unit_price', 0):,.2f} {currency}",
                    f"{item.get('total', 0):,.2f} {currency}",
                ]
            )

        table = Table(details_data, colWidths=[8 * cm, 3 * cm, 4 * cm, 4 * cm])
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#27AE60")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 11),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                    ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 1), (-1, -1), 10),
                    ("BOTTOMPADDING", (0, 1), (-1, -1), 8),
                    ("GRID", (0, 0), (-1, -1), 1, colors.grey),
                ]
            )
        )

        self.elements.append(table)
        self.elements.append(Spacer(1, 0.2 * inch))

        logger.debug("Line items table added to PDF")

    def add_totals_table(self, subtotal, tax=0, total=None, currency="MGA"):
        """Add totals summary table"""
        if total is None:
            total = subtotal + tax

        totals_data = [
            ["", ""],
            ["Sous-Total:", f"{subtotal:,.0f} {currency}"],
        ]

        if tax > 0:
            totals_data.append(["Taxes/Frais:", f"{tax:,.0f} {currency}"])

        totals_data.append(["TOTAL:", f"{total:,.0f} {currency}"])

        table = Table(totals_data, colWidths=[13 * cm, 5 * cm])
        table.setStyle(
            TableStyle(
                [
                    ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                    ("FONTNAME", (0, -1), (0, -1), "Helvetica-Bold"),
                    ("FONTNAME", (1, -1), (1, -1), "Helvetica-Bold"),
                    ("FONTSIZE", (1, -1), (1, -1), 12),
                    ("TEXTCOLOR", (0, -1), (1, -1), colors.HexColor("#27AE60")),
                    ("LINEABOVE", (0, -1), (-1, -1), 2, colors.HexColor("#27AE60")),
                    ("BOTTOMPADDING", (0, -1), (-1, -1), 12),
                ]
            )
        )

        self.elements.append(table)
        self.elements.append(Spacer(1, 0.3 * inch))

        logger.debug("Totals table added to PDF")

    def add_totals_table_with_breakdown(
        self, subtotal, margin_amount, tva_amount, total, currency="MGA"
    ):
        """Add totals table with margin and TVA breakdown"""
        totals_data = [
            ["", ""],
            ["Sous-Total:", f"{subtotal:,.2f} {currency}"],
        ]

        if margin_amount:
            totals_data.append(["Marge:", f"{margin_amount:,.2f} {currency}"])

        if tva_amount:
            totals_data.append(["TVA:", f"{tva_amount:,.2f} {currency}"])

        totals_data.append(["TOTAL:", f"{total:,.2f} {currency}"])

        table = Table(totals_data, colWidths=[13 * cm, 5 * cm])
        table.setStyle(
            TableStyle(
                [
                    ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                    ("FONTNAME", (0, -1), (0, -1), "Helvetica-Bold"),
                    ("FONTNAME", (1, -1), (1, -1), "Helvetica-Bold"),
                    ("FONTSIZE", (1, -1), (1, -1), 12),
                    ("TEXTCOLOR", (0, -1), (1, -1), colors.HexColor("#27AE60")),
                    ("LINEABOVE", (0, -1), (-1, -1), 2, colors.HexColor("#27AE60")),
                    ("BOTTOMPADDING", (0, -1), (-1, -1), 12),
                ]
            )
        )

        self.elements.append(table)
        self.elements.append(Spacer(1, 0.3 * inch))

        logger.debug("Totals breakdown table added to PDF")

    def add_terms(self, terms_text=""):
        """Add terms and conditions section"""
        default_terms = (
            "Conditions: Tarifs sujets à modification. Validité du devis: 30 jours."
        )

        text = terms_text or default_terms
        self.elements.append(Paragraph("<b>Conditions:</b>", self.styles["Normal"]))
        self.elements.append(Paragraph(text, self.styles["Normal"]))
        self.elements.append(Spacer(1, 0.2 * inch))

        logger.debug("Terms added to PDF")

    def add_footer(self, footer_text=""):
        """Add footer with contact information"""
        default_footer = "Lahimena Tours | Madagascar | info@lahimena.com"

        text = footer_text or default_footer
        self.elements.append(Spacer(1, 0.3 * inch))
        footer = Paragraph(
            f"<i>{text}</i>",
            ParagraphStyle(
                "Footer",
                parent=self.styles["Normal"],
                fontSize=9,
                textColor=colors.grey,
                alignment=TA_CENTER,
            ),
        )
        self.elements.append(footer)

    def generate(self):
        """Generate and save the PDF"""
        try:
            self.doc.build(self.elements)
            logger.info(f"PDF generated successfully: {self.filepath}")
            return self.filepath
        except Exception as e:
            logger.error(f"Error generating PDF: {e}", exc_info=True)
            raise


def generate_hotel_quotation_pdf(
    client_name,
    client_email,
    hotel_name,
    nights,
    adults,
    room_type,
    price_per_night,
    total_price,
    currency="MGA",
    quote_number=None,
    quote_date=None,
    hotel_location=None,
    hotel_category=None,
    hotel_contact=None,
    hotel_email=None,
    output_dir="devis",
):
    """
    Generate a hotel quotation PDF

    Args:
        client_name: Name of the client
        client_email: Email of the client
        hotel_name: Name of the hotel
        nights: Number of nights
        adults: Number of adults
        room_type: Type of room
        price_per_night: Price per night
        total_price: Total price
        currency: Currency code (default: MGA)
        quote_number: Quote number (optional, auto-generated if not provided)
        quote_date: Quote date (optional, uses today if not provided)
        hotel_location: Hotel location (optional)
        hotel_category: Hotel category (optional)
        hotel_contact: Hotel contact (optional)
        hotel_email: Hotel email (optional)
        output_dir: Output directory for PDF (default: devis)

    Returns:
        str: Path to generated PDF file
    """
    if not REPORTLAB_AVAILABLE:
        logger.error("ReportLab not available")
        raise ImportError("ReportLab required for PDF generation")

    try:
        # Create PDF document
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if not quote_number:
            quote_number = f"DEVIS_HOTEL_{timestamp}"
        if not quote_date:
            quote_date = datetime.now().strftime("%d/%m/%Y")

        # Ensure output directory exists
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename = os.path.join(output_dir, f"{quote_number}.pdf")

        pdf = QuotationPDF(filename)

        # Build PDF content
        pdf.add_header(
            COMPANY_NAME, COMPANY_TAGLINE, logo_path=LOGO_PATH
        )

        pdf.add_quotation_info(quote_number, quote_date, client_name, client_email)

        # Add hotel information if provided
        hotel_info = f"Séjour à {hotel_name}"
        if hotel_location:
            hotel_info += f" - {hotel_location}"
        pdf.add_section_title(hotel_info)

        pdf.add_quotation_details(
            nights=nights,
            adults=adults,
            children=0,
            room_type=room_type,
            price_per_night=price_per_night,
            total_price=total_price,
            currency=currency,
        )

        # Add totals
        pdf.add_totals_table(
            subtotal=total_price, tax=0, total=total_price, currency=currency
        )

        # Add hotel details if provided
        details_text = "Informations Hôtel:\n"
        if hotel_category:
            details_text += f"• Catégorie: {hotel_category}\n"
        if hotel_contact:
            details_text += f"• Contact: {hotel_contact}\n"
        if hotel_email:
            details_text += f"• Email: {hotel_email}\n"
        details_text += "\nTarif valide 30 jours"

        pdf.add_terms(details_text)
        pdf.add_footer(PDF_FOOTER_TEXT)

        # Generate PDF
        filepath = pdf.generate()

        logger.info(f"Hotel quotation PDF created: {filepath}")
        return filepath

    except Exception as e:
        logger.error(f"Error generating hotel quotation PDF: {e}", exc_info=True)
        raise


def generate_multi_hotel_quotation_pdf(
    client_name,
    client_email,
    client_phone,
    quote_number,
    quote_date,
    items,
    currency,
    subtotal,
    total,
    output_dir="devis",
):
    """
    Generate a multi-hotel quotation PDF (one line per hotel).
    """
    if not REPORTLAB_AVAILABLE:
        logger.error("ReportLab not available")
        raise ImportError("ReportLab required for PDF generation")

    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename = os.path.join(output_dir, f"{quote_number}.pdf")
        pdf = QuotationPDF(filename)

        pdf.add_header(
            COMPANY_NAME, COMPANY_TAGLINE, logo_path=LOGO_PATH
        )
        pdf.add_quotation_info(quote_number, quote_date, client_name, client_email)
        pdf.add_client_contact(client_phone)

        pdf.add_section_title("Détails des hôtels")
        pdf.add_line_items_table(items, currency=currency)

        pdf.add_totals_table(subtotal=subtotal, tax=0, total=total, currency=currency)

        pdf.add_terms(
            "Conditions: Tarifs sujets à modification. Validité du devis: 30 jours."
        )
        pdf.add_footer(PDF_FOOTER_TEXT)

        filepath = pdf.generate()
        logger.info(f"Multi-hotel quotation PDF created: {filepath}")
        return filepath

    except Exception as e:
        logger.error(f"Error generating multi-hotel quotation PDF: {e}", exc_info=True)
        raise


def generate_client_quotation_pdf(
    client_name,
    client_email,
    client_phone,
    quote_number,
    quote_date,
    items,
    currency,
    margin_pct,
    margin_amount,
    tva_pct,
    tva_amount,
    subtotal,
    total,
    output_dir="devis",
):
    """
    Generate a client quotation PDF with line items and margin/TVA breakdown.
    """
    if not REPORTLAB_AVAILABLE:
        logger.error("ReportLab not available")
        raise ImportError("ReportLab required for PDF generation")

    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename = os.path.join(output_dir, f"{quote_number}.pdf")
        pdf = QuotationPDF(filename)

        pdf.add_header(
            COMPANY_NAME, COMPANY_TAGLINE, logo_path=LOGO_PATH
        )
        pdf.add_quotation_info(quote_number, quote_date, client_name, client_email)
        pdf.add_client_contact(client_phone)

        pdf.add_section_title("Détails du devis")
        pdf.add_line_items_table(items, currency=currency)
        pdf.add_totals_table_with_breakdown(
            subtotal=subtotal,
            margin_amount=margin_amount,
            tva_amount=tva_amount,
            total=total,
            currency=currency,
        )

        pdf.add_terms(
            "Conditions: Tarifs sujets à modification. Validité du devis: 30 jours.\n"
            f"Marge appliquée: {margin_pct:,.2f}% | TVA appliquée: {tva_pct:,.2f}%."
        )
        pdf.add_footer(PDF_FOOTER_TEXT)

        filepath = pdf.generate()
        logger.info(f"Client quotation PDF created: {filepath}")
        return filepath

    except Exception as e:
        logger.error(f"Error generating client quotation PDF: {e}", exc_info=True)
        raise


def generate_air_ticket_cotation_pdf(client: dict, rows: list, output_dir: str = None) -> str:
    """
    Génère un devis PDF de cotation avion client.

    Args:
        client: dict avec au moins nom, prenom, ref_client, numero_dossier.
        rows:   liste de dicts cotation avion (_make_row format).
        output_dir: dossier de sortie (défaut: DEVIS_FOLDER).

    Returns:
        Chemin absolu du fichier PDF généré.
    """
    if not REPORTLAB_AVAILABLE:
        raise ImportError("ReportLab requis pour la génération PDF.")

    import re

    out_dir = output_dir or DEVIS_FOLDER
    os.makedirs(out_dir, exist_ok=True)

    nom    = str(client.get("nom", "") or "").strip()
    prenom = str(client.get("prenom", "") or "").strip()
    dossier = str(client.get("numero_dossier", "") or "").strip()
    ref     = str(client.get("ref_client", "") or "").strip()
    safe    = re.sub(r"[^\w\-]", "_", f"{ref}_{nom}_{prenom}")
    filename = os.path.join(out_dir, f"devis_avion_{safe}.pdf")

    try:
        pdf = QuotationPDF(filename)

        pdf.add_header(COMPANY_NAME, COMPANY_TAGLINE, logo_path=LOGO_PATH)
        pdf.add_quotation_info(
            quote_number=dossier or ref,
            quote_date=datetime.now().strftime("%d/%m/%Y"),
            client_name=f"{nom} {prenom}".strip(),
            client_email=str(client.get("email", "") or ""),
        )

        pdf.add_section_title("Cotation billets d'avion")

        # Tableau des vols
        col_headers = [
            "Date vol", "N° vol", "Trajet", "Compagnie", "Classe",
            "Adultes", "Enfants", "Tarif adulte", "Tarif enfant",
            "Sous-total", "Marge %", "Total (Ar)",
        ]

        def _f(v):
            try:
                return f"{float(v):,.0f}"
            except Exception:
                return str(v or "")

        data_rows = [col_headers]
        for r in rows:
            trajet = f"{r.get('ville_depart', '')} → {r.get('ville_arrivee', '')}"
            data_rows.append([
                r.get("date_vol", ""),
                r.get("numero_vol", ""),
                f"{r.get('type_trajet', '')} — {trajet}",
                r.get("compagnie", ""),
                r.get("classe", "Éco"),
                r.get("nb_adultes", ""),
                r.get("nb_enfants", ""),
                _f(r.get("tarif_adulte", 0)),
                _f(r.get("tarif_enfant", 0)),
                _f(r.get("sous_total", 0)),
                f"{r.get('marge_pct', '0')} %",
                _f(r.get("total", 0)),
            ])

        grand_total = sum(float(r.get("total") or 0) for r in rows)
        total_adultes = sum(float(r.get("montant_adultes") or 0) for r in rows)
        total_enfants = sum(float(r.get("montant_enfants") or 0) for r in rows)

        # Ligne totaux
        data_rows.append([
            "", "", "", "", "TOTAUX",
            "", "", _f(total_adultes), _f(total_enfants),
            "", "", _f(grand_total),
        ])

        col_widths = [55, 50, 130, 75, 55, 40, 40, 65, 65, 65, 45, 70]

        table = Table(data_rows, colWidths=[w * 0.75 for w in col_widths])
        table.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (-1, 0),  colors.HexColor("#0F7D8A")),
            ("TEXTCOLOR",   (0, 0), (-1, 0),  colors.white),
            ("FONTNAME",    (0, 0), (-1, 0),  "Helvetica-Bold"),
            ("FONTSIZE",    (0, 0), (-1, 0),  7),
            ("FONTSIZE",    (0, 1), (-1, -1), 7),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#EAF4F7")]),
            ("BACKGROUND",  (0, -1), (-1, -1), colors.HexColor("#DCEFF3")),
            ("FONTNAME",    (0, -1), (-1, -1), "Helvetica-Bold"),
            ("GRID",        (0, 0), (-1, -1),  0.3, colors.HexColor("#AABBCC")),
            ("ALIGN",       (5, 1), (-1, -1),  "RIGHT"),
            ("VALIGN",      (0, 0), (-1, -1),  "MIDDLE"),
            ("TOPPADDING",  (0, 0), (-1, -1),  3),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
        ]))

        from reportlab.platypus import Spacer
        pdf.elements.append(table)
        pdf.elements.append(Spacer(1, 0.4 * cm))

        # Récapitulatif
        pdf.add_totals_table_with_breakdown(
            subtotal=grand_total,
            margin_amount=0,
            tva_amount=0,
            total=grand_total,
            currency="Ar",
        )

        pdf.add_terms(
            "Tarifs indicatifs — sujets à modification selon disponibilités.\n"
            "Validité du devis : 7 jours."
        )
        pdf.add_footer(PDF_FOOTER_TEXT)

        filepath = pdf.generate()
        logger.info(f"Air ticket cotation PDF created: {filepath}")
        return filepath

    except Exception as exc:
        logger.error(f"Error generating air ticket cotation PDF: {exc}", exc_info=True)
        raise


def generate_invoice_pdf(
    invoice_number,
    invoice_date,
    client_name,
    client_email,
    client_phone,
    source_type,
    source_ref,
    currency,
    montant_ht,
    marge_pct,
    marge_amount,
    tva_pct,
    tva_amount,
    total_ttc,
    acompte,
    reste_a_payer,
    statut,
    items=None,
    base_taxable_ht=None,
    output_dir="devis",
):
    """Generate an invoice PDF from one invoice record."""
    if not REPORTLAB_AVAILABLE:
        logger.error("ReportLab not available")
        raise ImportError("ReportLab required for PDF generation")

    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename = os.path.join(output_dir, f"{invoice_number}.pdf")
        pdf = QuotationPDF(filename)

        pdf.add_header(
            COMPANY_NAME, COMPANY_TAGLINE, logo_path=LOGO_PATH
        )
        pdf.add_quotation_info(
            quote_number=invoice_number,
            quote_date=invoice_date,
            client_name=client_name,
            client_email=client_email,
        )
        pdf.add_client_contact(client_phone)

        pdf.add_section_title("Détails de la facture")
        raw_items = items or [
            {
                "designation": f"{source_type} - {source_ref}",
                "nights": 1,
                "unit_price": montant_ht,
                "total": montant_ht,
            }
        ]

        # Integrate margin/taxes directly into displayed unit prices.
        base_sum = sum(float(item.get("total", 0) or 0) for item in raw_items)
        final_sum = float(total_ttc or 0)
        coef = (final_sum / base_sum) if base_sum > 0 else 1.0
        line_items = []
        for item in raw_items:
            qty = max(1, int(item.get("nights", 1) or 1))
            base_total = float(item.get("total", 0) or 0)
            final_total = base_total * coef
            unit_price = final_total / qty if qty else final_total
            line_items.append(
                {
                    "designation": item.get("designation", ""),
                    "nights": qty,
                    "unit_price": unit_price,
                    "total": final_total,
                }
            )

        pdf.add_line_items_table(line_items, currency=currency)

        subtotal_display = sum(float(item.get("total", 0) or 0) for item in line_items)
        pdf.add_totals_table(
            subtotal=subtotal_display,
            tax=0,
            total=subtotal_display,
            currency=currency,
        )

        pdf.add_terms(
            "Informations de paiement:\n"
            f"Statut: {statut}\n"
            f"Acompte reçu: {acompte:,.2f} {currency}\n"
            f"Reste à payer: {reste_a_payer:,.2f} {currency}\n"
            "Prix unitaires affichés: frais et taxes inclus."
        )
        pdf.add_footer(PDF_FOOTER_TEXT)

        filepath = pdf.generate()
        logger.info(f"Invoice PDF created: {filepath}")
        return filepath
    except Exception as e:
        logger.error(f"Error generating invoice PDF: {e}", exc_info=True)
        raise
