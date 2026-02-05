"""
PDF generation utilities for quotations
Uses ReportLab to create professional PDF documents
"""

try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

import os
from datetime import datetime
from config import DEVIS_FOLDER
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
            raise ImportError("ReportLab is required for PDF generation. Install with: pip install reportlab")
        
        # Generate filename if not provided
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"DEVIS_{timestamp}.pdf"
        
        # Ensure devis folder exists
        os.makedirs(DEVIS_FOLDER, exist_ok=True)
        
        self.filepath = os.path.join(DEVIS_FOLDER, filename)
        self.doc = SimpleDocTemplate(
            self.filepath,
            pagesize=A4,
            rightMargin=2*cm,
            leftMargin=2*cm,
            topMargin=2*cm,
            bottomMargin=2*cm
        )
        
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
        self.elements = []
        
        logger.info(f"PDF generator initialized for: {filename}")
    
    def _setup_custom_styles(self):
        """Setup custom paragraph styles"""
        # Title style
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#1E3A5F'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))
        
        # Subtitle style
        self.styles.add(ParagraphStyle(
            name='CustomSubtitle',
            parent=self.styles['Normal'],
            fontSize=11,
            textColor=colors.HexColor('#666666'),
            spaceAfter=12,
            alignment=TA_CENTER
        ))
        
        # Header style
        self.styles.add(ParagraphStyle(
            name='CustomHeader',
            parent=self.styles['Heading2'],
            fontSize=14,
            textColor=colors.HexColor('#27AE60'),
            spaceAfter=10,
            spaceBefore=10
        ))
    
    def add_header(self, company_name="Lahimena Tours", address="Madagascar"):
        """Add document header with company info"""
        header_data = [
            [Paragraph(f"<b>{company_name}</b>", self.styles['CustomTitle'])],
            [Paragraph(address, self.styles['CustomSubtitle'])],
            [Paragraph("DEVIS / QUOTATION", self.styles['CustomSubtitle'])],
        ]
        
        self.elements.append(Table(header_data, colWidths=[19*cm]))
        self.elements.append(Spacer(1, 0.3*inch))
        
        logger.debug("Header added to PDF")
    
    def add_quotation_info(self, quote_number, quote_date, client_name, client_email):
        """Add quotation metadata"""
        info_data = [
            ["Devis N°:", quote_number],
            ["Date:", quote_date],
            ["Client:", client_name],
            ["Email:", client_email],
        ]
        
        table = Table(info_data, colWidths=[4*cm, 12*cm])
        table.setStyle(TableStyle([
            ('FONT', (0, 0), (0, -1), 'Helvetica-Bold', 10),
            ('FONT', (1, 0), (1, -1), 'Helvetica', 10),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMMARGIN', (0, 0), (-1, -1), 5),
        ]))
        
        self.elements.append(table)
        self.elements.append(Spacer(1, 0.3*inch))
        
        logger.debug("Quotation info added to PDF")
    
    def add_section_title(self, title):
        """Add section title"""
        self.elements.append(Paragraph(title, self.styles['CustomHeader']))
        self.elements.append(Spacer(1, 0.1*inch))
    
    def add_quotation_details(self, nights, adults, children, room_type, price_per_night, total_price, currency="MGA"):
        """Add quotation details table"""
        details_data = [
            ["Description", "Quantité", "Prix Unitaire", "Total"],
            [
                f"Hébergement - {room_type}",
                f"{nights} nuits",
                f"{price_per_night:,.0f} {currency}",
                f"{total_price:,.0f} {currency}"
            ],
        ]
        
        table = Table(details_data, colWidths=[8*cm, 3*cm, 4*cm, 4*cm])
        table.setStyle(TableStyle([
            # Header style
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#27AE60')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            
            # Data rows
            ('ALIGN', (0, 1), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ]))
        
        self.elements.append(table)
        self.elements.append(Spacer(1, 0.2*inch))
        
        logger.debug("Quotation details added to PDF")
    
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
        
        table = Table(totals_data, colWidths=[13*cm, 5*cm])
        table.setStyle(TableStyle([
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('FONTNAME', (0, -1), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, -1), (1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (1, -1), (1, -1), 12),
            ('TEXTCOLOR', (0, -1), (1, -1), colors.HexColor('#27AE60')),
            ('LINEABOVE', (0, -1), (-1, -1), 2, colors.HexColor('#27AE60')),
            ('BOTTOMPADDING', (0, -1), (-1, -1), 12),
        ]))
        
        self.elements.append(table)
        self.elements.append(Spacer(1, 0.3*inch))
        
        logger.debug("Totals table added to PDF")
    
    def add_terms(self, terms_text=""):
        """Add terms and conditions section"""
        default_terms = "Conditions: Tarifs sujets à modification. Validité du devis: 30 jours."
        
        text = terms_text or default_terms
        self.elements.append(Paragraph("<b>Conditions:</b>", self.styles['Normal']))
        self.elements.append(Paragraph(text, self.styles['Normal']))
        self.elements.append(Spacer(1, 0.2*inch))
        
        logger.debug("Terms added to PDF")
    
    def add_footer(self, footer_text=""):
        """Add footer with contact information"""
        default_footer = "Lahimena Tours | Madagascar | info@lahimena.com"
        
        text = footer_text or default_footer
        self.elements.append(Spacer(1, 0.3*inch))
        footer = Paragraph(
            f"<i>{text}</i>",
            ParagraphStyle(
                'Footer',
                parent=self.styles['Normal'],
                fontSize=9,
                textColor=colors.grey,
                alignment=TA_CENTER
            )
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


def generate_hotel_quotation_pdf(client_name, client_email, hotel_name, nights, adults, 
                                 room_type, price_per_night, total_price, currency="MGA",
                                 quote_number=None, quote_date=None, hotel_location=None,
                                 hotel_category=None, hotel_contact=None, hotel_email=None,
                                 output_dir="devis"):
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
        pdf.add_header("Lahimena Tours", "Madagascar - Tours & Travel")
        
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
            currency=currency
        )
        
        # Add totals
        pdf.add_totals_table(
            subtotal=total_price,
            tax=0,
            total=total_price,
            currency=currency
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
        pdf.add_footer("Lahimena Tours | Madagascar | Tel: +261-32-XXXX-XXXX")
        
        # Generate PDF
        filepath = pdf.generate()
        
        logger.info(f"Hotel quotation PDF created: {filepath}")
        return filepath
        
    except Exception as e:
        logger.error(f"Error generating hotel quotation PDF: {e}", exc_info=True)
        raise
