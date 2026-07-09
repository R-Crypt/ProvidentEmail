import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

def generate_pdf(filename="Provident_Classifier_Maintenance_Proposal.pdf"):
    # Target path setup
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )
    
    story = []
    
    # Styles
    styles = getSampleStyleSheet()
    
    # Color palette
    PRIMARY_COLOR = colors.HexColor("#0f172a") # Dark Slate / Navy
    TEXT_COLOR = colors.HexColor("#334155") # Muted charcoal body text
    BORDER_COLOR = colors.HexColor("#e2e8f0")
    
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontSize=18,
        leading=22,
        textColor=PRIMARY_COLOR,
        spaceAfter=20
    )
    
    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontSize=10,
        leading=12,
        fontName='Helvetica-Bold',
        textColor=colors.white
    )
    
    table_cell_style = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontSize=9.5,
        leading=13,
        textColor=TEXT_COLOR
    )

    table_cell_bold = ParagraphStyle(
        'TableCellBold',
        parent=styles['Normal'],
        fontSize=10,
        leading=13,
        fontName='Helvetica-Bold',
        textColor=PRIMARY_COLOR
    )

    # Document Header (Title and Table only)
    story.append(Paragraph("Provident Operations Copilot — Monthly Maintenance Costs", title_style))
    
    # Table data
    data = [
        [
            Paragraph("Service & Details", table_header_style),
            Paragraph("Billing Cycle", table_header_style),
            Paragraph("Amount", table_header_style)
        ],
        [
            Paragraph("<b>Cloud Server Hosting Subscription</b><br/>Monthly hosting plan to run the software core 24/7 in the cloud so it is always active inside Outlook.", table_cell_style),
            Paragraph("Monthly", table_cell_style),
            Paragraph("$24.50", table_cell_style)
        ],
        [
            Paragraph("<b>Managed Database Service Subscription</b><br/>Storage instance subscription to save email triage records, including automated daily backup snapshot streams.", table_cell_style),
            Paragraph("Monthly", table_cell_style),
            Paragraph("$22.50", table_cell_style)
        ],
        [
            Paragraph("<b>Error Tracking & Diagnostics Platform License</b><br/>Software diagnostics subscription that automatically logs errors and system crashes for troubleshooting.", table_cell_style),
            Paragraph("Monthly", table_cell_style),
            Paragraph("$20.00", table_cell_style)
        ],
        [
            Paragraph("<b>OpenAI AI Engine Token Subscription</b><br/>Access credit pool subscription for the artificial intelligence engine to read email text and generate draft responses.", table_cell_style),
            Paragraph("Monthly", table_cell_style),
            Paragraph("$18.50", table_cell_style)
        ],
        [
            Paragraph("<b>Domain & SSL Security Certificate Upkeep</b><br/>Subscription to keep the domain active and secure HTTPS traffic certification required by Microsoft's security policy.", table_cell_style),
            Paragraph("Monthly", table_cell_style),
            Paragraph("$14.50", table_cell_style)
        ],
        [
            Paragraph("<b>TOTAL MONTHLY MAINTENANCE</b>", table_cell_bold),
            Paragraph("—", table_cell_bold),
            Paragraph("<b>$100.00</b>", table_cell_bold)
        ]
    ]
    
    # 532 pt total width for margins
    col_widths = [350, 102, 80]
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), PRIMARY_COLOR),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('LINEBELOW', (0,0), (-1,-2), 0.5, BORDER_COLOR),
        ('LINEABOVE', (0,-1), (-1,-1), 1.5, PRIMARY_COLOR),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#f8fafc")),
    ]))
    
    story.append(t)
    
    # Build Document
    doc.build(story)
    print(f"Success: Generated {filename}")

if __name__ == "__main__":
    generate_pdf()
