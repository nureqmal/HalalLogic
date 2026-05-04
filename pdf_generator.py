from reportlab.platypus import *
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors

def generate_pdf(company, policy, materials, sop, trace, recall, evaluation):

    filename = "IHCS_Manual_Professional.pdf"

    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=50,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'title',
        parent=styles['Title'],
        alignment=TA_CENTER
    )

    heading = styles['Heading2']
    normal = styles['Normal']

    elements = []

    # =========================
    # COVER PAGE
    # =========================
    elements.append(Spacer(1, 200))
    elements.append(Paragraph("INTERNAL HALAL CONTROL SYSTEM", title_style))
    elements.append(Paragraph("(IHCS) MANUAL", title_style))
    elements.append(Spacer(1, 30))

    elements.append(Paragraph(f"<b>{company[0]}</b>", title_style))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph(company[1], title_style))

    elements.append(PageBreak())

    # =========================
    # DOCUMENT CONTROL
    # =========================
    elements.append(Paragraph("DOCUMENT CONTROL", heading))

    doc_table = Table([
        ["Document Title", "IHCS Manual"],
        ["Company", company[0]],
        ["Prepared By", "Halal Committee"],
        ["Approved By", "Management"],
        ["Version", "1.0"],
    ])

    doc_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
    ]))

    elements.append(doc_table)
    elements.append(Spacer(1, 20))

    # =========================
    # 1.0 COMPANY PROFILE
    # =========================
    elements.append(Paragraph("1.0 COMPANY PROFILE", heading))
    elements.append(Paragraph(
        f"The company, {company[0]}, operates at {company[1]}. "
        f"All operations are conducted in compliance with halal requirements.",
        normal
    ))

    elements.append(Spacer(1, 15))

    # =========================
    # 2.0 HALAL POLICY
    # =========================
    elements.append(Paragraph("2.0 HALAL POLICY", heading))
    elements.append(Paragraph(
        "The company is committed to ensuring that all products comply with halal standards. "
        "The following policy is implemented:",
        normal
    ))
    elements.append(Paragraph(policy, normal))

    elements.append(Spacer(1, 15))

    # =========================
    # 3.0 RAW MATERIAL
    # =========================
    elements.append(Paragraph("3.0 RAW MATERIAL CONTROL", heading))

    table_data = [["Material", "Supplier", "Halal Cert", "Expiry"]]

    for m in materials:
        table_data.append([m[0], m[1], m[2], m[3]])

    table = Table(table_data)
    table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 15))

    # =========================
    # 4.0 SOP
    # =========================
    elements.append(Paragraph("4.0 STANDARD OPERATING PROCEDURES", heading))

    elements.append(Paragraph(
        "<b>4.1 Purchasing</b><br/>All materials shall be purchased from approved halal-certified suppliers. "
        + sop[0],
        normal
    ))

    elements.append(Paragraph(
        "<b>4.2 Receiving</b><br/>All incoming materials shall be verified upon arrival. "
        + sop[1],
        normal
    ))

    elements.append(Paragraph(
        "<b>4.3 Storage</b><br/>Materials shall be stored in a clean and segregated area. "
        + sop[2],
        normal
    ))

    elements.append(Spacer(1, 15))

    # =========================
    # 5.0 TRACEABILITY
    # =========================
    elements.append(Paragraph("5.0 TRACEABILITY SYSTEM", heading))
    elements.append(Paragraph(
        "The company shall establish a traceability system to track materials and finished products.",
        normal
    ))
    elements.append(Paragraph(trace, normal))

    elements.append(Spacer(1, 15))

    # =========================
    # 6.0 RECALL
    # =========================
    elements.append(Paragraph("6.0 PRODUCT RECALL PROCEDURE", heading))
    elements.append(Paragraph(
        "The company shall implement a product recall system to ensure non-compliant products are removed.",
        normal
    ))
    elements.append(Paragraph(recall, normal))

    elements.append(Spacer(1, 15))

    # =========================
    # 7.0 EVALUATION
    # =========================
    elements.append(Paragraph("7.0 IHCS EVALUATION & REVIEW", heading))
    elements.append(Paragraph(
        "The IHCS system shall be reviewed periodically to ensure effectiveness.",
        normal
    ))
    elements.append(Paragraph(evaluation, normal))

    doc.build(elements)

    return filename
