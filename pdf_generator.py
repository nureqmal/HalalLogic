from reportlab.lib.pagesizes import A4
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

def generate_pdf(company, policy, materials, sop, trace, recall, evaluation, filename="IHCS_Manual.pdf"):

    doc = SimpleDocTemplate(filename, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    # TITLE
    elements.append(Paragraph("INTERNAL HALAL CONTROL SYSTEM (IHCS) MANUAL", styles['Title']))
    elements.append(Spacer(1, 20))

    # 1 COMPANY
    elements.append(Paragraph("1. Company Profile", styles['Heading2']))
    elements.append(Paragraph(f"Name: {company[0]}", styles['Normal']))
    elements.append(Paragraph(f"Address: {company[1]}", styles['Normal']))
    elements.append(Paragraph(f"Contact: {company[2]}", styles['Normal']))
    elements.append(Spacer(1, 15))

    # 2 POLICY
    elements.append(Paragraph("2. Halal Policy", styles['Heading2']))
    elements.append(Paragraph(policy, styles['Normal']))
    elements.append(Spacer(1, 15))

    # 3 RAW MATERIAL
    elements.append(Paragraph("3. Raw Material Masterlist", styles['Heading2']))
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

    # 4 SOP
    elements.append(Paragraph("4. Standard Operating Procedures (SOP)", styles['Heading2']))
    elements.append(Paragraph(f"Purchasing: {sop[0]}", styles['Normal']))
    elements.append(Paragraph(f"Receiving: {sop[1]}", styles['Normal']))
    elements.append(Paragraph(f"Storage: {sop[2]}", styles['Normal']))
    elements.append(Spacer(1, 15))

    # 5 TRACEABILITY
    elements.append(Paragraph("5. Traceability Program", styles['Heading2']))
    elements.append(Paragraph(trace, styles['Normal']))
    elements.append(Spacer(1, 15))

    # 6 RECALL
    elements.append(Paragraph("6. Product Recall Procedure", styles['Heading2']))
    elements.append(Paragraph(recall, styles['Normal']))
    elements.append(Spacer(1, 15))

    # 7 EVALUATION
    elements.append(Paragraph("7. IHCS Evaluation & Review", styles['Heading2']))
    elements.append(Paragraph(evaluation, styles['Normal']))

    doc.build(elements)
    return filename
