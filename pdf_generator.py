from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

def generate_pdf(company, policy, materials, filename="IHCS_Manual.pdf"):

    doc = SimpleDocTemplate(filename, pagesize=A4)
    styles = getSampleStyleSheet()

    elements = []

    # TITLE
    elements.append(Paragraph("IHCS MANUAL", styles['Title']))
    elements.append(Spacer(1, 20))

    # COMPANY
    elements.append(Paragraph("1. Company Profile", styles['Heading2']))
    elements.append(Paragraph(f"Name: {company[0]}", styles['Normal']))
    elements.append(Paragraph(f"Address: {company[1]}", styles['Normal']))
    elements.append(Paragraph(f"Contact: {company[2]}", styles['Normal']))
    elements.append(Spacer(1, 15))

    # POLICY
    elements.append(Paragraph("2. Halal Policy", styles['Heading2']))
    elements.append(Paragraph(policy, styles['Normal']))
    elements.append(Spacer(1, 15))

    # MATERIAL TABLE
    elements.append(Paragraph("3. Raw Material Masterlist", styles['Heading2']))

    table_data = [["Material", "Supplier", "Halal Cert", "Expiry"]]

    for m in materials:
        table_data.append([m[0], m[1], m[2], m[3]])

    table = Table(table_data)

    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
    ]))

    elements.append(table)

    doc.build(elements)

    return filename
